<#
.SYNOPSIS
    Aplica atributos de diretorio (Entra ID) e relacao de manager nos usuarios TRINCA
    a partir da planilha de atributos. Esses atributos alimentam as regras dos
    grupos DINAMICOS (DG-TRINCA-*), que populam automaticamente.

.DESCRIPTION
    Colunas esperadas:
      userPrincipalName, mail, accountEnabled, givenName, surname, displayName,
      jobTitle, manager, department, extensionAttribute1, companyName,
      officeLocation, mobilePhone

    Caracteristicas de producao:
      * Cache unico de usuarios (1 Get-MgUser -All) com os atributos atuais -> diff real.
      * Idempotente: so faz PATCH dos campos que mudaram; reporta UPDATED/UNCHANGED.
      * Dry-run por padrao. So aplica com -Execute.
      * Fase 1 = atributos; Fase 2 = manager (resolve UPN; pula nao-resolviveis).
      * Log estruturado + relatorio CSV de auditoria.
      * Erros do Graph viram status reais no CSV (ERROR-SYNC/ERROR-PRIV/ERROR-MIGRATING).

    Mapeamento que alimenta os DGs:
      department          -> DG-TRINCA-Area-*   (cuidado: a rule deve usar a string real)
      employeeType        -> DG-TRINCA-Dir-*    (CEO/COO/CCO/CIO/CDO)
      companyName=TRINCA  -> DG-TRINCA-Users

.PARAMETER ExcelPath
    Caminho do .xlsx de atributos.

.PARAMETER Execute
    Aplica de fato. Sem o switch, roda em simulacao (dry-run).

.PARAMETER SkipManager
    Nao processa a Fase 2 (manager).

.PARAMETER ReportPath
    CSV de relatorio. Default: .\AttrReport_<timestamp>.csv

.EXAMPLE
    .\Set-TrincaUserAttributes.ps1 -ExcelPath .\users_final_para_seven-91.xlsx
    .\Set-TrincaUserAttributes.ps1 -ExcelPath .\users_final_para_seven-91.xlsx -Execute

.NOTES
    Requer Microsoft.Graph + ImportExcel. Scope: User.ReadWrite.All
    A diretoria vem da coluna extensionAttribute1 da planilha e e gravada em employeeType/employeeId.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ExcelPath,
    [switch]$Execute,
    [switch]$SkipManager,
    [string]$ReportPath = ".\AttrReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

#region ---------- Bootstrap ----------
function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "[setup] Instalando $Name..." -ForegroundColor Yellow
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Name -ErrorAction Stop
}
Ensure-Module ImportExcel
Ensure-Module Microsoft.Graph.Authentication
Ensure-Module Microsoft.Graph.Users
#endregion

#region ---------- Conexao ----------
if (-not (Get-MgContext)) {
    Write-Host "[auth] Conectando ao Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes 'User.ReadWrite.All' -NoWelcome
}
$ctx = Get-MgContext
Write-Host "[auth] $($ctx.Account) | Tenant: $($ctx.TenantId)" -ForegroundColor DarkGray
#endregion

#region ---------- Leitura ----------
if (-not (Test-Path $ExcelPath)) { throw "Arquivo nao encontrado: $ExcelPath" }
$rows = Import-Excel -Path $ExcelPath -WorksheetName ((Get-ExcelSheetInfo -Path $ExcelPath)[0].Name)
Write-Host ("[read] {0} usuarios na planilha." -f $rows.Count) -ForegroundColor Green
#endregion

#region ---------- Cache de usuarios (com atributos atuais p/ diff) ----------
Write-Host "[cache] Carregando usuarios do tenant..." -ForegroundColor Cyan
$props = 'Id','UserPrincipalName','GivenName','Surname','DisplayName','JobTitle',
         'Department','CompanyName','OfficeLocation','MobilePhone','AccountEnabled',
         'EmployeeType','EmployeeId','OnPremisesSyncEnabled'
$byUpn = @{}
Get-MgUser -All -Property $props | ForEach-Object { $byUpn[$_.UserPrincipalName.ToLower()] = $_ }
Write-Host ("[cache] {0} usuarios." -f $byUpn.Count) -ForegroundColor DarkGray
#endregion

#region ---------- Relatorio ----------
$report = [System.Collections.Generic.List[object]]::new()
function Log-Result {
    param($Upn,$Phase,$Status,$Detail)
    $report.Add([pscustomobject]@{ Upn=$Upn; Phase=$Phase; Status=$Status; Detail=$Detail })
    $color = switch -Wildcard ($Status) {
        'UPDATED'    { 'Green' }; 'WOULD-UPDATE' { 'Cyan' }
        'MGR-SET'    { 'Green' }; 'WOULD-SET-MGR' { 'Cyan' }
        'UNCHANGED'  { 'DarkGray' }; 'MGR-NONE' { 'DarkGray' }
        'ERROR-*'    { 'Red' }; default { 'White' }
    }
    Write-Host ("  [{0,-15}] {1,-32} {2} {3}" -f $Status,$Upn,$Phase,$Detail) -ForegroundColor $color
}

function Get-GraphFailureStatus {
    param(
        [object]$User,
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [string]$DefaultStatus
    )

    $message = $ErrorRecord.Exception.Message
    if ($User.OnPremisesSyncEnabled -eq $true) { return 'ERROR-SYNC' }
    if ($message -match 'Insufficient privileges|Authorization_RequestDenied|Forbidden') { return 'ERROR-PRIV' }
    if ($message -match 'undergoing migration|currently undergoing migration') { return 'ERROR-MIGRATING' }
    if ($message -match 'on-premises mastered|Directory Sync') { return 'ERROR-SYNC' }
    return $DefaultStatus
}
#endregion

$mode = if ($Execute) { 'EXECUCAO' } else { 'DRY-RUN (simulacao)' }
Write-Host "`n========== MODO: $mode ==========" -ForegroundColor Magenta

#region ---------- FASE 1: Atributos ----------
Write-Host "`n--- FASE 1: Atributos ---" -ForegroundColor Magenta
# Graph property -> coluna planilha (atributos escalares simples)
$map = [ordered]@{
    GivenName      = 'givenName'
    Surname        = 'surname'
    DisplayName    = 'displayName'
    JobTitle       = 'jobTitle'
    Department     = 'department'
    CompanyName    = 'companyName'
    OfficeLocation = 'officeLocation'
    MobilePhone    = 'mobilePhone'
}

foreach ($r in $rows) {
    $upn = "$($r.userPrincipalName)".Trim().ToLower()
    if (-not $upn) { continue }
    $u = $byUpn[$upn]
    if (-not $u) { Log-Result $upn 'ATTR' 'ERROR-USER-NOTFOUND' 'inexistente no tenant'; continue }

    $patch  = @{}
    $changed = [System.Collections.Generic.List[string]]::new()

    foreach ($g in $map.Keys) {
        $desired = "$($r.($map[$g]))".Trim()
        if ($desired -eq '') { continue }
        if ("$($u.$g)" -ne $desired) { $patch[$g] = $desired; $changed.Add($g) }
    }

    # accountEnabled (bool)
    $desiredEnabled = [bool]$r.accountEnabled
    if ($u.AccountEnabled -ne $desiredEnabled) { $patch['AccountEnabled'] = $desiredEnabled; $changed.Add('AccountEnabled') }

    # Diretoria (CEO/COO/CCO/CIO/CDO): coluna extensionAttribute1 da planilha -> employeeType/employeeId
    $desiredDir = "$($r.extensionAttribute1)".Trim()
    if ($desiredDir -ne '' -and "$($u.EmployeeType)" -ne $desiredDir) {
        $patch['EmployeeType'] = $desiredDir
        $changed.Add('EmployeeType')
    }
    if ($desiredDir -ne '' -and "$($u.EmployeeId)" -ne $desiredDir) {
        $patch['EmployeeId'] = $desiredDir
        $changed.Add('EmployeeId')
    }

    if ($changed.Count -eq 0) { Log-Result $upn 'ATTR' 'UNCHANGED' ''; continue }
    $detail = ($changed -join ',')

    if (-not $Execute) { Log-Result $upn 'ATTR' 'WOULD-UPDATE' $detail; continue }

    if ($u.OnPremisesSyncEnabled -eq $true) {
        Log-Result $upn 'ATTR' 'ERROR-SYNC' "usuario sincronizado do AD local; aplicar no on-prem: $detail"
        continue
    }

    try {
        Update-MgUser -UserId $u.Id @patch -ErrorAction Stop
        Log-Result $upn 'ATTR' 'UPDATED' $detail
    } catch {
        $status = Get-GraphFailureStatus -User $u -ErrorRecord $_ -DefaultStatus 'ERROR-UPDATE'
        Log-Result $upn 'ATTR' $status $_.Exception.Message
    }
}
#endregion

#region ---------- FASE 2: Manager ----------
if (-not $SkipManager) {
    Write-Host "`n--- FASE 2: Manager ---" -ForegroundColor Magenta
    foreach ($r in $rows) {
        $upn = "$($r.userPrincipalName)".Trim().ToLower()
        $u = $byUpn[$upn]
        if (-not $u) { continue }

        $mgrRaw = "$($r.manager)".Trim()
        if ($mgrRaw -eq '') { Log-Result $upn 'MGR' 'MGR-NONE' ''; continue }

        $mgrUpn = $mgrRaw.ToLower()
        # so aceita se parecer email e existir no tenant (sem chutar correcoes)
        if ($mgrUpn -notmatch '^[^@\s]+@[^@\s]+$' -or -not $byUpn.ContainsKey($mgrUpn)) {
            Log-Result $upn 'MGR' 'ERROR-MANAGER-NOTFOUND' $mgrRaw
            continue
        }
        $mgrId = $byUpn[$mgrUpn].Id

        if (-not $Execute) { Log-Result $upn 'MGR' 'WOULD-SET-MGR' $mgrRaw; continue }

        if ($u.OnPremisesSyncEnabled -eq $true) {
            Log-Result $upn 'MGR' 'ERROR-SYNC' "usuario sincronizado do AD local; aplicar manager no on-prem: $mgrRaw"
            continue
        }

        try {
            $ref = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/users/$mgrId" }
            Set-MgUserManagerByRef -UserId $u.Id -BodyParameter $ref -ErrorAction Stop
            Log-Result $upn 'MGR' 'MGR-SET' $mgrRaw
        } catch {
            $status = Get-GraphFailureStatus -User $u -ErrorRecord $_ -DefaultStatus 'ERROR-MANAGER-SET'
            Log-Result $upn 'MGR' $status $_.Exception.Message
        }
    }
}
#endregion

#region ---------- Resumo ----------
$report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
Write-Host "`n========== RESUMO ==========" -ForegroundColor Magenta
$report | Group-Object Phase,Status | Sort-Object Name |
    ForEach-Object { Write-Host ("  {0,-30} {1}" -f $_.Name, $_.Count) }
Write-Host ("`nRelatorio: {0}" -f (Resolve-Path $ReportPath)) -ForegroundColor Green
if (-not $Execute) { Write-Host "`n>>> DRY-RUN. Revise o CSV e rode com -Execute." -ForegroundColor Yellow }
#endregion
