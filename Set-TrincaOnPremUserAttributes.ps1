<#
.SYNOPSIS
    Aplica atributos TRINCA no Active Directory on-premises para usuarios DirSync.

.DESCRIPTION
    Use este script depois de revisar possiveis duplicados no Entra ID/AD.
    Ele le a planilha de atributos e atualiza no AD local os campos que o Graph
    nao consegue gravar em objetos sincronizados.

    Mapeamento principal:
      givenName            -> givenName
      surname              -> sn
      displayName          -> displayName
      jobTitle             -> title
      department           -> department
      companyName          -> company
      officeLocation       -> physicalDeliveryOfficeName
      mobilePhone          -> mobile
      extensionAttribute1  -> extensionAttribute1
      manager              -> manager (distinguishedName do manager)

    Dry-run por padrao. So altera com -Execute.

.PARAMETER ExcelPath
    Caminho do .xlsx de atributos.

.PARAMETER Execute
    Aplica de fato. Sem este switch, roda em simulacao.

.PARAMETER SkipManager
    Nao altera manager.

.PARAMETER SearchBase
    OU/base DN opcional para limitar buscas no AD.

.PARAMETER Server
    Domain Controller opcional.

.EXAMPLE
    .\Set-TrincaOnPremUserAttributes.ps1 -ExcelPath .\users_final_para_seven-91_PRONTO.xlsx

.EXAMPLE
    .\Set-TrincaOnPremUserAttributes.ps1 -ExcelPath .\users_final_para_seven-91_PRONTO.xlsx -Execute
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ExcelPath,

    [switch]$Execute,

    [switch]$SkipManager,

    [string]$SearchBase,

    [string]$Server,

    [string]$ReportPath = ".\OnPremAttrReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Modulo nao encontrado: $Name. Instale/rode em uma maquina com RSAT/ActiveDirectory e ImportExcel."
    }
    Import-Module $Name -ErrorAction Stop
}

function Normalize-Text {
    param($Value)
    return "$Value".Trim()
}

function Escape-LdapFilterValue {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    return $Value.Replace('\','\5c').Replace('*','\2a').Replace('(','\28').Replace(')','\29').Replace([string][char]0,'\00')
}

function Get-AdParams {
    $p = @{}
    if (-not [string]::IsNullOrWhiteSpace($SearchBase)) { $p.SearchBase = $SearchBase }
    if (-not [string]::IsNullOrWhiteSpace($Server)) { $p.Server = $Server }
    return $p
}

function Find-AdUserByIdentityData {
    param(
        [string]$Upn,
        [string]$Mail
    )

    $adParams = Get-AdParams
    $props = @(
        'UserPrincipalName','mail','proxyAddresses','givenName','sn','displayName','title',
        'department','company','physicalDeliveryOfficeName','mobile','extensionAttribute1',
        'manager','distinguishedName'
    )

    $filters = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($Upn)) {
        $e = Escape-LdapFilterValue $Upn
        $filters.Add("(userPrincipalName=$e)")
        $filters.Add("(mail=$e)")
        $filters.Add("(proxyAddresses=smtp:$e)")
        $filters.Add("(proxyAddresses=SMTP:$e)")
    }
    if (-not [string]::IsNullOrWhiteSpace($Mail) -and $Mail -ne $Upn) {
        $e = Escape-LdapFilterValue $Mail
        $filters.Add("(mail=$e)")
        $filters.Add("(proxyAddresses=smtp:$e)")
        $filters.Add("(proxyAddresses=SMTP:$e)")
    }

    if ($filters.Count -eq 0) { return @() }
    $ldap = "(|$($filters -join ''))"
    return @(Get-ADUser -LDAPFilter $ldap -Properties $props @adParams)
}

function Find-AdUserByUpnOnly {
    param([string]$Upn)
    if ([string]::IsNullOrWhiteSpace($Upn)) { return $null }

    $matches = Find-AdUserByIdentityData -Upn $Upn -Mail ''
    if ($matches.Count -eq 1) { return $matches[0] }
    return $null
}

function Add-ReplaceIfChanged {
    param(
        [hashtable]$Replace,
        [System.Collections.Generic.List[string]]$Changed,
        [object]$User,
        [string]$AdAttr,
        [string]$Desired
    )

    $value = Normalize-Text $Desired
    if ($value -eq '') { return }
    if ("$($User.$AdAttr)" -ne $value) {
        $Replace[$AdAttr] = $value
        $Changed.Add($AdAttr)
    }
}

function Log-Result {
    param($Upn,$Phase,$Status,$Detail)
    $script:Report.Add([pscustomobject]@{
        Upn = $Upn
        Phase = $Phase
        Status = $Status
        Detail = $Detail
    })

    $color = switch -Wildcard ($Status) {
        'UPDATED' { 'Green' }
        'WOULD-UPDATE' { 'Cyan' }
        'UNCHANGED' { 'DarkGray' }
        'MGR-SET' { 'Green' }
        'WOULD-SET-MGR' { 'Cyan' }
        'MGR-NONE' { 'DarkGray' }
        'ERROR-*' { 'Red' }
        default { 'White' }
    }
    Write-Host ("  [{0,-18}] {1,-32} {2} {3}" -f $Status,$Upn,$Phase,$Detail) -ForegroundColor $color
}

Ensure-Module ImportExcel
Ensure-Module ActiveDirectory

if (-not (Test-Path $ExcelPath)) { throw "Arquivo nao encontrado: $ExcelPath" }

$rows = Import-Excel -Path $ExcelPath -WorksheetName ((Get-ExcelSheetInfo -Path $ExcelPath)[0].Name)
Write-Host ("[read] {0} usuarios na planilha." -f $rows.Count) -ForegroundColor Green

$script:Report = [System.Collections.Generic.List[object]]::new()
$resolvedUsers = @{}

$mode = if ($Execute) { 'EXECUCAO' } else { 'DRY-RUN (simulacao)' }
Write-Host "`n========== MODO: $mode ==========" -ForegroundColor Magenta

Write-Host "`n--- FASE 1: Atributos AD ---" -ForegroundColor Magenta

foreach ($r in $rows) {
    $upn = (Normalize-Text $r.userPrincipalName).ToLower()
    $mail = (Normalize-Text $r.mail).ToLower()
    if ($upn -eq '') { continue }

    $matches = Find-AdUserByIdentityData -Upn $upn -Mail $mail
    if ($matches.Count -eq 0) {
        Log-Result $upn 'ATTR' 'ERROR-USER-NOTFOUND' 'usuario inexistente no AD local'
        continue
    }
    if ($matches.Count -gt 1) {
        $ids = ($matches | Select-Object -ExpandProperty UserPrincipalName) -join '; '
        Log-Result $upn 'ATTR' 'ERROR-DUPLICATE' "mais de um candidato no AD: $ids"
        continue
    }

    $u = $matches[0]
    $resolvedUsers[$upn] = $u

    $replace = @{}
    $changed = [System.Collections.Generic.List[string]]::new()

    Add-ReplaceIfChanged $replace $changed $u 'givenName' $r.givenName
    Add-ReplaceIfChanged $replace $changed $u 'sn' $r.surname
    Add-ReplaceIfChanged $replace $changed $u 'displayName' $r.displayName
    Add-ReplaceIfChanged $replace $changed $u 'title' $r.jobTitle
    Add-ReplaceIfChanged $replace $changed $u 'department' $r.department
    Add-ReplaceIfChanged $replace $changed $u 'company' $r.companyName
    Add-ReplaceIfChanged $replace $changed $u 'physicalDeliveryOfficeName' $r.officeLocation
    Add-ReplaceIfChanged $replace $changed $u 'mobile' $r.mobilePhone
    Add-ReplaceIfChanged $replace $changed $u 'extensionAttribute1' $r.extensionAttribute1

    if ($changed.Count -eq 0) {
        Log-Result $upn 'ATTR' 'UNCHANGED' ''
        continue
    }

    $detail = $changed -join ','
    if (-not $Execute) {
        Log-Result $upn 'ATTR' 'WOULD-UPDATE' $detail
        continue
    }

    try {
        $adParams = Get-AdParams
        Set-ADUser -Identity $u.DistinguishedName -Replace $replace @adParams -ErrorAction Stop
        Log-Result $upn 'ATTR' 'UPDATED' $detail
    } catch {
        Log-Result $upn 'ATTR' 'ERROR-UPDATE' $_.Exception.Message
    }
}

if (-not $SkipManager) {
    Write-Host "`n--- FASE 2: Manager AD ---" -ForegroundColor Magenta

    foreach ($r in $rows) {
        $upn = (Normalize-Text $r.userPrincipalName).ToLower()
        if ($upn -eq '') { continue }

        $u = $resolvedUsers[$upn]
        if (-not $u) { continue }

        $mgrRaw = (Normalize-Text $r.manager).ToLower()
        if ($mgrRaw -eq '') {
            Log-Result $upn 'MGR' 'MGR-NONE' ''
            continue
        }
        if ($mgrRaw -notmatch '^[^@\s]+@[^@\s]+$') {
            Log-Result $upn 'MGR' 'ERROR-MANAGER-NOTFOUND' $mgrRaw
            continue
        }

        $mgr = Find-AdUserByUpnOnly -Upn $mgrRaw
        if (-not $mgr) {
            Log-Result $upn 'MGR' 'ERROR-MANAGER-NOTFOUND' $mgrRaw
            continue
        }

        if ("$($u.manager)" -eq "$($mgr.DistinguishedName)") {
            Log-Result $upn 'MGR' 'UNCHANGED' $mgrRaw
            continue
        }

        if (-not $Execute) {
            Log-Result $upn 'MGR' 'WOULD-SET-MGR' $mgrRaw
            continue
        }

        try {
            $adParams = Get-AdParams
            Set-ADUser -Identity $u.DistinguishedName -Manager $mgr.DistinguishedName @adParams -ErrorAction Stop
            Log-Result $upn 'MGR' 'MGR-SET' $mgrRaw
        } catch {
            Log-Result $upn 'MGR' 'ERROR-MANAGER-SET' $_.Exception.Message
        }
    }
}

$Report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

Write-Host "`n========== RESUMO ==========" -ForegroundColor Magenta
$Report | Group-Object Phase,Status | Sort-Object Name |
    ForEach-Object { Write-Host ("  {0,-34} {1}" -f $_.Name, $_.Count) }

Write-Host ("`nRelatorio: {0}" -f (Resolve-Path $ReportPath)) -ForegroundColor Green
if (-not $Execute) {
    Write-Host "`n>>> DRY-RUN. Revise o CSV e rode com -Execute no AD on-premises." -ForegroundColor Yellow
}
