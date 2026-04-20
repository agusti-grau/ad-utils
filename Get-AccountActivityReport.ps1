#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication, ActiveDirectory, ImportExcel

<#
.SYNOPSIS
    Reports latest observed activity for AD/Entra accounts filtered by one or more employeeType values.

.DESCRIPTION
    Consolidates activity from:
      - Entra ID interactive and non-interactive sign-ins (Graph API, cert auth)
      - AD LastLogon queried from every DC (max across DCs)
      - DC Security log: Event 4768 (Kerberos), 4776 (NTLM), 4624 LogonType 3 (Network/LDAP)
      - DC Directory Service log: Event 2889 (LDAP unsigned bind — requires diagnostic logging)

    One row per account. Columns include per-source latest timestamp and an overall LastSeen.
    Accounts with no activity in the window are highlighted in the output.

.PARAMETER TenantId
    Entra tenant ID (GUID).

.PARAMETER ClientId
    App registration client ID for Graph authentication.

.PARAMETER CertThumbprint
    Thumbprint of the cert in the local machine/user store used for Graph auth.

.PARAMETER EmployeeType
    One or more employeeType values to include (e.g. 'Contractor', 'Vendor', 'SA').
    Accepts an array: -EmployeeType 'SA','G1'

.PARAMETER OutputPath
    Output Excel file path. Defaults to AccountActivityReport_<timestamp>.xlsx in the current directory.

.PARAMETER DaysBack
    Lookback window in days. Defaults to 60.

.NOTES
    Event 4624 LogonType 3 covers all network logons on DCs (LDAP, SMB, etc.) — there is no
    standard Windows event that distinguishes LDAP/LDAPS from other network protocols.
    Event 2889 is LDAP-specific but requires LDAP Interface Events logging to be enabled on each DC.

    Required Graph API permissions: AuditLog.Read.All, User.Read.All

.EXAMPLE
    .\Get-AccountActivityReport.ps1 `
        -TenantId      "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId      "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -CertThumbprint "AABBCCDD..." `
        -EmployeeType  "Contractor"

.EXAMPLE
    .\Get-AccountActivityReport.ps1 `
        -TenantId      "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId      "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -CertThumbprint "AABBCCDD..." `
        -EmployeeType  "SA","G1" `
        -DaysBack      90
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ if ($_ -notmatch '^[0-9a-fA-F-]{36}$') { throw "TenantId must be a GUID." } $true })]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [ValidateScript({ if ($_ -notmatch '^[0-9a-fA-F-]{36}$') { throw "ClientId must be a GUID." } $true })]
    [string]$ClientId,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$CertThumbprint,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string[]]$EmployeeType,

    [string]$OutputPath,

    [ValidateRange(1, 365)]
    [int]$DaysBack = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$since    = (Get-Date).AddDays(-$DaysBack).ToUniversalTime()
$sinceStr = $since.ToString("yyyy-MM-ddTHH:mm:ssZ")
$sinceXml = $since.ToString("yyyy-MM-ddTHH:mm:ss.000Z")

if (-not $OutputPath) {
    $OutputPath = ".\AccountActivityReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
}

#region Helpers

function Update-MaxTimestamp ([hashtable]$record, [string]$field, $ts) {
    if ($null -eq $ts) { return }
    if ($null -eq $record[$field] -or $ts -gt $record[$field]) { $record[$field] = $ts }
}

function Build-AccountXPath ([string[]]$names, [string]$fieldName) {
    ($names | ForEach-Object { "Data[@Name='$fieldName']='$_'" }) -join ' or '
}

#endregion

#region Graph — connect and fetch target users

Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertThumbprint -NoWelcome

$graphFilter = ($EmployeeType | ForEach-Object { "employeeType eq '$_'" }) -join ' or '
Write-Host "Fetching users where $graphFilter ..."

$userMap = @{}   # userId (GUID) → activity record
$samMap  = @{}   # samAccountName (lowercase) → activity record

$uri = ("https://graph.microsoft.com/v1.0/users" +
        "?`$filter=$graphFilter" +
        "&`$select=id,userPrincipalName,displayName,employeeType,onPremisesSamAccountName" +
        "&`$top=999")

do {
    $page = Invoke-MgGraphRequest -Uri $uri -Method GET -OutputType PSObject
    foreach ($u in $page.value) {
        $record = @{
            SamAccountName                = $u.onPremisesSamAccountName
            UserPrincipalName             = $u.userPrincipalName
            DisplayName                   = $u.displayName
            EmployeeType                  = $u.employeeType
            EntraInteractiveLastSignIn    = $null
            EntraNonInteractiveLastSignIn = $null
            ADLastLogon                   = $null
            KerberosLastAuth              = $null
            NTLMLastAuth                  = $null
            LDAPNetworkLastLogon          = $null
            LDAPUnsignedLastBind          = $null
        }
        $userMap[$u.id] = $record
        if ($u.onPremisesSamAccountName) {
            $samMap[$u.onPremisesSamAccountName.ToLower()] = $record
        }
    }
    $uri = $page.'@odata.nextLink'
} while ($uri)

if ($userMap.Count -eq 0) {
    Write-Warning "No accounts found matching employeeType filter: $graphFilter"
    exit 0
}

Write-Host "  Found $($userMap.Count) account(s)."

#endregion

#region Graph — sign-ins (interactive + non-interactive)

Write-Host "Fetching Entra ID sign-ins (last $DaysBack days)..."

$signinUri = ("https://graph.microsoft.com/v1.0/auditLogs/signIns" +
              "?`$filter=createdDateTime ge $sinceStr" +
              "&`$select=userId,createdDateTime,signInEventTypes" +
              "&`$top=999")

$countInteractive = 0; $countNonInteractive = 0

do {
    $page = Invoke-MgGraphRequest -Uri $signinUri -Method GET -OutputType PSObject
    foreach ($s in $page.value) {
        if (-not $userMap.ContainsKey($s.userId)) { continue }
        $record = $userMap[$s.userId]
        $ts     = [datetime]$s.createdDateTime
        if ($s.signInEventTypes -contains 'nonInteractiveUser') {
            Update-MaxTimestamp $record 'EntraNonInteractiveLastSignIn' $ts
            $countNonInteractive++
        } else {
            Update-MaxTimestamp $record 'EntraInteractiveLastSignIn' $ts
            $countInteractive++
        }
    }
    $signinUri = $page.'@odata.nextLink'
} while ($signinUri)

Write-Host "  Interactive sign-in events matched: $countInteractive. Non-interactive: $countNonInteractive."

#endregion

#region AD — discover DCs; skip on-prem queries if no SAM names

$samNames = @($samMap.Keys)

if ($samNames.Count -eq 0) {
    Write-Warning "No accounts have an onPremisesSamAccountName — skipping all DC queries."
} else {
    Write-Host "Discovering Domain Controllers..."
    $dcs = @(Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName)
    Write-Host "  Found $($dcs.Count) DC(s): $($dcs -join ', ')"

    $batchSize = 25

    #region AD — LastLogon per DC

    Write-Host "Querying AD LastLogon from all DCs..."

    foreach ($dc in $dcs) {
        Write-Progress -Activity 'AD LastLogon' -Status $dc -PercentComplete ([array]::IndexOf($dcs, $dc) / $dcs.Count * 100)
        try {
            for ($i = 0; $i -lt $samNames.Count; $i += $batchSize) {
                $batch    = $samNames[$i..([math]::Min($i + $batchSize - 1, $samNames.Count - 1))]
                $adFilter = ($batch | ForEach-Object { "SamAccountName -eq '$_'" }) -join ' -or '
                foreach ($au in @(Get-ADUser -Server $dc -Filter $adFilter -Properties LastLogon -ErrorAction SilentlyContinue)) {
                    if ($au.LastLogon -gt 0) {
                        $key = $au.SamAccountName.ToLower()
                        if ($samMap.ContainsKey($key)) {
                            Update-MaxTimestamp $samMap[$key] 'ADLastLogon' ([DateTime]::FromFileTime($au.LastLogon))
                        }
                    }
                }
            }
        } catch { Write-Warning "  LastLogon query failed on '$dc': $_" }
    }
    Write-Progress -Activity 'AD LastLogon' -Completed

    #endregion

    #region DC Security Event Logs — 4768 (Kerberos), 4776 (NTLM), 4624 LogonType 3 (Network/LDAP)
    # XPath filters are batched to stay within query length limits.
    # Event 4624 LogonType 3 covers LDAP simple binds and all other network authentications;
    # Windows does not emit a separate event that isolates LDAP/LDAPS from SMB or RPC.

    Write-Host "Querying DC Security event logs (Events 4768, 4776, 4624)..."

    foreach ($dc in $dcs) {
        Write-Progress -Activity 'DC Security logs' -Status $dc -PercentComplete ([array]::IndexOf($dcs, $dc) / $dcs.Count * 100)

        for ($i = 0; $i -lt $samNames.Count; $i += $batchSize) {
            $batch = $samNames[$i..([math]::Min($i + $batchSize - 1, $samNames.Count - 1))]

            # 4768 — Kerberos TGT; Properties[0] = TargetUserName
            try {
                $xp = ("*[System[(EventID=4768) and TimeCreated[@SystemTime>='$sinceXml']]] and " +
                       "*[EventData[" + (Build-AccountXPath $batch 'TargetUserName') + "]]")
                Get-WinEvent -ComputerName $dc -LogName Security -FilterXPath $xp -ErrorAction SilentlyContinue |
                    ForEach-Object {
                        $key = ([string]$_.Properties[0].Value).ToLower()
                        if ($samMap.ContainsKey($key)) { Update-MaxTimestamp $samMap[$key] 'KerberosLastAuth' $_.TimeCreated }
                    }
            } catch { Write-Warning "  4768 batch failed on '$dc': $_" }

            # 4776 — NTLM credential validation; Properties[1] = LogonAccount
            try {
                $xp = ("*[System[(EventID=4776) and TimeCreated[@SystemTime>='$sinceXml']]] and " +
                       "*[EventData[" + (Build-AccountXPath $batch 'LogonAccount') + "]]")
                Get-WinEvent -ComputerName $dc -LogName Security -FilterXPath $xp -ErrorAction SilentlyContinue |
                    ForEach-Object {
                        $key = ([string]$_.Properties[1].Value).ToLower()
                        if ($samMap.ContainsKey($key)) { Update-MaxTimestamp $samMap[$key] 'NTLMLastAuth' $_.TimeCreated }
                    }
            } catch { Write-Warning "  4776 batch failed on '$dc': $_" }

            # 4624 LogonType 3 — Network logon (covers LDAP and other network auth); Properties[5] = TargetUserName
            try {
                $xp = ("*[System[(EventID=4624) and TimeCreated[@SystemTime>='$sinceXml']]] and " +
                       "*[EventData[Data[@Name='LogonType']='3']] and " +
                       "*[EventData[" + (Build-AccountXPath $batch 'TargetUserName') + "]]")
                Get-WinEvent -ComputerName $dc -LogName Security -FilterXPath $xp -ErrorAction SilentlyContinue |
                    ForEach-Object {
                        $key = ([string]$_.Properties[5].Value).ToLower()
                        if ($samMap.ContainsKey($key)) { Update-MaxTimestamp $samMap[$key] 'LDAPNetworkLastLogon' $_.TimeCreated }
                    }
            } catch { Write-Warning "  4624 batch failed on '$dc': $_" }
        }
    }
    Write-Progress -Activity 'DC Security logs' -Completed

    #endregion

    #region DC Directory Service Log — Event 2889 (LDAP unsigned bind)
    # Requires 'LDAP Interface Events' diagnostic logging enabled on each DC.
    # Property[1] contains the binding account DN; CN is extracted for matching.

    Write-Host "Querying DC Directory Service logs (Event 2889 — LDAP unsigned bind)..."

    foreach ($dc in $dcs) {
        try {
            $xp = "*[System[(EventID=2889) and TimeCreated[@SystemTime>='$sinceXml']]]"
            Get-WinEvent -ComputerName $dc -LogName 'Directory Service' -FilterXPath $xp -ErrorAction SilentlyContinue |
                ForEach-Object {
                    $dn  = [string]$_.Properties[1].Value
                    $key = if ($dn -match '^CN=([^,]+)') { $Matches[1].ToLower() } else { $dn.ToLower() }
                    if ($samMap.ContainsKey($key)) { Update-MaxTimestamp $samMap[$key] 'LDAPUnsignedLastBind' $_.TimeCreated }
                }
        } catch { Write-Warning "  2889 unavailable on '$dc' (diagnostic logging may not be enabled): $_" }
    }

    #endregion
}

#endregion

#region Build and export report

Write-Host "Building report..."

$sourceLabels = [ordered]@{
    EntraInteractiveLastSignIn    = 'Entra Interactive'
    EntraNonInteractiveLastSignIn = 'Entra Non-Interactive'
    ADLastLogon                   = 'AD LastLogon'
    KerberosLastAuth              = 'Kerberos (4768)'
    NTLMLastAuth                  = 'NTLM (4776)'
    LDAPNetworkLastLogon          = 'Network Logon / LDAP (4624)'
    LDAPUnsignedLastBind          = 'LDAP Unsigned Bind (2889)'
}

$report = foreach ($record in $userMap.Values) {
    $lastSeen = $null; $lastSeenSource = 'No activity found'
    foreach ($kv in $sourceLabels.GetEnumerator()) {
        $ts = $record[$kv.Key]
        if ($ts -and ($null -eq $lastSeen -or $ts -gt $lastSeen)) {
            $lastSeen = $ts; $lastSeenSource = $kv.Value
        }
    }

    [PSCustomObject][ordered]@{
        SamAccountName                = $record['SamAccountName']
        UserPrincipalName             = $record['UserPrincipalName']
        DisplayName                   = $record['DisplayName']
        EmployeeType                  = $record['EmployeeType']
        EntraInteractiveLastSignIn    = $record['EntraInteractiveLastSignIn']
        EntraNonInteractiveLastSignIn = $record['EntraNonInteractiveLastSignIn']
        ADLastLogon                   = $record['ADLastLogon']
        KerberosLastAuth              = $record['KerberosLastAuth']
        NTLMLastAuth                  = $record['NTLMLastAuth']
        LDAPNetworkLastLogon          = $record['LDAPNetworkLastLogon']
        LDAPUnsignedLastBind          = $record['LDAPUnsignedLastBind']
        LastSeenOverall               = $lastSeen
        LastSeenSource                = $lastSeenSource
    }
}

$noActivityHighlight = New-ConditionalText -Text 'No activity found' `
    -BackgroundColor '#FFD7D7' -ConditionalTextColor '#CC0000'

$report | Sort-Object LastSeenOverall |
    Export-Excel -Path $OutputPath -AutoSize -BoldTopRow -FreezeTopRow `
        -ConditionalText $noActivityHighlight

Write-Host "Done. $($userMap.Count) account(s) reported → $OutputPath"

#endregion
