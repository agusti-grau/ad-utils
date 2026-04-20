#Requires -Version 5.1
#Requires -Modules ImportExcel, GroupPolicy

<#
.SYNOPSIS
    Creates per-browser GPOs that allowlist extension IDs read from an Excel file.
    GPOs are created unlinked — link them to the target OU after review.

.DESCRIPTION
    Reads 'Browser' and 'Extension ID' columns from the Excel file.
    Supported browsers: Chrome, Edge, Firefox.

    Chrome / Edge: sets numbered REG_SZ values under ExtensionInstallAllowlist.
    Firefox:       sets a single ExtensionSettings JSON that blocks all extensions
                   except those listed. The existing Firefox blocking GPO must be
                   disabled to avoid conflicts — both write to the same registry key.

    GPO naming convention: "<Environment> - Browser - <BrowserName> - Extension policy"

.PARAMETER ExcelPath
    Path to the input Excel file (.xlsx / .xls / .xlsm).
    Must contain 'Browser' and 'Extension ID' columns.

.PARAMETER Domain
    Target AD domain FQDN (e.g. "contoso.com").

.PARAMETER Environment
    Environment label used in the GPO name. Defaults to 'Test'.

.EXAMPLE
    .\New-BrowserExtensionAllowlistGPO.ps1 -ExcelPath ".\extensions.xlsx" -Domain "contoso.com"

.EXAMPLE
    .\New-BrowserExtensionAllowlistGPO.ps1 -ExcelPath ".\extensions.xlsx" -Domain "contoso.com" -Environment "Prod"

.EXAMPLE
    .\New-BrowserExtensionAllowlistGPO.ps1 -ExcelPath ".\extensions.xlsx" -Domain "contoso.com" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [ValidateScript({
        if (-not (Test-Path -LiteralPath $_ -PathType Leaf)) { throw "File not found: $_" }
        if ([System.IO.Path]::GetExtension($_) -notin @('.xlsx', '.xls', '.xlsm')) { throw "Not an Excel file: $_" }
        $true
    })]
    [string]$ExcelPath,

    [Parameter(Mandatory)]
    [ValidateScript({
        if ($_ -notmatch '\.') { throw "Domain must be an FQDN (e.g. contoso.com), not a NetBIOS name." }
        $true
    })]
    [string]$Domain,

    [ValidateScript({
        if ($_ -match '[\\\/\[\]:;|=,+*?<>"&]') { throw "Environment contains characters that are invalid in a GPO name." }
        $true
    })]
    [string]$Environment = 'Test'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Extension ID validation patterns.
# Chrome/Edge: exactly 32 lowercase letters a-p (CRX hash encoding).
# Firefox:     {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx} GUID  OR  name@domain.
$idPatterns = @{
    Chrome  = '^[a-p]{32}$'
    Edge    = '^[a-p]{32}$'
    Firefox = '^(\{[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}|[a-zA-Z0-9._%-]+@[a-zA-Z0-9._-]+)$'
}

function Test-ExtensionIds {
    param([string[]]$ids, [string]$browser, [string]$pattern)
    $valid   = [System.Collections.Generic.List[string]]::new()
    $invalid = [System.Collections.Generic.List[string]]::new()
    foreach ($id in $ids) {
        if ($id -match $pattern) { $valid.Add($id) }
        else                     { $invalid.Add($id) }
    }
    if ($invalid.Count -gt 0) {
        Write-Warning ("  $browser — $($invalid.Count) invalid extension ID(s) skipped: " + ($invalid -join ', '))
    }
    return $valid.ToArray()
}

$browserConfig = @{
    Chrome  = @{
        RegistryKey = 'HKLM\SOFTWARE\Policies\Google\Chrome\ExtensionInstallAllowlist'
        Mode        = 'Indexed'
    }
    Edge    = @{
        RegistryKey = 'HKLM\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallAllowlist'
        Mode        = 'Indexed'
    }
    Firefox = @{
        RegistryKey = 'HKLM\SOFTWARE\Policies\Mozilla\Firefox'
        ValueName   = 'ExtensionSettings'
        Mode        = 'Json'
    }
}

#region Load and validate data

$rows = Import-Excel -Path $ExcelPath

if ($null -eq $rows -or @($rows).Count -eq 0) {
    Write-Warning "No data rows found in '$ExcelPath'."
    exit 0
}

$rowArray = @($rows)

foreach ($col in @('Browser', 'Extension ID')) {
    if (-not ($rowArray[0].PSObject.Properties.Name -contains $col)) {
        Write-Error "Column '$col' not found in '$ExcelPath'."
    }
}

$grouped = $rowArray |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_.'Extension ID') -and
                   -not [string]::IsNullOrWhiteSpace($_.'Browser') } |
    Group-Object { $_.'Browser'.Trim() }

if (-not $grouped) {
    Write-Warning "No valid rows with extension IDs found."
    exit 0
}

#endregion

#region Create GPOs

foreach ($group in $grouped) {
    $browserName = $group.Name
    $cfgKey      = $browserConfig.Keys |
                       Where-Object { $_ -ieq $browserName } |
                       Select-Object -First 1

    if (-not $cfgKey) {
        Write-Warning "Unknown browser '$browserName' — skipping. Supported: Chrome, Edge, Firefox."
        continue
    }

    $cfg     = $browserConfig[$cfgKey]
    $gpoName = "$Environment - Browser - $cfgKey - Extension policy"
    $ids     = $group.Group |
                   ForEach-Object { ($_.'Extension ID').Trim() } |
                   Where-Object   { $_ -ne '' } |
                   Sort-Object    -Unique

    Write-Host "`n[$cfgKey] $($ids.Count) extension(s) — GPO: '$gpoName'"

    # Track whether GPO is ready for registry writes.
    # In -WhatIf mode ShouldProcess returns $false but we still proceed to show
    # what registry values would be set. With -Confirm and user declines, we skip.
    $gpoReady = $false
    if ($PSCmdlet.ShouldProcess($gpoName, 'Create GPO')) {
        if (Get-GPO -Name $gpoName -Domain $Domain -ErrorAction SilentlyContinue) {
            Write-Error "GPO '$gpoName' already exists. Delete it first or choose a different -Environment value."
        } else {
            $null = New-GPO -Name $gpoName -Domain $Domain
            Write-Host "  GPO created (unlinked)."
        }
        $gpoReady = $true
    } elseif ($WhatIfPreference) {
        $gpoReady = $true
    }

    if (-not $gpoReady) {
        Write-Warning "  GPO creation declined — skipping registry writes for '$gpoName'."
        continue
    }

    # Validate IDs against the browser-specific pattern; skip invalid ones.
    $validIds = Test-ExtensionIds -ids $ids -browser $cfgKey -pattern $idPatterns[$cfgKey]

    if ($validIds.Count -eq 0) {
        Write-Warning "  No valid extension IDs remain for $cfgKey after validation — skipping GPO write."
        continue
    }

    switch ($cfg.Mode) {

        'Indexed' {
            # Chrome / Edge: allowlist key is SEPARATE from the blocklist key.
            # ExtensionInstallBlocklist (*) and ExtensionInstallAllowlist coexist without conflict —
            # Chrome/Edge explicitly exempts allowlisted IDs from the blocklist.
            # The allowlist is only effective when a blocking GPO with ExtensionInstallBlocklist = *
            # is also applied. Without it, all extensions are already allowed and this GPO is a no-op.
            Write-Host "  Mixed policy model: blocklist (*) + allowlist (specific IDs) — no key conflict."
            Write-Warning "  ${cfgKey}: this allowlist only has effect if a separate GPO sets ExtensionInstallBlocklist = *. Without it, all extensions are already allowed."

            # Chrome / Edge: numbered REG_SZ values (1, 2, 3 ...) under the allowlist key.
            $index = 1
            foreach ($id in $validIds) {
                if ($PSCmdlet.ShouldProcess("$cfgKey entry $index = '$id'", 'Set-GPRegistryValue')) {
                    Set-GPRegistryValue -Name $gpoName -Domain $Domain `
                        -Key       $cfg.RegistryKey `
                        -ValueName ([string]$index) `
                        -Type      String `
                        -Value     $id | Out-Null
                }
                $index++
            }
            if (-not $WhatIfPreference) {
                Write-Host "  Wrote $($validIds.Count) allowlist entr$(if ($validIds.Count -eq 1) { 'y' } else { 'ies' })."
            }
        }

        'Json' {
            # Firefox: single ExtensionSettings JSON — block-all wildcard + per-extension allows.
            # This GPO is self-contained; disable the existing Firefox blocking GPO to avoid
            # conflicts since both write to the same HKLM\...\Firefox\ExtensionSettings value.
            #
            # JSON is built via [PSCustomObject] (not a piped hashtable) to guarantee PS 5.1
            # produces a clean JSON object rather than a wrapped PSCustomObject serialization.
            $jsonObj = [PSCustomObject]@{}
            $jsonObj | Add-Member -NotePropertyName '*' -NotePropertyValue ([PSCustomObject]@{ installation_mode = 'blocked' })
            foreach ($id in $validIds) {
                $jsonObj | Add-Member -NotePropertyName $id -NotePropertyValue ([PSCustomObject]@{ installation_mode = 'allowed' })
            }
            $json = $jsonObj | ConvertTo-Json -Compress -Depth 3

            # Verify the produced JSON is parseable before writing.
            try   { $null = $json | ConvertFrom-Json }
            catch { Write-Error "Firefox ExtensionSettings JSON is invalid — aborting write for '$gpoName': $_" }

            if ($PSCmdlet.ShouldProcess("$cfgKey ExtensionSettings JSON", 'Set-GPRegistryValue')) {
                # Mozilla documents ExtensionSettings as REG_MULTI_SZ on Windows GPO.
                # Value is a single-element array; Firefox concatenates all elements into one JSON string.
                Set-GPRegistryValue -Name $gpoName -Domain $Domain `
                    -Key       $cfg.RegistryKey `
                    -ValueName $cfg.ValueName `
                    -Type      MultiString `
                    -Value     @($json) | Out-Null
            }
            if (-not $WhatIfPreference) {
                Write-Host "  Wrote ExtensionSettings JSON (REG_SZ): block-all + $($validIds.Count) allowed extension(s)."
            }
            Write-Warning "  ACTION REQUIRED: disable your existing Firefox blocking GPO — it conflicts with '$gpoName' on the ExtensionSettings key."
        }
    }
}

#endregion

Write-Host "`nDone. All GPOs are unlinked. Link them to the target OU after review before applying to production."
