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
    [ValidateNotNullOrEmpty()]
    [string]$Domain,

    [string]$Environment = 'Test'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
        exit 1
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
            Write-Warning "  GPO '$gpoName' already exists — stale entries will be cleared before rewriting."
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

    switch ($cfg.Mode) {

        'Indexed' {
            # Remove stale entries left from previous runs before writing fresh values.
            if ($PSCmdlet.ShouldProcess("$cfgKey existing allowlist entries", 'Remove-GPRegistryValue')) {
                $existing = Get-GPRegistryValue -Name $gpoName -Domain $Domain `
                    -Key $cfg.RegistryKey -ErrorAction SilentlyContinue
                if ($existing) {
                    @($existing) | ForEach-Object {
                        Remove-GPRegistryValue -Name $gpoName -Domain $Domain `
                            -Key $cfg.RegistryKey -ValueName $_.ValueName | Out-Null
                    }
                }
            }

            # Chrome / Edge: numbered REG_SZ values (1, 2, 3 ...) under the allowlist key.
            $index = 1
            foreach ($id in $ids) {
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
                Write-Host "  Wrote $($ids.Count) allowlist entr$(if ($ids.Count -eq 1) { 'y' } else { 'ies' })."
            }
        }

        'Json' {
            # Firefox: single ExtensionSettings JSON — block-all wildcard + per-extension allows.
            # This GPO is self-contained; disable the existing Firefox blocking GPO to avoid
            # conflicts since both write to the same HKLM\...\Firefox\ExtensionSettings value.
            $settings = [ordered]@{
                '*' = [ordered]@{ installation_mode = 'blocked' }
            }
            foreach ($id in $ids) {
                $settings[$id] = [ordered]@{ installation_mode = 'allowed' }
            }
            $json = $settings | ConvertTo-Json -Compress -Depth 3

            if ($PSCmdlet.ShouldProcess("$cfgKey ExtensionSettings JSON", 'Set-GPRegistryValue')) {
                Set-GPRegistryValue -Name $gpoName -Domain $Domain `
                    -Key       $cfg.RegistryKey `
                    -ValueName $cfg.ValueName `
                    -Type      String `
                    -Value     $json | Out-Null
            }
            if (-not $WhatIfPreference) {
                Write-Host "  Wrote ExtensionSettings JSON: block-all + $($ids.Count) allowed extension(s)."
            }
            Write-Warning "  ACTION REQUIRED: disable your existing Firefox blocking GPO — it conflicts with '$gpoName' on the ExtensionSettings key."
        }
    }
}

#endregion

Write-Host "`nDone. All GPOs are unlinked. Link them to the target OU after review before applying to production."
