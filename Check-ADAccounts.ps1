#Requires -Version 5.1
#Requires -Modules ImportExcel, ActiveDirectory

<#
.SYNOPSIS
    Checks Active Directory for SamAccountNames from an Excel file after stripping a given suffix.
    Outputs a copy of the file with an 'Exists in AD' column appended.

.PARAMETER ExcelPath
    Full or relative path to the input Excel file (.xlsx / .xls / .xlsm).

.PARAMETER Suffix
    Suffix to strip from each SamAccountName before querying AD (e.g. "_TEMP").

.OUTPUTS
    A new Excel file in the same directory as the input, named <original>_checked_AD<ext>.

.EXAMPLE
    .\Check-ADAccounts.ps1 -ExcelPath ".\users.xlsx" -Suffix "_TEMP"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateScript({
        if (-not (Test-Path -LiteralPath $_ -PathType Leaf)) {
            throw "File not found: $_"
        }
        if ([System.IO.Path]::GetExtension($_) -notin @('.xlsx', '.xls', '.xlsm')) {
            throw "File must be an Excel file (.xlsx / .xls / .xlsm): $_"
        }
        $true
    })]
    [string]$ExcelPath,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Suffix
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Load data

$rows = Import-Excel -Path $ExcelPath

if ($null -eq $rows -or @($rows).Count -eq 0) {
    Write-Warning "No data rows found in '$ExcelPath'. Nothing to process."
    exit 0
}

$rowArray = @($rows)

if (-not ($rowArray[0].PSObject.Properties.Name -contains 'SamAccountName')) {
    Write-Error "Column 'SamAccountName' not found in '$ExcelPath'."
    exit 1
}

#endregion

#region Process rows

$total   = $rowArray.Count
$counter = 0

$results = foreach ($row in $rowArray) {
    $counter++
    Write-Progress -Activity 'Checking AD accounts' `
                   -Status "$counter / $total" `
                   -PercentComplete ($counter / $total * 100)

    $raw = [string]$row.SamAccountName

    if ($raw.ToLower().EndsWith($Suffix.ToLower())) {
        $samToCheck = $raw.Substring(0, $raw.Length - $Suffix.Length)
    } else {
        $samToCheck = $raw
    }

    $exists = $false
    if (-not [string]::IsNullOrWhiteSpace($samToCheck)) {
        try {
            $null = Get-ADUser -Identity $samToCheck -ErrorAction Stop
            $exists = $true
        } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $exists = $false
        }
        # All other AD exceptions (network, permissions) propagate and stop the script.
    }

    $row | Add-Member -NotePropertyName 'Exists in AD' -NotePropertyValue $exists -Force -PassThru
}

Write-Progress -Activity 'Checking AD accounts' -Completed

#endregion

#region Export

$inputItem  = Get-Item -LiteralPath $ExcelPath
$outputPath = Join-Path $inputItem.DirectoryName ($inputItem.BaseName + '_checked_AD' + $inputItem.Extension)

if (Test-Path -LiteralPath $outputPath) {
    Write-Warning "Output file already exists and will be overwritten: $outputPath"
}

$results | Export-Excel -Path $outputPath -AutoSize -BoldTopRow

$found    = @($results | Where-Object { $_.'Exists in AD' -eq $true }).Count
$notFound = $total - $found

Write-Host "Done. $total account(s) checked: $found found in AD, $notFound not found."
Write-Host "Output: $outputPath"

#endregion
