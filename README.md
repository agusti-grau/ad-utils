# ad-utils

PowerShell utilities for Active Directory management tasks.

---

## Check-ADAccounts.ps1

Reads a list of SamAccountNames from an Excel file, strips a given suffix from each one, queries Active Directory to check existence, and outputs a copy of the file with a new **Exists in AD** column.

### Requirements

| Requirement | How to install |
|-------------|---------------|
| PowerShell 5.1+ | Built into Windows |
| [ImportExcel](https://github.com/dfinke/ImportExcel) module | `Install-Module ImportExcel -Scope CurrentUser` |
| ActiveDirectory module | Install RSAT: `Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0` |

### Parameters

| Parameter | Mandatory | Description |
|-----------|-----------|-------------|
| `-ExcelPath` | Yes | Path to the input Excel file (`.xlsx`, `.xls`, `.xlsm`). Must contain a `SamAccountName` column. |
| `-Suffix` | Yes | Suffix to strip from each SamAccountName before querying AD (e.g. `_TEMP`). |

### Usage

```powershell
.\Check-ADAccounts.ps1 -ExcelPath ".\users.xlsx" -Suffix "_TEMP"
```

### Input / Output

**Input** — any Excel file with a `SamAccountName` column in the first row:

| SamAccountName |
|----------------|
| john.doe_TEMP  |
| jane.smith_TEMP |

**Output** — a new file named `<original>_checked_AD<ext>` in the same directory:

| SamAccountName  | Exists in AD |
|-----------------|--------------|
| john.doe_TEMP   | True         |
| jane.smith_TEMP | False        |

### Notes

- Rows with an empty or whitespace-only SamAccountName are skipped and marked `False`.
- If the suffix is not present on a given row, the full value is queried as-is.
- Network or permission errors from AD will stop the script immediately — only "user not found" is handled silently.
- Re-running on an already-processed file is safe; the column is overwritten.
