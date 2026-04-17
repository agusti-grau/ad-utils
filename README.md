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

---

## New-BrowserExtensionAllowlistGPO.ps1

Reads browser extension IDs from an Excel file and creates one unlinked GPO per browser (Chrome, Edge, Firefox) that allowlists those extensions. GPOs are never linked automatically — an admin must link them to the target OU before they take effect.

### Requirements

| Requirement | How to install |
|-------------|---------------|
| PowerShell 5.1+ | Built into Windows |
| [ImportExcel](https://github.com/dfinke/ImportExcel) module | `Install-Module ImportExcel -Scope CurrentUser` |
| GroupPolicy module | Install RSAT: `Add-WindowsCapability -Online -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0` |

### Parameters

| Parameter | Mandatory | Default | Description |
|-----------|-----------|---------|-------------|
| `-ExcelPath` | Yes | — | Path to the input Excel file (`.xlsx`, `.xls`, `.xlsm`). Must contain `Browser` and `Extension ID` columns. |
| `-Domain` | Yes | — | Target AD domain FQDN (e.g. `contoso.com`). |
| `-Environment` | No | `Test` | Label used in the GPO name. Change to `Prod` for production GPOs. |

### GPO naming convention

```
<Environment> - Browser - <BrowserName> - Extension policy
```

Examples: `Test - Browser - Chrome - Extension policy`, `Prod - Browser - Firefox - Extension policy`.

### Usage

```powershell
# Dry run — recommended before any GPO changes
.\New-BrowserExtensionAllowlistGPO.ps1 -ExcelPath ".\extensions.xlsx" -Domain "contoso.com" -WhatIf

# Create GPOs (unlinked, Test environment)
.\New-BrowserExtensionAllowlistGPO.ps1 -ExcelPath ".\extensions.xlsx" -Domain "contoso.com"

# Production naming
.\New-BrowserExtensionAllowlistGPO.ps1 -ExcelPath ".\extensions.xlsx" -Domain "contoso.com" -Environment "Prod"
```

### Input format

Excel file with `Browser` and `Extension ID` columns:

| Browser | Extension ID |
|---------|-------------|
| Chrome  | cjpalhdlnbpafiamejdnhcphjbkeiagm |
| Chrome  | aapbdbdomjkkjkaonfhkkikfgjllcleb |
| Edge    | jmjflgjpcpepeafmmgdpfkogkghcpiha |
| Firefox | uBlock0@raymondhill.net |

### How it works per browser

| Browser | Policy key | Mechanism |
|---------|-----------|-----------|
| Chrome | `HKLM\SOFTWARE\Policies\Google\Chrome\ExtensionInstallAllowlist` | Numbered REG_SZ values (`1`, `2`, `3`…) |
| Edge | `HKLM\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallAllowlist` | Numbered REG_SZ values (`1`, `2`, `3`…) |
| Firefox | `HKLM\SOFTWARE\Policies\Mozilla\Firefox\ExtensionSettings` | Single JSON string with block-all wildcard + per-extension `allowed` entries |

### Notes

- All GPOs are created **unlinked**. Link them to the appropriate OU after review.
- **Firefox conflict:** Chrome and Edge use separate registry keys for blocklist and allowlist, so both GPOs coexist. Firefox uses a single `ExtensionSettings` JSON value — the new GPO embeds a block-all rule to remain self-contained. **Disable the existing Firefox blocking GPO** to avoid the two GPOs conflicting on the same registry key.
- Re-running the script on an existing GPO is safe — stale allowlist entries are cleared before writing the new set.
- Rows with an empty Browser or Extension ID are silently skipped.
