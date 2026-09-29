# Invoke-ProfilePurge

> **PowerShell script for automated local user profile purge across Windows server fleets.**

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)](https://github.com/PowerShell/PowerShell)
[![PS7 Parallel](https://img.shields.io/badge/PS7-Parallel%20mode-blueviolet?logo=powershell)](https://github.com/PowerShell/PowerShell)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

📖 [Lire en français](README_FR.md)

---

## Overview

`Invoke-ProfilePurge` is a production-grade PowerShell script for Windows sysadmins who need to regularly clean stale user profiles on servers (RDS, file servers, VDI, domain controllers). It combines three cleanup phases, structured logging, HTML reports, parallel execution, and Windows Event Log integration.

---

## Features

| Feature | Detail |
|---|---|
| **Profile purge** | Inactive profiles removed via `Win32_UserProfile` CIM (folder + registry key) |
| **Date criterion** | `Max(ntuser.dat LastWriteTime, CIM LastUseTime)` — WSearch-resistant |
| **Registry .bak cleanup** | Orphaned `ProfileList\*.bak` keys with SID→username resolution |
| **Backup folder cleanup** | `*BACKUP*` folders with active session guard |
| **Domain duplicate repair** | `username` / `username.DOMAIN` pairs repointed to local profile |
| **WSearch stop/start** | Releases `ntuser.dat.LOG` locks; dates computed BEFORE WSearch stops |
| **Parallel execution** | `ForEach-Object -Parallel` on PS7 with named mutex for thread-safe logging |
| **Structured logging** | 6-level `Write-Log`, colored console + UTF-8 BOM `.log` file, auto-rotation |
| **HTML report** | Light-theme report with sticky topbar, KPI strip, filtered action tables |
| **Windows Event Log** | EventID 4100 (summary) and 4199 (critical error) |
| **WhatIf simulation** | Full dry-run, no changes made |
| **PS5.1 + PS7** | Dual compatibility; parallel mode PS7 only |

---

## How the date criterion works

The inactivity criterion is `Max(ntuser.dat LastWriteTime, CIM LastUseTime)`:

| Source | Updated by | Reliable? |
|---|---|---|
| `ntuser.dat` LastWriteTime | Windows at user logoff (registry hive commit) | ✅ Yes |
| `Win32_UserProfile.LastUseTime` | Windows at logoff, stored in `ProfileList\<SID>` | ✅ Yes (may be null) |
| Profile folder LastWriteTime | Any process writing anywhere in the profile tree | ❌ Contaminated |

`ntuser.dat` is opened read-only by Windows Search for indexing — it is never modified by WSearch, antivirus, shadow copy, or backup agents. It reflects the actual last user session, making it the most reliable filesystem date source.

The two-pass architecture ensures WSearch is stopped **after** all dates are computed, preventing contamination during the current run.

---

## Requirements

- PowerShell **5.1+** (parallel mode: **7.0+**)
- **Elevated admin rights** on each target server (`SeRestorePrivilege` required for `Win32_UserProfile.Delete()`)
- **WinRM** enabled on remote targets (`Enable-PSRemoting`)
- Event Log source creation: admin rights on **first run only**

---

## Installation

```powershell
git clone https://github.com/9lives-it/Invoke-ProfilePurge.git
cd Invoke-ProfilePurge
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

No external dependencies. No module installation required.

---

## Quick Start

```powershell
# Always simulate first
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -WhatIf

# Live run — local server
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -StopWSearch -LogPath C:\Logs

# Full purge — all phases
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 `
    -PurgeProfileListBak -PurgeBackupFolders `
    -RepairDomainDuplicates -StopWSearch `
    -WriteEventLog -LogPath C:\Logs
```

---

## Parameters

### Targets

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-ComputerName` | `string[]` | — | Target servers inline |
| `-ComputerList` | `string` | — | Text file of servers (1/line, `#` ignored) |
| `-Credential` | `PSCredential` | — | Credential for WinRM sessions |

### Profile purge

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-DaysInactive` | `int` | `90` | Inactivity threshold in days |
| `-ExcludeUsers` | `string[]` | — | Inline exclusions (wildcards: `svc_*`) |
| `-ExcludeFile` | `string` | — | Whitelist file (1/line, wildcards OK) |
| `-DeleteUnknownDate` | `switch` | off | Delete profiles with no date info (⚠ service accounts) |
| `-StopWSearch` | `switch` | off | Stop Windows Search before purge |
| `-RepairDomainDuplicates` | `switch` | off | Repair `user` / `user.DOMAIN` pairs |

### Additional phases

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-PurgeProfileListBak` | `switch` | off | Remove orphaned `.bak` registry keys |
| `-PurgeBackupFolders` | `switch` | off | Delete `*BACKUP*` folders |
| `-UsersPath` | `string` | `C:\Users` | Root path for BACKUP folder search |

### Execution

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-Parallel` | `switch` | off | Parallel server processing (PS7+) |
| `-ThrottleLimit` | `int` | `5` | Max concurrent servers |
| `-WhatIf` | `switch` | off | Simulation — no changes |
| `-PassThru` | `switch` | off | Emit result objects to pipeline |

### Output

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-LogPath` | `string` | Script dir | Log + HTML report destination |
| `-LogRetentionDays` | `int` | `30` | Log file retention |
| `-ReportPath` | `string` | Auto | Custom HTML report path |
| `-WriteEventLog` | `switch` | off | Write to Windows Event Log |
| `-EventSource` | `string` | `ProfilePurge` | Event source name |
| `-EventLogName` | `string` | `Application` | Target event log |

---

## Usage Examples

```powershell
# Simulate locally
.\Invoke-ProfilePurge.ps1 -DaysInactive 60 -WhatIf

# Multi-server production run
.\Invoke-ProfilePurge.ps1 `
    -ComputerList .\servers.txt `
    -DaysInactive 90 `
    -ExcludeFile  .\whitelist.txt `
    -PurgeProfileListBak -PurgeBackupFolders `
    -StopWSearch -WriteEventLog `
    -LogPath C:\Logs\ProfilePurge

# Parallel execution (PS7)
$cred = Get-Credential
.\Invoke-ProfilePurge.ps1 `
    -ComputerList .\servers.txt -DaysInactive 90 `
    -Credential $cred -Parallel -ThrottleLimit 8 `
    -WriteEventLog -LogPath C:\Logs

# Repair domain duplicates (username vs username.DOMAIN)
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -RepairDomainDuplicates -WhatIf

# Unknown-date profiles — review first, then force
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -DeleteUnknownDate -WhatIf
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -DeleteUnknownDate -StopWSearch

# Export results to CSV
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -PassThru |
    Export-Csv -Path C:\Logs\purge.csv -NoTypeInformation -Encoding UTF8

# Alert on errors only
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -PassThru |
    Where-Object { $_.Status -eq 'Error' } |
    Select-Object ComputerName, Identifier, Reason
```

---

## Exit Codes

| Code | Meaning |
|---|---|
| `0` | Success |
| `1` | Critical unhandled error |
| `2` | Partial — WinRM failures or per-profile errors |

---

## Windows Event Log

| EventID | Type | Trigger |
|---|---|---|
| `4100` | Information / Warning / Error | End-of-run summary (type maps to exit code) |
| `4199` | Error | Unhandled critical exception + stack trace |

```powershell
# Create the source once (requires admin, only needed on first run)
New-EventLog -LogName Application -Source ProfilePurge
```

---

## Console Output

Only actionable entries are shown — `Kept`, `Excluded`, and standard `Skipped` are counted in the summary but not printed:

| Prefix | Status | Meaning |
|---|---|---|
| `DEL` | `Deleted` | Profile removed |
| `SIM` | `WhatIf` | Would be removed (simulation) |
| `REP` | `Repaired` | Domain duplicate fixed |
| `ERR` | `Error` | Operation failed |
| `INF` | `Info` | Service event (WSearch restart) |
| `SKP` | `Skipped` | Significant skip (unknown date, active session on duplicate) |

---

## Domain Duplicate Repair

When `-RepairDomainDuplicates` is active, the script detects folders with a dot (`username.DOMAIN`) that have a matching local profile (`username`):

1. `Set-ItemProperty ProfileImagePath` → repoints the domain SID registry entry to the local profile
2. `Remove-Item -Recurse -Force` → removes the `.DOMAIN` folder

Safety guards: skipped if the `.DOMAIN` session is active, or if the local profile is also eligible for deletion.

---

## Unknown Date Profiles

Profiles with no `ntuser.dat` and no `LastUseTime` are **skipped by default** to avoid infinite delete/recreate loops (service accounts recreate their profile on next start). Use `-DeleteUnknownDate` after reviewing the WhatIf report.

---

## Task Scheduler Setup

```
Program:   powershell.exe
Arguments: -NonInteractive -NoProfile -ExecutionPolicy Bypass
           -File "C:\Scripts\Invoke-ProfilePurge.ps1"
           -ComputerList "C:\Scripts\servers.txt"
           -DaysInactive 90 -PurgeProfileListBak -StopWSearch
           -WriteEventLog -LogPath "C:\Logs\ProfilePurge"
Run as:    SYSTEM  (or elevated admin — SeRestorePrivilege required)
```

---

## File Structure

```
Invoke-ProfilePurge/
├── Invoke-ProfilePurge.ps1   # Main script (EN)
├── README.md                 # This file
├── README_FR.md              # French documentation
├── CHANGELOG.md              # Version history
├── LICENSE                   # MIT License
├── .gitignore
├── servers.txt               # (not tracked) server list
└── whitelist.txt             # (not tracked) exclusion list
```

---

## License

MIT — see [LICENSE](LICENSE).

---

*9 Lives IT Solutions — Healthcare IT & Infrastructure Automation*
