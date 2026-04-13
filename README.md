# Invoke-ProfilePurge

> **PowerShell script for automated local user profile purge across Windows server fleets.**

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)](https://github.com/PowerShell/PowerShell)
[![PS7 Parallel](https://img.shields.io/badge/PS7-Parallel%20mode-blueviolet?logo=powershell)](https://github.com/PowerShell/PowerShell)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

📖 [Lire en français](README_FR.md)

---

## Overview

`Invoke-ProfilePurge` is a production-grade PowerShell script designed for Windows sysadmins who need to regularly clean up stale user profiles on servers (RDS, file servers, VDI, domain controllers). It combines three cleanup phases, a structured logging system, HTML reports, and optional parallel execution.

---

## Features

| Feature | Details |
|---|---|
| **Profile purge** | Inactive profiles removed via `Win32_UserProfile` CIM (folder + registry key) |
| **Registry .bak cleanup** | Orphaned `ProfileList\*.bak` keys removed with SID → username resolution |
| **Backup folder cleanup** | `*BACKUP*` folders deleted with active session guard |
| **Domain duplicate repair** | `username` / `username.DOMAIN` pairs detected and repointed |
| **WSearch stop/start** | Releases `ntuser.dat.LOG` locks before purge |
| **Parallel execution** | PS7 `ForEach-Object -Parallel` for multi-server fleets |
| **Structured logging** | 6-level `Write-Log` with colored console + UTF-8 BOM `.log` file |
| **HTML report** | Modern light-theme report with KPI cards and filtered tables |
| **Windows Event Log** | EventID 4100 (summary) and 4199 (critical error) |
| **WhatIf simulation** | Full dry-run mode, no changes made |
| **PS5.1 + PS7** | Dual compatibility, parallel only on PS7 |

---

## Requirements

- PowerShell **5.1+** (parallel mode requires **7.0+**)
- **Local administrator** rights on each target server
- **WinRM** enabled on remote targets (`Enable-PSRemoting`)
- For Event Log source creation: admin rights on **first run only**

---

## Installation

```powershell
# Clone or download
git clone https://github.com/9LivesITSolutions/Invoke-ProfilePurge.git
cd Invoke-ProfilePurge

# Optional: allow script execution
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

No external dependencies. No module installation required.

---

## Quick Start

```powershell
# 1. Simulate first — always
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -WhatIf

# 2. Review the HTML report, then run for real
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -StopWSearch -LogPath C:\Logs

# 3. Full purge — all phases
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
| `-PassThru` | `switch` | off | Emit objects to pipeline |

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
# ── Local simulation ───────────────────────────────────────────────────────────
.\Invoke-ProfilePurge.ps1 -DaysInactive 60 -WhatIf

# ── Production purge on a server list ─────────────────────────────────────────
.\Invoke-ProfilePurge.ps1 `
    -ComputerList .\servers.txt `
    -DaysInactive 90 `
    -ExcludeFile  .\whitelist.txt `
    -PurgeProfileListBak -PurgeBackupFolders `
    -StopWSearch -WriteEventLog `
    -LogPath C:\Logs\ProfilePurge

# ── Parallel execution (PS7) ───────────────────────────────────────────────────
$cred = Get-Credential
.\Invoke-ProfilePurge.ps1 `
    -ComputerList .\servers.txt `
    -DaysInactive 90 `
    -Credential $cred `
    -Parallel -ThrottleLimit 8 `
    -WriteEventLog -LogPath C:\Logs

# ── Unknown-date profiles (review before running) ──────────────────────────────
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -DeleteUnknownDate -WhatIf
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -DeleteUnknownDate -StopWSearch

# ── Repair domain duplicates ───────────────────────────────────────────────────
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -RepairDomainDuplicates -WhatIf

# ── Export results to CSV ─────────────────────────────────────────────────────
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -PassThru |
    Export-Csv -Path C:\Logs\purge_detail.csv -NoTypeInformation -Encoding UTF8

# ── Alert on errors only ──────────────────────────────────────────────────────
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -PassThru |
    Where-Object { $_.Status -eq 'Error' } |
    Select-Object ComputerName, Identifier, Reason
```

---

## Exit Codes

| Code | Meaning |
|---|---|
| `0` | Success — all operations completed |
| `1` | Critical error — unhandled exception |
| `2` | Partial — WinRM failures or per-profile errors |

---

## Event Log

| EventID | Type | Trigger |
|---|---|---|
| `4100` | Information / Warning / Error | End-of-run summary |
| `4199` | Error | Unhandled critical exception |

EventID type maps to exit code: `0` → Information, `2` → Warning, `1` → Error.

**First-run setup** (if the event source doesn't exist yet):
```powershell
# Run once as administrator, or pre-create the source via GPO
New-EventLog -LogName Application -Source ProfilePurge
```

---

## Console Output

Only actionable entries are shown:

| Prefix | Status | Meaning |
|---|---|---|
| `DEL` | `Deleted` | Profile successfully removed |
| `SIM` | `WhatIf` | Would be removed (simulation) |
| `REP` | `Repaired` | Domain duplicate fixed |
| `ERR` | `Error` | Operation failed |
| `INF` | `Info` | Service event (WSearch restart) |
| `SKP` | `Skipped` | Significant skip (unknown date, active session on duplicate) |

`Kept`, `Excluded`, and standard `Skipped` entries are counted in the summary but not printed.

---

## Domain Duplicate Repair

When `-RepairDomainDuplicates` is active, the script detects profile folders containing a dot (`username.DOMAIN`) that have a corresponding dot-free profile (`username`).

**Actions performed:**
1. `Set-ItemProperty ProfileImagePath` → repoints the domain SID registry entry to the local profile path
2. `Remove-Item -Recurse -Force` → deletes the `.DOMAIN` profile folder

**Safety guards:**
- Session active on `.DOMAIN` profile → `Skipped`
- Local profile not viable or also eligible for deletion → `Skipped`
- No data migration — assumes local profile has the authoritative data

---

## Unknown Date Profiles

Profiles with no `LastWriteTime` (folder unreadable) and no `LastUseTime` (CIM) are **skipped by default** (`-DeleteUnknownDate` not passed).

**Why?** These are typically service accounts (`IIS APPPOOL\...`, `NT SERVICE\...`, SCCM agents) that automatically recreate their profile folder on next service startup — causing an infinite delete/recreate loop.

**Recommended workflow:**
```powershell
# 1. Identify them in the HTML report (Status: Skipped, unknown date)
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -WhatIf

# 2. Add confirmed orphans to the exclusion list
# whitelist.txt: svc_backup, svc_monitor, IIS*, ...

# 3. Force-delete remaining unknown-date profiles
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -DeleteUnknownDate -StopWSearch
```

---

## Task Scheduler Setup

```
Program : powershell.exe   (or pwsh.exe for PS7 parallel)
Arguments:
  -NonInteractive -NoProfile -ExecutionPolicy Bypass
  -File "C:\Scripts\Invoke-ProfilePurge.ps1"
  -ComputerList "C:\Scripts\servers.txt"
  -DaysInactive 90
  -PurgeProfileListBak
  -StopWSearch
  -WriteEventLog
  -LogPath "C:\Logs\ProfilePurge"
Run as: SYSTEM (or a dedicated service account with local admin on targets)
```

---

## File Structure

```
Invoke-ProfilePurge/
├── Invoke-ProfilePurge.ps1   # Main script
├── README.md                 # This file
├── README_FR.md              # French documentation
├── CHANGELOG.md              # Version history
├── LICENSE                   # MIT License
├── .gitignore
├── servers.txt               # (not tracked) server list
└── whitelist.txt             # (not tracked) exclusion list
```

---

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Test on PS5.1 AND PS7 before submitting
4. Submit a pull request with a description and test evidence

---

## License

MIT — see [LICENSE](LICENSE).

---

*9 Lives IT Solutions — Healthcare IT & Infrastructure Automation*
