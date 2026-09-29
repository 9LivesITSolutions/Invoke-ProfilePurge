# Changelog

All notable changes to **Invoke-ProfilePurge** are documented here.
Format: [Keep a Changelog](https://keepachangelog.com/en/1.0.0/)

---

## [2.7.2] - 2026-09-29

### Fixed
- **Date criterion: replaced folder `LastWriteTime` with `ntuser.dat LastWriteTime`** as the primary filesystem indicator. The profile root folder's `LastWriteTime` is contaminated by Windows Search, antivirus agents, shadow copy, and scheduled tasks writing anywhere in the profile tree — making it unreliable. `ntuser.dat` is written by Windows only at user logoff (registry hive commit), giving an accurate "last real user session" timestamp. WSearch opens `ntuser.dat` read-only for indexing and does not modify it.
- New criterion order: **`Max(ntuser.dat LastWriteTime, CIM LastUseTime)`**. Falls back to profile folder `LastWriteTime` only when both primary sources are unavailable.

---

## [2.7.1] - 2026-09-29

### Fixed
- Reverted over-correction from v2.6.1: restored `Max(LastWriteTime, LastUseTime)` as the criterion (replaced by `ntuser.dat` approach in v2.7.2).
- Script version check for `SeRestorePrivilege` warning now writes to the log file (not console only).

---

## [2.7.0] - 2026-09-29

### Changed
- **Reverted to `Remove-CimInstance`** as the profile deletion method. All intermediate approaches (direct registry deletion, `DeleteProfile()` P/Invoke, `takeown`/`icacls`) were over-engineering. `Remove-CimInstance` on `Win32_UserProfile` calls `Win32_UserProfile.Delete()` internally which handles ACL-protected folders, locked files, and registry cleanup correctly.
- Root cause of all deletion failures confirmed: script was not running with admin elevation. `SeRestorePrivilege` is required by the WMI provider for `Win32_UserProfile.Delete()` and is available in elevated admin tokens.

### Added
- Early privilege check: if `SeRestorePrivilege` cannot be enabled via `AdjustTokenPrivileges`, a clear `[WARN]` with the current identity is written to both console and log file.
- Fixed `Add-Type` C# struct visibility: `TPriv` struct must be `public` (not default `internal`) for the type to compile correctly across PS sessions.

---

## [2.6.1] - 2026-09-29 *(over-corrected, reverted in 2.7.1)*

### Changed
- Changed date criterion to `LastUseTime` primary, `LastWriteTime` fallback — caused loss of detection for profiles with null `LastUseTime`.

---

## [2.6.0] - 2026-09-29

### Changed
- **Two-pass architecture** in `$PurgeProfilesBlock`: all date computations now happen in Pass 1 (before WSearch is stopped). Pass 2 stops WSearch then executes deletions. Prevents WSearch from contaminating `LastWriteTime` during the current run by flushing its index files into profile folders at service stop.
- WSearch is no longer stopped in WhatIf mode or when no profiles are eligible for deletion.

### Fixed
- `$LogPath` default value now resolved after the `param()` block. `$PSScriptRoot` evaluates to empty inside `param()` when invoked via `powershell.exe -File`, causing a binding error on `-Path`.

---

## [2.5.8] - 2026-09-29

### Fixed
- `$LogPath = $PSScriptRoot` in `param()` block evaluates to empty when called via `powershell.exe -File`. Moved resolution to after the param block with `$MyInvocation.MyCommand.Path` fallback.

---

## [2.5.7] - 2026-04-12

### Changed
- **System account exclusion is now SID-based** (language-independent). Replaced hardcoded names (`SYSTEM`, `LOCAL SERVICE`, `NetworkService`, `Administrateur`...) with:
  - Exact SID: `S-1-5-18` (SYSTEM), `S-1-5-19` (LOCAL SERVICE), `S-1-5-20` (NETWORK SERVICE)
  - Suffix: `*-500` (built-in Administrator, all languages), `*-501` (Guest), `*-503` (DefaultAccount)
  - Machine accounts: SAM name ending with `$`
- Exclusion reason in log now identifies which rule triggered: `System SID (S-1-5-18)`, `Machine account (trailing $)`, or `Exclusion rule match`.

---

## [2.5.6] - 2026-04-11

### Fixed
- **Root cause of WhatIf in LIVE mode**: `$WhatIfPreference = 'SilentlyContinue'` assigned a non-empty string (truthy bool) — WhatIf was permanently active. Fixed to `$WhatIfPreference = $false`.
- Added `-WhatIf:$false` to `Remove-CimInstance` calls to prevent ShouldProcess propagation from parent script scope.

---

## [2.5.5] - 2026-04-11

### Fixed
- **Critical argument binding bug**: `$MergedExclusions` (`[string[]]`) was positioned before bool parameters in `ArgumentList`. PowerShell flattens arrays in `@()`, causing the string array to consume all subsequent bool arguments. Fixed by moving `[string[]]$Exclusions` to the last position in the scriptblock `param()` block.

---

## [2.5.4] - 2026-04-11

### Added
- `-PassThru` switch: pipeline output is silent by default. All result objects were printing to console in interactive mode.

---

## [2.5.3] - 2026-04-11

### Added
- `-DeleteUnknownDate`: profiles with no `LastWriteTime` and no CIM `LastUseTime` are **skipped by default**. Service accounts recreate their profile on next start, causing infinite delete/recreate loops. Use `-DeleteUnknownDate` to force deletion after manual review.

### Changed
- Console output filtered: only actionable entries (`Deleted`, `WhatIf`, `Repaired`, `Error`, significant `Skipped`) are printed. `Kept`, `Excluded`, standard `Skipped` are counted but not printed.
- HTML report redesigned: light theme, sticky topbar, KPI strip, filtered tables (non-impacted rows hidden).

---

## [2.5.2] - 2026-04-11

### Added
- `-StopWSearch`: stops Windows Search before purge to release `ntuser.dat.LOG` file handles. Restarts only if it was running before. Executes inside the remoting scriptblock (works on each remote target).

---

## [2.5.1] - 2026-04-10

### Fixed
- `Remove-CimInstance` silent failure: `Win32_UserProfile::Delete()` removes the `ProfileList` registry key but can silently fail on the folder. Added post-deletion `Test-Path` check with `Remove-Item -Recurse -Force` fallback.

---

## [2.5.0] - 2026-04-10

### Added
- `-RepairDomainDuplicates`: detects `username` / `username.DOMAIN` profile pairs. Repoints the domain SID's `ProfileImagePath` to the local profile path and removes the `.DOMAIN` folder. Safety guard: skipped if local profile is also eligible for deletion.
- `Repaired` status in log and HTML report (cyan badge).

---

## [2.4.1] - 2026-04-10

### Fixed
- `-WhatIf` propagation to `Add-Content`, `Out-File`, `New-Item`: added `-WhatIf:$false` explicitly.
- `.Count` on `Where-Object` results under PS5.1 StrictMode: simplified `-in` syntax returns a non-array single object. Fixed with explicit `Where-Object { $_.Status -eq 'X' -or ... }` and `[array]` type forcing.

---

## [2.4.0] - 2026-04-09

### Added
- **Parallel mode** (`-Parallel`, `-ThrottleLimit`): `ForEach-Object -Parallel` on PS7 for simultaneous multi-server processing. Named mutex for thread-safe log writes. Local stats per runspace aggregated after join. Falls back to sequential on PS5.1 with a warning.

---

## [2.3.0] - 2026-04-09

### Added
- **Windows Event Log** (`-WriteEventLog`, `-EventSource`, `-EventLogName`): EventID `4100` summary (Information/Warning/Error by exit code), EventID `4199` critical error with stack trace. Auto-creates the event source on first run.

---

## [2.2.0] - 2026-04-09

### Added
- **Phase 2** (`-PurgeProfileListBak`): removes orphaned `ProfileList\*.bak` registry keys with SID→username resolution.
- **Phase 3** (`-PurgeBackupFolders`, `-UsersPath`): deletes `*BACKUP*` folders. Active session check via `Win32_UserProfile.Loaded` + `query.exe` fallback. Regex-based base name extraction.

---

## [2.1.0] - 2026-04-08

### Added
- Logging system ported from `Invoke-FilePurge`: `Write-Log` 6 levels, colored console + UTF-8 BOM log file, structured header/footer, per-server separators, automatic log rotation, normalized exit codes.

---

## [2.0.0] - 2026-04-08

### Added
- Initial production release.
- Phase 1: inactive profile purge via `Win32_UserProfile` CIM.
- Inactivity criterion: `Max(folder LastWriteTime, CIM LastUseTime)`.
- System account auto-exclusion, inline and file-based whitelist with wildcard support.
- Local and WinRM remote execution (`-ComputerName`, `-ComputerList`, `-Credential`).
- `-WhatIf` full simulation mode. HTML report (dark theme). Pipeline output.
