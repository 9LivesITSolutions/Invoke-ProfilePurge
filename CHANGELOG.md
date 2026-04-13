# Changelog

All notable changes to **Invoke-ProfilePurge** are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [2.5.7] - 2026-04-12

### Changed
- **System account exclusion is now SID-based** (language-independent). The previous hardcoded name list (`SYSTEM`, `LOCAL SERVICE`, `NetworkService`, `Administrateur`…) failed on non-English Windows. Replaced by:
  - Exact SID match: `S-1-5-18` (SYSTEM), `S-1-5-19` (LOCAL SERVICE), `S-1-5-20` (NETWORK SERVICE)
  - Suffix match: `*-500` (built-in Administrator in any language), `*-501` (Guest), `*-503` (DefaultAccount)
  - Machine account pattern: SAM name ending with `$` (domain computer accounts)
- `$ExcludeUsers` / `$ExcludeFile` now handle **user-supplied exclusions only**. The system default list is removed.
- Exclusion reason now shows which type triggered: `System SID (S-1-5-18)`, `Machine account (trailing $)`, or `Exclusion rule match`.

---

## [2.5.6] - 2026-04-11

### Fixed
- **Root cause of phantom WhatIf in LIVE mode**: `$WhatIfPreference = 'SilentlyContinue'` assigned a non-empty string (truthy) to a `[bool]` variable, activating WhatIf globally on every run. Changed to `$WhatIfPreference = $false`.
- Added `-WhatIf:$false` to all `Remove-CimInstance` calls in the purge scriptblock to prevent ShouldProcess propagation from parent script scope.

---

## [2.5.5] - 2026-04-11

### Fixed
- **Critical argument binding bug**: `@($DaysInactive, $MergedExclusions, $bool1, $bool2...)` unrolls the string array into the outer array. PowerShell's greedy `[string[]]` parameter binding then consumed all subsequent bool arguments (including `-DeleteUnknownDate`, `-StopWSearch`, `-RepairDomainDuplicates`), which always received `$false`. Fixed by moving `[string[]]$Exclusions` to the **last position** in the scriptblock param block.

---

## [2.5.4] - 2026-04-11

### Added
- `-PassThru` switch: pipeline output is now **silent by default**. Without `-PassThru`, PSCustomObject results no longer flood the interactive console.

---

## [2.5.3] - 2026-04-11

### Added
- `-DeleteUnknownDate` switch: profiles with no `LastWriteTime` and no CIM `LastUseTime` are now **skipped by default** instead of being deleted. Service accounts recreate their profile folder on next startup, causing an infinite delete/recreate loop. Use `-DeleteUnknownDate` to force deletion after manual review.

### Changed
- Console output: `Kept`, `Excluded`, and standard `Skipped` entries no longer print to console. Only actionable entries (`Deleted`, `WhatIf`, `Repaired`, `Error`) and significant skips (domain duplicate session active, unknown date) are shown.

### Redesigned
- HTML report: completely new light-theme design with sticky topbar, KPI strip, chip-based metadata, and filtered tables (non-impacted rows hidden).

---

## [2.5.2] - 2026-04-11

### Added
- `-StopWSearch` switch: stops the Windows Search service (`WSearch`) before the profile purge to release `ntuser.dat.LOG` file handles. Restarts the service after purge **only if it was running before**. Executes inside the remoting scriptblock — works on each target server individually.

---

## [2.5.1] - 2026-04-10

### Fixed
- **`Remove-CimInstance` silent failure**: `Win32_UserProfile::Delete()` can remove the `ProfileList` registry key while silently failing to delete the profile folder (locked files: `ntuser.dat.LOG`, SearchIndexer, AV). Added post-deletion `Test-Path` check with `Remove-Item -Recurse -Force` fallback. If the fallback also fails, status is `Error` instead of a false `Deleted`.

---

## [2.5.0] - 2026-04-10

### Added
- `-RepairDomainDuplicates` switch: auto-detects `username` / `username.DOMAIN` profile pairs. Repoints the domain SID's `ProfileImagePath` registry value to the local profile path and removes the `.DOMAIN` folder. Safety guard: skipped if local profile is also eligible for age-based deletion.
- New `Repaired` status in log and HTML report (cyan badge).
- `DomainDupsFound`, `DomainDupsRepaired`, `DomainDupsWhatIf` counters in global stats.

---

## [2.4.1] - 2026-04-10

### Fixed
- `-WhatIf` propagation to `Add-Content`, `Out-File`, `New-Item`: `[CmdletBinding(SupportsShouldProcess)]` activates `ShouldProcess` at the pipeline level, intercepting these cmdlets even after `$WhatIfPreference` neutralization. Fixed by adding `-WhatIf:$false` explicitly to each internal cmdlet.
- `.Count` on `Where-Object` results in PS5.1 StrictMode: simplified syntax with `-in` returns a single object (not array) when one match is found. Fixed by using explicit `Where-Object { $_.Status -eq 'X' -or ... }` and `[array]` type forcing.
- Replaced `Where-Object Status -in 'A','B'` with explicit OR conditions throughout for PS5.1 compatibility.

---

## [2.4.0] - 2026-04-09

### Added
- **Parallel mode** (`-Parallel`, `-ThrottleLimit`): PS7 `ForEach-Object -Parallel` for simultaneous multi-server processing. Named mutex for thread-safe log writes. Local stats per runspace, aggregated after join. Falls back to sequential on PS5.1 with a warning.
- `-Parallel` auto-ignored for single-server / local mode.

### Changed
- Complete script rewrite for PS5.1 + PS7 dual compatibility.
- All non-ASCII characters removed from PowerShell code strings (UTF-8 BOM retained).

---

## [2.3.0] - 2026-04-09

### Added
- **Windows Event Log** support (`-WriteEventLog`, `-EventSource`, `-EventLogName`).
  - EventID `4100`: summary event (Information / Warning / Error based on exit code).
  - EventID `4199`: critical unhandled error with full stack trace.
  - Auto-creates the event source on first run (requires admin rights once).

---

## [2.2.0] - 2026-04-09

### Added
- **Phase 2** (`-PurgeProfileListBak`): removes orphaned `ProfileList\*.bak` registry keys. SID resolved to username. Detects active twin key.
- **Phase 3** (`-PurgeBackupFolders`, `-UsersPath`): deletes `*BACKUP*` folders. Active session check via `Win32_UserProfile.Loaded` + `query.exe` fallback. Robust base name extraction (regex suffix stripping vs naive `.Split('.')`).

---

## [2.1.0] - 2026-04-08

### Added
- Logging system ported from `Invoke-FilePurge`: `Write-Log` with 6 levels, double console + file output, structured header/footer, per-server separators, automatic log rotation, normalized exit codes.
- `-LogPath`, `-LogRetentionDays` parameters.

---

## [2.0.0] - 2026-04-08

### Added
- Initial production release.
- Phase 1: inactive profile purge via `Win32_UserProfile` CIM (folder + registry key).
- Conservative age criterion: `Max(folder LastWriteTime, CIM LastUseTime)`.
- System account auto-exclusion (`Default*`, `Public`, `NetworkService`, machine accounts `*$`).
- Inline (`-ExcludeUsers`) and file-based (`-ExcludeFile`) whitelist with wildcard support.
- Local and WinRM remote execution (`-ComputerName`, `-ComputerList`, `-Credential`).
- `-WhatIf` full simulation mode.
- HTML report (dark theme, per-server tables, 6 KPI cards).
- Pipeline output of result objects.
