
#Requires -Version 5.1
<#
.SYNOPSIS
    Purge of inactive local user profiles, orphaned registry keys and backup
    folders across a fleet of Windows servers.

.DESCRIPTION
    Three combinable cleanup phases:

    1. INACTIVE PROFILES (always active)
       Deletes profiles inactive for more than N days.
       Conservative criterion: Max(folder LastWriteTime, CIM LastUseTime).
       Clean removal via Win32_UserProfile (folder + ProfileList registry key).

    2. REGISTRY *.BAK KEYS  (-PurgeProfileListBak)
       Removes HKLM:\...\ProfileList\<SID>.bak orphaned keys created by Windows
       when a profile fails to load (corruption, concurrent load, account migration).
       SID is resolved to a username (best-effort).

    3. *BACKUP* FOLDERS  (-PurgeBackupFolders)
       Deletes *BACKUP* folders under -UsersPath (default C:\Users).
       Active session check via Win32_UserProfile.Loaded + query.exe fallback.

    Logging system mirrors Invoke-FilePurge:
      - Write-Log with 6 levels (INFO / WARN / ERROR / SUCCESS / DEBUG / SECTION)
      - Colored console output + timestamped UTF-8 BOM .log file
      - Structured header/footer, per-server and per-phase separators
      - Automatic log rotation, normalized exit codes for task scheduler

    Parallel mode (PS7 only):
      - -Parallel: simultaneous processing of multiple servers
      - Named mutex for thread-safe log file writes
      - Stats aggregated after runspace join
      - Ignored for single-server / local execution

.PARAMETER ComputerName
    Target servers (inline). Absent + no ComputerList -> local execution.

.PARAMETER ComputerList
    Text file listing servers (1/line, # lines ignored).

.PARAMETER DaysInactive
    Inactivity threshold in days for profile purge. Default: 90.

.PARAMETER ExcludeUsers
    Accounts to exclude inline (wildcards OK, e.g. svc_*).

.PARAMETER ExcludeFile
    Whitelist file (1 account/line, wildcards OK).

.PARAMETER PurgeProfileListBak
    Remove orphaned ProfileList\*.bak registry keys.

.PARAMETER PurgeBackupFolders
    Delete *BACKUP* folders under -UsersPath.

.PARAMETER UsersPath
    Profile root path for BACKUP folder search. Default: C:\Users

.PARAMETER LogPath
    Log files destination folder. Default: script folder.

.PARAMETER LogRetentionDays
    Log file retention in days. Default: 30.

.PARAMETER ReportPath
    HTML report path. Default: <LogPath>\ProfilePurge_<timestamp>.html

.PARAMETER Credential
    PSCredential for remote WinRM sessions.

.PARAMETER Parallel
    Enable parallel server processing (PS7+ only).

.PARAMETER ThrottleLimit
    Number of servers processed simultaneously in parallel mode. Default: 5.

.PARAMETER WriteEventLog
    Write a summary event to the Windows Event Log at end of execution.

.PARAMETER EventSource
    Windows event source name. Default: ProfilePurge

.PARAMETER EventLogName
    Target event log. Default: Application

.PARAMETER RepairDomainDuplicates
    Auto-detect and repair username / username.DOMAIN profile pairs.
    Repoints the domain account SID to the local profile path and
    removes the .DOMAIN folder. Skipped if local profile is not viable.

.PARAMETER StopWSearch
    Stop Windows Search (WSearch) before purge to release ntuser.dat.LOG
    file handles. Restarts the service only if it was running before.

.PARAMETER DeleteUnknownDate
    Delete profiles with no LastWriteTime and no CIM LastUseTime.
    By default these profiles are skipped to avoid infinite delete/recreate
    cycles (service accounts recreate their profile on next start).

.PARAMETER PassThru
    Emit result objects to the pipeline. Silent by default.
    Use for: ... | Export-Csv or ... | Where-Object { $_.Status -eq 'Error' }

.PARAMETER WhatIf
    Full simulation -- no deletions performed.

.EXAMPLE
    # Simulate all phases locally
    .\Invoke-ProfilePurge.ps1 -DaysInactive 60 -PurgeProfileListBak -PurgeBackupFolders -WhatIf

.EXAMPLE
    # Live purge across multiple servers with whitelist
    .\Invoke-ProfilePurge.ps1 -ComputerList .\servers.txt -DaysInactive 90 `
        -ExcludeFile .\whitelist.txt -PurgeProfileListBak -PurgeBackupFolders `
        -Parallel -ThrottleLimit 8 -WriteEventLog -LogPath C:\Logs

.EXAMPLE
    # Delete unknown-date profiles (service account orphans) after review
    .\Invoke-ProfilePurge.ps1 -DaysInactive 90 -StopWSearch -DeleteUnknownDate -WhatIf

.EXAMPLE
    # Repair domain duplicates (username vs username.DOMAIN)
    .\Invoke-ProfilePurge.ps1 -DaysInactive 90 -RepairDomainDuplicates -WhatIf

.NOTES
    Requires  : PowerShell 5.1+ (parallel: PS7+)
    Rights    : Local administrator on each target, WinRM enabled
    Exit codes: 0=OK | 1=Critical error | 2=Partial (errors / WinRM failures)
    Event IDs : 4100=Summary | 4199=Critical error
    Author    : 9 Lives IT Solutions
    Version   : 2.5.6
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [string[]]$ComputerName,
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ComputerList,
    [ValidateRange(1, 3650)]
    [int]$DaysInactive = 90,
    [string[]]$ExcludeUsers,
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ExcludeFile,
    [switch]$PurgeProfileListBak,
    [switch]$PurgeBackupFolders,
    [string]$UsersPath = 'C:\Users',
    [string]$LogPath = $PSScriptRoot,
    [ValidateRange(1, 365)]
    [int]$LogRetentionDays = 30,
    [string]$ReportPath,
    [PSCredential]$Credential,
    [switch]$Parallel,
    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 5,
    [switch]$WriteEventLog,
    [string]$EventSource   = 'ProfilePurge',
    [string]$EventLogName  = 'Application',
    [switch]$RepairDomainDuplicates,
    [switch]$StopWSearch,
    [switch]$DeleteUnknownDate,
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Capture -WhatIf BEFORE neutralization.
# SupportsShouldProcess propagates -WhatIf to all cmdlets in scope (CIM aliases included).
# We capture the flag first, then neutralize propagation.
# IMPORTANT: $WhatIfPreference is a [bool], NOT an action preference string.
# 'SilentlyContinue' is a non-empty string = truthy = WhatIf ON. Use $false.
$Script:IsWhatIf  = $PSBoundParameters.ContainsKey('WhatIf')
$WhatIfPreference = $false

$Script:Version   = '2.5.6'
$Script:StartTime = Get-Date
$Script:ExitCode  = 0

# Log initialization
if (-not (Test-Path $LogPath -PathType Container)) {
    New-Item -Path $LogPath -ItemType Directory -Force -WhatIf:$false | Out-Null
}
$LogFile    = Join-Path $LogPath "ProfilePurge_$($Script:StartTime.ToString('yyyyMMdd_HHmmss')).log"
$ReportFile = if ($ReportPath) { $ReportPath } else {
    Join-Path $LogPath "ProfilePurge_$($Script:StartTime.ToString('yyyyMMdd_HHmmss')).html"
}

# Global counters
$GlobalStats = [PSCustomObject]@{
    ProfilesScanned  = [long]0; ProfilesDeleted  = [long]0; ProfilesWhatIf   = [long]0
    ProfilesKept     = [long]0; ProfilesExcluded = [long]0; ProfilesSkipped  = [long]0
    BakKeysFound     = [long]0; BakKeysDeleted   = [long]0; BakKeysWhatIf    = [long]0
    BackupFolders    = [long]0; BackupDeleted    = [long]0; BackupWhatIf     = [long]0; BackupSkipped = [long]0
    DomainDupsFound  = [long]0; DomainDupsRepaired = [long]0; DomainDupsWhatIf = [long]0
    Errors           = [long]0; ServersOK        = [long]0; ServersError     = [long]0
}

# ─────────────────────────────────────────────────────────────────────────────
#region  FONCTIONS UTILITAIRES
# ─────────────────────────────────────────────────────────────────────────────

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG','SECTION')]
        [string]$Level = 'INFO'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $icon = switch ($Level) {
        'INFO'   {'[i]'} 'WARN' {'[!]'} 'ERROR'   {'[X]'}
        'SUCCESS'{'[+]'} 'DEBUG'{'...'} 'SECTION' {'==='}
    }
    $line  = "$ts $icon [$Level] $Message"
    $color = switch ($Level) {
        'INFO'   {'Cyan'   } 'WARN'   {'Yellow' } 'ERROR'   {'Red'    }
        'SUCCESS'{'Green'  } 'DEBUG'  {'Gray'   } 'SECTION' {'Magenta'}
        default  {'White'  }
    }
    Write-Host $line -ForegroundColor $color
    # -WhatIf:$false : CmdletBinding propage ShouldProcess meme apres neutralisation de $WhatIfPreference
    Add-Content -Path $LogFile -Value $line -Encoding UTF8 -WhatIf:$false
}

function Write-Sep {
    param([char]$Char = '=', [int]$Width = 70)
    Write-Log ($Char.ToString() * $Width) -Level SECTION
}

function Write-EventLogEntry {
    param(
        [string]$Message,
        [System.Diagnostics.EventLogEntryType]$EntryType = 'Information',
        [int]$EventId = 4100
    )
    if (-not $WriteEventLog) { return }
    try {
        if (-not [System.Diagnostics.EventLog]::SourceExists($EventSource)) {
            [System.Diagnostics.EventLog]::CreateEventSource($EventSource, $EventLogName)
            Write-Log "Event Log source created: '$EventSource' -> '$EventLogName'" -Level SUCCESS
        }
    }
    catch {
        Write-Log "Cannot create Event Log source '$EventSource': $_ (admin rights required on first run)" -Level WARN
        return
    }
    try {
        Write-EventLog -LogName $EventLogName -Source $EventSource `
                       -EventId $EventId -EntryType $EntryType -Message $Message
        Write-Log "Event Log: EventID $EventId written to '$EventLogName' (source: $EventSource)" -Level DEBUG
    }
    catch { Write-Log "Failed to write to Event Log: $_" -Level WARN }
}

function Invoke-LogRotation {
    $cutoff = (Get-Date).AddDays(-$LogRetentionDays)
    $old = Get-ChildItem -Path $LogPath -Filter 'ProfilePurge_*.log' -ErrorAction SilentlyContinue |
           Where-Object { $_.LastWriteTime -lt $cutoff }
    foreach ($f in $old) {
        try   { Remove-Item $f.FullName -Force; Write-Log "Old log deleted: $($f.Name)" -Level DEBUG }
        catch { Write-Log "Cannot delete old log: $($f.Name)" -Level WARN }
    }
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region  SCRIPTBLOCK PHASE 1 : PROFILS INACTIFS  (self-contained pour remoting)
# ─────────────────────────────────────────────────────────────────────────────

$PurgeProfilesBlock = {
    param([int]$Days, [bool]$DryRun, [bool]$RepairDups, [bool]$DoStopWSearch, [bool]$DelUnknownDate, [string[]]$Exclusions)
    $results = [System.Collections.Generic.List[PSObject]]::new()

    # ── Arret de Windows Search avant la purge ────────────────────────────────
    # WSearch locks ntuser.dat.LOG and prevents Remove-CimInstance from deleting
    # the profile folder. Stop it if running, restart after purge.
    $wSearchWasRunning = $false
    if ($DoStopWSearch -and -not $DryRun) {
        try {
            $svc = Get-Service -Name 'WSearch' -ErrorAction Stop
            if ($svc.Status -eq 'Running') {
                $wSearchWasRunning = $true
                Stop-Service -Name 'WSearch' -Force -ErrorAction Stop
                # Wait for actual stop (max 30s)
                $svc.WaitForStatus('Stopped', [TimeSpan]::FromSeconds(30))
            }
        }
        catch {
            $results.Add([PSCustomObject]@{
                Phase='Profile'; ComputerName=$env:COMPUTERNAME; Identifier='WSearch'
                Detail='Stop-Service WSearch'; LastActivity=$null; DaysInactive=$null
                Loaded=$false; Status='Error'; Reason="Failed to stop WSearch service: $_"
            })
        }
    }
    try {
        $allProfiles = @(Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop |
                         Where-Object { -not $_.Special })
    }
    catch {
        return [PSCustomObject]@{
            Phase='Profile'; ComputerName=$env:COMPUTERNAME; Identifier='N/A'; Detail='N/A'
            LastActivity=$null; DaysInactive=$null; Loaded=$false; Status='Error'
            Reason="Win32_UserProfile not accessible: $_"
        }
    }

    # Pre-pass: index of profiles with no dot in name (local profile candidates)
    # Key = lowercase basename, value = CIM object
    $localIndex = @{}
    foreach ($p in $allProfiles) {
        if ($p.LocalPath) {
            $base = Split-Path $p.LocalPath -Leaf
            if ($base.IndexOf('.') -lt 0) {
                $localIndex[$base.ToLower()] = $p
            }
        }
    }

    foreach ($profile in $allProfiles) {
        $localPath = $profile.LocalPath
        if (-not $localPath) { continue }
        $username = Split-Path $localPath -Leaf

        # Exclusions (wildcards + comptes machine)
        $excluded = ($username -match '\$$')
        if (-not $excluded) {
            foreach ($rule in $Exclusions) { if ($username -like $rule) { $excluded = $true; break } }
        }

        # Domain duplicate detection: name contains a dot AND base name exists in local index
        $isDomainDup     = $false
        $localCounterpart = $null
        if ($RepairDups -and -not $excluded -and $username.IndexOf('.') -ge 0) {
            $baseName = $username.Substring(0, $username.IndexOf('.'))
            if ($localIndex.ContainsKey($baseName.ToLower())) {
                $localCounterpart = $localIndex[$baseName.ToLower()]
                $isDomainDup      = $true
            }
        }

        # Compute this profile's last activity
        $dates = [System.Collections.Generic.List[datetime]]::new()
        if (Test-Path $localPath) { try { $dates.Add((Get-Item $localPath -Force).LastWriteTime) } catch {} }
        if ($profile.LastUseTime) { try { $dates.Add([datetime]$profile.LastUseTime) } catch {} }
        $lastActivity = if ($dates.Count -gt 0) { ($dates | Sort-Object -Descending)[0] } else { $null }
        $daysOld      = if ($lastActivity) { [int]((Get-Date) - $lastActivity).TotalDays } else { $null }

        $status = $reason = ''

        if ($excluded) {
            $status = 'Excluded'; $reason = 'Whitelist / system account'
        }
        elseif ($isDomainDup) {
            # --- Logique reparation doublon domaine ---
            if ($profile.Loaded) {
                $status = 'Skipped'
                $reason = "Domain duplicate -- session active, reparation impossible"
            }
            else {
                # Verification viabilite du profil local cible
                $lDates = [System.Collections.Generic.List[datetime]]::new()
                if ($localCounterpart.LocalPath -and (Test-Path $localCounterpart.LocalPath)) {
                    try { $lDates.Add((Get-Item $localCounterpart.LocalPath -Force).LastWriteTime) } catch {}
                }
                if ($localCounterpart.LastUseTime) {
                    try { $lDates.Add([datetime]$localCounterpart.LastUseTime) } catch {}
                }
                $lLastActivity = if ($lDates.Count -gt 0) { ($lDates | Sort-Object -Descending)[0] } else { $null }
                $lDaysOld      = if ($lLastActivity) { [int]((Get-Date) - $lLastActivity).TotalDays } else { $null }
                # Local profile viable = exists on disk AND not eligible for age-based deletion
                $localViable = (Test-Path $localCounterpart.LocalPath) -and
                               ($null -ne $lLastActivity) -and ($lDaysOld -lt $Days)

                if (-not $localViable) {
                    $status = 'Skipped'
                    $reason = "Domain duplicate -- local profile not viable or also eligible for deletion"
                }
                elseif ($DryRun) {
                    $status = 'WhatIf'
                    $reason = "Domain duplicate -- would repoint to $($localCounterpart.LocalPath) and delete folder"
                }
                else {
                    try {
                        # 1. Repointer l entree SID du compte domaine vers le profil local
                        $sidRegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($profile.SID)"
                        Set-ItemProperty -Path $sidRegPath -Name 'ProfileImagePath' `
                                         -Value $localCounterpart.LocalPath -WhatIf:$false -ErrorAction Stop
                        # 2. Supprimer le dossier .DOMAINE (pas Remove-CimInstance -- on conserve l entree SID)
                        Remove-Item -Path $localPath -Recurse -Force -WhatIf:$false -ErrorAction Stop
                        $status = 'Repaired'
                        $reason = "Repointed SID -> $($localCounterpart.LocalPath) | Domain folder deleted"
                    }
                    catch {
                        $status = 'Error'
                        $reason = "Duplicate repair failed: $($_.Exception.Message)"
                    }
                }
            }
        }
        elseif ($profile.Loaded) {
            $status = 'Skipped'; $reason = 'Active session (profile loaded)'
        }
        elseif ($null -eq $lastActivity) {
            # Date totalement inconnue (LastWriteTime illisible + LastUseTime CIM absent)
            # Ces profils sont souvent des comptes de service recrees automatiquement au
            # demarrage -- les supprimer cree une boucle infinie. Conserves par defaut.
            if ($DelUnknownDate) {
                if ($DryRun) { $status = 'WhatIf'; $reason = 'Would be deleted -- unknown date (-DeleteUnknownDate active)' }
                else {
                    try {
                        $profile | Remove-CimInstance -WhatIf:$false -ErrorAction Stop
                        if (Test-Path $localPath) {
                            Remove-Item -Path $localPath -Recurse -Force -WhatIf:$false -ErrorAction Stop
                        }
                        $status = 'Deleted'; $reason = 'Deleted -- unknown date (-DeleteUnknownDate active)'
                    }
                    catch { $status = 'Error'; $reason = $_.Exception.Message }
                }
            }
            else {
                $status = 'Skipped'
                $reason = 'Unknown date -- skipped (use -DeleteUnknownDate to force)'
            }
        }
        elseif ($daysOld -ge $Days) {
            $age = "${daysOld}j"
            if ($DryRun) { $status = 'WhatIf'; $reason = "Would be deleted -- inactive for $age" }
            else {
                try {
                    $profile | Remove-CimInstance -WhatIf:$false -ErrorAction Stop
                    if (Test-Path $localPath) {
                        Remove-Item -Path $localPath -Recurse -Force -WhatIf:$false -ErrorAction Stop
                    }
                    $status = 'Deleted'; $reason = "Deleted -- inactive for $age"
                }
                catch { $status = 'Error'; $reason = $_.Exception.Message }
            }
        }
        else { $status = 'Kept'; $reason = "Active ${daysOld}d ago" }

        $results.Add([PSCustomObject]@{
            Phase        = 'Profile'
            ComputerName = $env:COMPUTERNAME
            Identifier   = $username
            Detail       = $localPath
            LastActivity = $lastActivity
            DaysInactive = $daysOld
            Loaded       = $profile.Loaded
            Status       = $status
            Reason       = $reason
        })
    }

    # ── Redemarrage de Windows Search si il tournait avant ────────────────────
    if ($wSearchWasRunning) {
        try {
            Start-Service -Name 'WSearch' -ErrorAction Stop
            $results.Add([PSCustomObject]@{
                Phase='Profile'; ComputerName=$env:COMPUTERNAME; Identifier='WSearch'
                Detail='Start-Service WSearch'; LastActivity=$null; DaysInactive=$null
                Loaded=$false; Status='Info'; Reason='WSearch service restarted successfully'
            })
        }
        catch {
            $results.Add([PSCustomObject]@{
                Phase='Profile'; ComputerName=$env:COMPUTERNAME; Identifier='WSearch'
                Detail='Start-Service WSearch'; LastActivity=$null; DaysInactive=$null
                Loaded=$false; Status='Error'; Reason="Failed to restart WSearch service: $_"
            })
        }
    }

    return $results
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region  SCRIPTBLOCK PHASE 2 : CLES REGISTRE *.BAK
# ─────────────────────────────────────────────────────────────────────────────

$PurgeBakBlock = {
    param([bool]$DryRun)
    $results = [System.Collections.Generic.List[PSObject]]::new()
    $regPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    try {
        $bakKeys = Get-ChildItem -Path $regPath -ErrorAction Stop |
                   Where-Object { $_.PSChildName -like '*.bak' }
    }
    catch {
        return [PSCustomObject]@{
            Phase='BakKey'; ComputerName=$env:COMPUTERNAME; Identifier='N/A'; Detail=$regPath
            LastActivity=$null; DaysInactive=$null; Loaded=$false; Status='Error'
            Reason="Registry access failed: $_"
        }
    }
    foreach ($key in $bakKeys) {
        $sid      = $key.PSChildName -replace '\.bak$',''
        $username = $sid
        try { $username = ([System.Security.Principal.SecurityIdentifier]$sid).Translate([System.Security.Principal.NTAccount]).Value } catch {}
        $fullPath    = $key.Name -replace '^HKEY_LOCAL_MACHINE','HKLM:'
        $twinExists  = Test-Path ($fullPath -replace '\.bak$','')
        $profilePath = try { (Get-ItemProperty $fullPath -Name 'ProfileImagePath' -ErrorAction Stop).ProfileImagePath } catch { '(unknown)' }
        $detail = "$fullPath | $profilePath$(if ($twinExists) {' | [!] active twin key present'})"
        $status = $reason = ''
        if ($DryRun) {
            $status='WhatIf'; $reason="Would be deleted -- .bak key$(if ($twinExists) {' (active twin key present)'})"
        }
        else {
            try   { Remove-Item -Path $fullPath -Force -Recurse -ErrorAction Stop; $status='Deleted'; $reason="Deleted .bak key$(if ($twinExists) {' (active twin key preserved)'})" }
            catch { $status='Error'; $reason=$_.Exception.Message }
        }
        $results.Add([PSCustomObject]@{
            Phase='BakKey'; ComputerName=$env:COMPUTERNAME; Identifier=$username; Detail=$detail
            LastActivity=$null; DaysInactive=$null; Loaded=$false; Status=$status; Reason=$reason
        })
    }
    return $results
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region  SCRIPTBLOCK PHASE 3 : DOSSIERS *BACKUP*
# ─────────────────────────────────────────────────────────────────────────────

$PurgeBackupFoldersBlock = {
    param([string]$RootPath, [bool]$DryRun)
    $results = [System.Collections.Generic.List[PSObject]]::new()
    if (-not (Test-Path $RootPath -PathType Container)) {
        return [PSCustomObject]@{
            Phase='BackupFolder'; ComputerName=$env:COMPUTERNAME; Identifier='N/A'; Detail=$RootPath
            LastActivity=$null; DaysInactive=$null; Loaded=$false; Status='Error'
            Reason="Path not found: $RootPath"
        }
    }
    $loadedUsers = @{}
    try {
        Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop |
            Where-Object { $_.Loaded -and $_.LocalPath } |
            ForEach-Object { $loadedUsers[(Split-Path $_.LocalPath -Leaf).ToLower()] = $true }
    } catch {}
    try {
        $queryOut = & query.exe user 2>$null
        if ($queryOut -and $queryOut.Count -gt 1) {
            foreach ($line in ($queryOut | Select-Object -Skip 1)) {
                if ($line -match '^\s*(\S+)') { $loadedUsers[$Matches[1].ToLower()] = $true }
            }
        }
    } catch {}
    try {
        $backupFolders = Get-ChildItem -Path $RootPath -Directory -ErrorAction Stop |
                         Where-Object { $_.Name -like '*BACKUP*' }
    }
    catch {
        return [PSCustomObject]@{
            Phase='BackupFolder'; ComputerName=$env:COMPUTERNAME; Identifier='N/A'; Detail=$RootPath
            LastActivity=$null; DaysInactive=$null; Loaded=$false; Status='Error'
            Reason="Failed to list $RootPath: $_"
        }
    }
    foreach ($folder in $backupFolders) {
        $baseName = ($folder.Name -replace '[._-]*(BACKUP|BAK|OLD)[._-]*\d*$','').TrimEnd('.-_')
        if (-not $baseName) { $baseName = $folder.Name }
        $isLoaded = $loadedUsers.ContainsKey($baseName.ToLower())
        $lastWrite = $folder.LastWriteTime
        $daysOld   = [int]((Get-Date) - $lastWrite).TotalDays
        $status = $reason = ''
        if ($isLoaded) { $status='Skipped'; $reason="Active session detected for '$baseName'" }
        elseif ($DryRun) { $status='WhatIf'; $reason="Would be deleted (${daysOld}d, user: $baseName)" }
        else {
            try   { Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction Stop; $status='Deleted'; $reason="Deleted (${daysOld}d, user: $baseName)" }
            catch { $status='Error'; $reason=$_.Exception.Message }
        }
        $results.Add([PSCustomObject]@{
            Phase='BackupFolder'; ComputerName=$env:COMPUTERNAME; Identifier=$folder.Name; Detail=$folder.FullName
            LastActivity=$lastWrite; DaysInactive=$daysOld; Loaded=$isLoaded; Status=$status; Reason=$reason
        })
    }
    return $results
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region  EXECUTION PRINCIPALE
# ─────────────────────────────────────────────────────────────────────────────

$allResults = [System.Collections.Generic.List[PSObject]]::new()

# Validation mode parallele
$useParallel = $Parallel.IsPresent -and ($PSVersionTable.PSVersion.Major -ge 7)
if ($Parallel.IsPresent -and -not $useParallel) {
    Write-Log "-Parallel requires PS7+. Running sequentially on PS $($PSVersionTable.PSVersion)." -Level WARN
}

try {

    # Build exclusion list
    [string[]]$MergedExclusions = @(
        'Default','Default User','Public','NetworkService',
        'LocalService','systemprofile','SYSTEM','NETWORK SERVICE','LOCAL SERVICE'
    )
    if ($ExcludeUsers) { $MergedExclusions += $ExcludeUsers }
    if ($ExcludeFile) {
        $MergedExclusions += Get-Content $ExcludeFile |
            Where-Object { $_ -notmatch '^\s*#' -and $_ -match '\S' } |
            ForEach-Object { $_.Trim() }
    }

    # Build server list
    [string[]]$Targets = @()
    if ($ComputerName) { $Targets += $ComputerName }
    if ($ComputerList) {
        $Targets += Get-Content $ComputerList |
            Where-Object { $_ -notmatch '^\s*#' -and $_ -match '\S' } |
            ForEach-Object { $_.Trim() }
    }
    $RunLocally = ($Targets.Count -eq 0)
    $TargetStr  = if ($RunLocally) { $env:COMPUTERNAME } else { ($Targets | Select-Object -Unique) -join ', ' }

    # Parallel ne s applique pas en mode local (un seul hote)
    if ($useParallel -and $RunLocally) {
        Write-Log "Local mode detected -- -Parallel ignored (single server)." -Level WARN
        $useParallel = $false
    }

    $phases = @('Inactive profiles')
    if ($PurgeProfileListBak) { $phases += 'Registry *.bak keys' }
    if ($PurgeBackupFolders)  { $phases += "BACKUP folders ($UsersPath)" }

    Invoke-LogRotation

    # HEADER
    Write-Sep '=' 70
    Write-Log "INVOKE-PROFILEPURGE v$($Script:Version)  --  $(if ($Script:IsWhatIf) {'SIMULATION (WhatIf)'} else {'EXECUTION REELLE'})" -Level SECTION
    Write-Sep '=' 70
    Write-Log "Started        : $($Script:StartTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    Write-Log "Targets        : $TargetStr"
    Write-Log "Phases         : $($phases -join ' | ')"
    Write-Log "Execution      : $(if ($useParallel) {"Parallel PS7 (ThrottleLimit: $ThrottleLimit)"} else {'Sequential'})"
    Write-Log "Threshold      : $DaysInactive days inactive"
    Write-Log "Criterion      : Max(folder LastWriteTime, CIM LastUseTime)"
    Write-Log "Exclusions     : $($MergedExclusions.Count) rules loaded"
    Write-Log "Log file       : $LogFile"
    Write-Log "Log retention  : $LogRetentionDays days"
    if ($WriteEventLog) { Write-Log "Event Log      : $EventLogName (source: $EventSource, EventID 4100/4199)" }

    $serverList = if ($RunLocally) { @($env:COMPUTERNAME) } else { @($Targets | Select-Object -Unique) }

    # ─────────────────────────────────────────────────────────────────────────
    # MODE PARALLELE  (PS7 uniquement)
    # ─────────────────────────────────────────────────────────────────────────
    if ($useParallel) {

        Write-Sep '-' 70
        Write-Log "Launching parallel: $($serverList.Count) servers (ThrottleLimit $ThrottleLimit)" -Level INFO

        # Named mutex: thread-safe log file writing
        $logMutex = [System.Threading.Mutex]::new($false)

        # Config passed to runspaces
        $pCfg = [PSCustomObject]@{
            LogFile               = $LogFile
            IsWhatIf              = $Script:IsWhatIf
            DaysInactive          = $DaysInactive
            MergedExclusions      = $MergedExclusions
            Credential            = $Credential
            UsersPath             = $UsersPath
            PurgeProfileListBak   = [bool]$PurgeProfileListBak
            PurgeBackupFolders    = [bool]$PurgeBackupFolders
            RepairDomainDuplicates = [bool]$RepairDomainDuplicates
            StopWSearch           = [bool]$StopWSearch
            DeleteUnknownDate     = [bool]$DeleteUnknownDate
        }
        # Scriptblocks via intermediate variables (PS7.0+ compatibility)
        $sb_Prof = $PurgeProfilesBlock
        $sb_Bak  = $PurgeBakBlock
        $sb_Bkp  = $PurgeBackupFoldersBlock

        $parallelResults = $serverList | ForEach-Object -Parallel {
            $target  = $_
            $c       = $using:pCfg
            $mutex   = $using:logMutex
            $sbProf  = $using:sb_Prof
            $sbBak   = $using:sb_Bak
            $sbBkp   = $using:sb_Bkp

            # Thread-safe Write-Log for this runspace
            # (local function: accesses $c and $mutex via PS scope chain)
            function PLog {
                param([string]$Msg, [string]$Lvl = 'INFO')
                $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                $icon = switch ($Lvl) {
                    'INFO'   {'[i]'} 'WARN' {'[!]'} 'ERROR'   {'[X]'}
                    'SUCCESS'{'[+]'} 'DEBUG'{'...'} 'SECTION' {'==='}
                    default  {'[i]'}
                }
                $line  = "$ts $icon [$Lvl] $Msg"
                $color = switch ($Lvl) {
                    'INFO'   {'Cyan'   } 'WARN'   {'Yellow' } 'ERROR'   {'Red'    }
                    'SUCCESS'{'Green'  } 'DEBUG'  {'Gray'   } 'SECTION' {'Magenta'}
                    default  {'White'  }
                }
                Write-Host $line -ForegroundColor $color
                $null = $mutex.WaitOne(5000)
                try   { Add-Content -Path $c.LogFile -Value $line -Encoding UTF8 }
                finally { $mutex.ReleaseMutex() }
            }

            # Helper: WinRM call for a phase
            function CallPhase {
                param([scriptblock]$Blk, [object[]]$PhaseArgs, [string]$Name)
                PLog ('-' * 70) 'SECTION'
                PLog "PHASE : $Name" 'SECTION'
                PLog ('-' * 70) 'SECTION'
                PLog "WinRM -> $target"
                try {
                    $p = @{ ComputerName=$target; ScriptBlock=$Blk; ArgumentList=$PhaseArgs; ErrorAction='Stop' }
                    if ($c.Credential) { $p['Credential'] = $c.Credential }
                    return @(Invoke-Command @p)
                }
                catch {
                    PLog "WinRM: $_" 'ERROR'
                    return @([PSCustomObject]@{
                        Phase='?'; ComputerName=$target; Identifier='N/A'
                        Detail='N/A'; LastActivity=$null; DaysInactive=$null; Loaded=$false
                        Status='Error'; Reason="WinRM: $_"
                    })
                }
            }

            # Local stats for this runspace
            $ls = @{
                ProfilesScanned=0L; ProfilesDeleted=0L; ProfilesWhatIf=0L
                ProfilesKept=0L;    ProfilesExcluded=0L; ProfilesSkipped=0L
                BakKeysFound=0L;   BakKeysDeleted=0L;   BakKeysWhatIf=0L
                BackupFolders=0L;  BackupDeleted=0L;    BackupWhatIf=0L; BackupSkipped=0L
                Errors=0L
            }
            $serverOK   = $true
            $serverRows = [System.Collections.Generic.List[PSObject]]::new()

            PLog ('=' * 70) 'SECTION'
            PLog "SERVER: $target" 'SECTION'
            PLog ('=' * 70) 'SECTION'

            # Phase 1 : profils inactifs
            $rows1 = CallPhase -Blk $sbProf -PhaseArgs @($c.DaysInactive, $c.IsWhatIf, $c.RepairDomainDuplicates, $c.StopWSearch, $c.DeleteUnknownDate, $c.MergedExclusions) -Name 'Inactive profiles'
            foreach ($r in $rows1) {
                $ls.ProfilesScanned++
                switch ($r.Status) {
                    'Deleted'  { $ls.ProfilesDeleted++;  PLog "  DEL  $($r.Identifier)  --  $($r.Reason)" 'SUCCESS' }
                    'WhatIf'   { $ls.ProfilesWhatIf++;   PLog "  SIM  $($r.Identifier)  --  $($r.Reason)" 'WARN'    }
                    'Kept'     { $ls.ProfilesKept++;      PLog "  OK   $($r.Identifier)  --  $($r.Reason)" 'DEBUG'   }
                    'Excluded' { $ls.ProfilesExcluded++;  PLog "  EXC  $($r.Identifier)  --  $($r.Reason)" 'DEBUG'   }
                    'Skipped'  { $ls.ProfilesSkipped++;   PLog "  SKP  $($r.Identifier)  --  $($r.Reason)"           }
                    'Error'    { $ls.Errors++;             PLog "  ERR  $($r.Identifier)  --  $($r.Reason)" 'ERROR'; $serverOK = $false }
                }
                [void]$serverRows.Add($r)
            }
            $d1 = @($rows1 | Where-Object Status -in 'Deleted','WhatIf').Count
            $e1 = @($rows1 | Where-Object Status -eq 'Error').Count
            PLog "  -> $($rows1.Count) profiles | $d1 purged/simulated | $e1 error(s)" $(if ($e1 -gt 0){'WARN'} elseif ($d1 -gt 0){'INFO'} else{'SUCCESS'})

            # Phase 2 : cles .bak
            if ($c.PurgeProfileListBak) {
                $rows2 = CallPhase -Blk $sbBak -PhaseArgs @($c.IsWhatIf) -Name 'Registry *.bak keys'
                foreach ($r in $rows2) {
                    $ls.BakKeysFound++
                    switch ($r.Status) {
                        'Deleted' { $ls.BakKeysDeleted++; PLog "  DEL  $($r.Identifier)  --  $($r.Reason)" 'SUCCESS' }
                        'WhatIf'  { $ls.BakKeysWhatIf++;  PLog "  SIM  $($r.Identifier)  --  $($r.Reason)" 'WARN'    }
                        'Error'   { $ls.Errors++;           PLog "  ERR  $($r.Identifier)  --  $($r.Reason)" 'ERROR'; $serverOK = $false }
                    }
                    [void]$serverRows.Add($r)
                }
                $d2 = @($rows2 | Where-Object Status -in 'Deleted','WhatIf').Count
                $e2 = @($rows2 | Where-Object Status -eq 'Error').Count
                PLog "  -> $($rows2.Count) .bak keys | $d2 purged/simulated | $e2 error(s)" $(if ($e2 -gt 0){'WARN'} elseif ($d2 -gt 0){'INFO'} else{'SUCCESS'})
            }

            # Phase 3 : dossiers BACKUP
            if ($c.PurgeBackupFolders) {
                $rows3 = CallPhase -Blk $sbBkp -PhaseArgs @($c.UsersPath, $c.IsWhatIf) -Name "BACKUP folders ($($c.UsersPath))"
                foreach ($r in $rows3) {
                    $ls.BackupFolders++
                    switch ($r.Status) {
                        'Deleted' { $ls.BackupDeleted++; PLog "  DEL  $($r.Identifier)  --  $($r.Reason)" 'SUCCESS' }
                        'WhatIf'  { $ls.BackupWhatIf++;  PLog "  SIM  $($r.Identifier)  --  $($r.Reason)" 'WARN'    }
                        'Skipped' { $ls.BackupSkipped++; PLog "  SKP  $($r.Identifier)  --  $($r.Reason)"           }
                        'Error'   { $ls.Errors++;          PLog "  ERR  $($r.Identifier)  --  $($r.Reason)" 'ERROR'; $serverOK = $false }
                    }
                    [void]$serverRows.Add($r)
                }
                $d3 = @($rows3 | Where-Object Status -in 'Deleted','WhatIf').Count
                $e3 = @($rows3 | Where-Object Status -eq 'Error').Count
                PLog "  -> $($rows3.Count) BACKUP folders | $d3 purged/simulated | $e3 error(s)" $(if ($e3 -gt 0){'WARN'} elseif ($d3 -gt 0){'INFO'} else{'SUCCESS'})
            }

            [PSCustomObject]@{ Target=$target; ServerOK=$serverOK; Rows=$serverRows; Stats=$ls }

        } -ThrottleLimit $ThrottleLimit

        $logMutex.Dispose()

        # Agregation des resultats paralleles
        foreach ($sr in @($parallelResults)) {
            if ($sr.ServerOK) { $GlobalStats.ServersOK++ } else { $GlobalStats.ServersError++ }
            $s = $sr.Stats
            $GlobalStats.ProfilesScanned  += $s.ProfilesScanned;  $GlobalStats.ProfilesDeleted  += $s.ProfilesDeleted
            $GlobalStats.ProfilesWhatIf   += $s.ProfilesWhatIf;   $GlobalStats.ProfilesKept     += $s.ProfilesKept
            $GlobalStats.ProfilesExcluded += $s.ProfilesExcluded; $GlobalStats.ProfilesSkipped  += $s.ProfilesSkipped
            $GlobalStats.BakKeysFound     += $s.BakKeysFound;     $GlobalStats.BakKeysDeleted   += $s.BakKeysDeleted
            $GlobalStats.BakKeysWhatIf    += $s.BakKeysWhatIf;    $GlobalStats.BackupFolders    += $s.BackupFolders
            $GlobalStats.BackupDeleted    += $s.BackupDeleted;     $GlobalStats.BackupWhatIf     += $s.BackupWhatIf
            $GlobalStats.BackupSkipped    += $s.BackupSkipped;     $GlobalStats.Errors           += $s.Errors
            if ($sr.Rows) { $allResults.AddRange([PSObject[]]@($sr.Rows)) }
        }
    }

    # ─────────────────────────────────────────────────────────────────────────
    # MODE SEQUENTIEL  (PS5.1 et PS7 sans -Parallel)
    # ─────────────────────────────────────────────────────────────────────────
    else {

        function Invoke-Phase {
            param([scriptblock]$Block, [object[]]$PhaseArgs, [string]$PhaseName, [string]$TargetHost, [bool]$IsLocal)
            Write-Sep '-' 70
            Write-Log "PHASE: $PhaseName" -Level SECTION
            Write-Sep '-' 70
            if ($IsLocal) {
                Write-Log "Local execution"
                try { return @(& $Block @PhaseArgs) }
                catch {
                    Write-Log "Local error: $_" -Level ERROR
                    $Script:ExitCode = [Math]::Max($Script:ExitCode, 2)
                    return @([PSCustomObject]@{
                        Phase='?'; ComputerName=$env:COMPUTERNAME; Identifier='N/A'
                        Detail='N/A'; LastActivity=$null; DaysInactive=$null; Loaded=$false
                        Status='Error'; Reason="Local error: $_"
                    })
                }
            }
            else {
                Write-Log "WinRM -> $TargetHost"
                try {
                    $p = @{ ComputerName=$TargetHost; ScriptBlock=$Block; ArgumentList=$PhaseArgs; ErrorAction='Stop' }
                    if ($Credential) { $p['Credential'] = $Credential }
                    $r = @(Invoke-Command @p)
                    Write-Log "Session opened" -Level SUCCESS
                    return $r
                }
                catch {
                    Write-Log "WinRM: $_" -Level ERROR
                    $Script:ExitCode = [Math]::Max($Script:ExitCode, 2)
                    return @([PSCustomObject]@{
                        Phase='?'; ComputerName=$TargetHost; Identifier='N/A'
                        Detail='N/A'; LastActivity=$null; DaysInactive=$null; Loaded=$false
                        Status='Error'; Reason="WinRM: $_"
                    })
                }
            }
        }

        foreach ($target in $serverList) {

            Write-Sep '=' 70
            Write-Log "SERVER: $target" -Level SECTION
            Write-Sep '=' 70

            $serverOK   = $true
            $serverRows = [System.Collections.Generic.List[PSObject]]::new()

            # Phase 1 : profils inactifs
            # @() force le tableau meme si un seul profil est retourne (StrictMode PS5.1)
            [array]$rows1 = @(Invoke-Phase -Block $PurgeProfilesBlock `
                                  -PhaseArgs @($DaysInactive, $Script:IsWhatIf, $RepairDomainDuplicates.IsPresent, $StopWSearch.IsPresent, $DeleteUnknownDate.IsPresent, $MergedExclusions) `
                                  -PhaseName 'Inactive profiles' -TargetHost $target -IsLocal $RunLocally)
            foreach ($r in $rows1) {
                $GlobalStats.ProfilesScanned++
                switch ($r.Status) {
                    'Deleted'  { $GlobalStats.ProfilesDeleted++;    Write-Log "  DEL  $($r.Identifier)  --  $($r.Reason)" -Level SUCCESS }
                    'WhatIf'   { $GlobalStats.ProfilesWhatIf++;     Write-Log "  SIM  $($r.Identifier)  --  $($r.Reason)" -Level WARN    }
                    'Kept'     { $GlobalStats.ProfilesKept++     }
                    'Excluded' { $GlobalStats.ProfilesExcluded++ }
                    'Skipped'  { $GlobalStats.ProfilesSkipped++;
                                 # Afficher uniquement les Skipped significatifs (session active sur doublon, date inconnue)
                                 if ($r.Reason -match 'Domain duplicate|Unknown date|repair not possible') {
                                     Write-Log "  SKP  $($r.Identifier)  --  $($r.Reason)"
                                 }
                               }
                    'Repaired' { $GlobalStats.DomainDupsRepaired++;  $GlobalStats.DomainDupsFound++
                                 Write-Log "  REP  $($r.Identifier)  --  $($r.Reason)" -Level SUCCESS }
                    'Info'     { Write-Log "  INF  $($r.Identifier)  --  $($r.Reason)" -Level INFO }
                    'Error'    { $GlobalStats.Errors++;               Write-Log "  ERR  $($r.Identifier)  --  $($r.Reason)" -Level ERROR; $serverOK = $false }
                }
                if ($r.Status -eq 'WhatIf' -and $r.Reason -match 'Domain duplicate') {
                    $GlobalStats.DomainDupsFound++; $GlobalStats.DomainDupsWhatIf++
                }
                [void]$serverRows.Add($r)
            }
            # Where-Object syntaxe explicite (evite bug PS5.1 StrictMode avec -in simplifie)
            $p1d = @($rows1 | Where-Object { $_.Status -eq 'Deleted' -or $_.Status -eq 'WhatIf' -or $_.Status -eq 'Repaired' }).Count
            $p1e = @($rows1 | Where-Object { $_.Status -eq 'Error' }).Count
            Write-Log "  -> $($rows1.Count) profiles | $p1d purged/repaired/simulated | $p1e error(s)" `
                      -Level $(if ($p1e -gt 0) {'WARN'} elseif ($p1d -gt 0) {'INFO'} else {'SUCCESS'})

            # Phase 2 : cles .bak
            if ($PurgeProfileListBak) {
                [array]$rows2 = @(Invoke-Phase -Block $PurgeBakBlock -PhaseArgs @($Script:IsWhatIf) `
                                      -PhaseName 'Registry ProfileList *.bak keys' -TargetHost $target -IsLocal $RunLocally)
                foreach ($r in $rows2) {
                    $GlobalStats.BakKeysFound++
                    switch ($r.Status) {
                        'Deleted' { $GlobalStats.BakKeysDeleted++; Write-Log "  DEL  $($r.Identifier)  --  $($r.Reason)" -Level SUCCESS }
                        'WhatIf'  { $GlobalStats.BakKeysWhatIf++;  Write-Log "  SIM  $($r.Identifier)  --  $($r.Reason)" -Level WARN    }
                        'Error'   { $GlobalStats.Errors++;           Write-Log "  ERR  $($r.Identifier)  --  $($r.Reason)" -Level ERROR; $serverOK = $false }
                    }
                    [void]$serverRows.Add($r)
                }
                $p2d = @($rows2 | Where-Object { $_.Status -eq 'Deleted' -or $_.Status -eq 'WhatIf' }).Count
                $p2e = @($rows2 | Where-Object { $_.Status -eq 'Error' }).Count
                Write-Log "  -> $($rows2.Count) .bak keys | $p2d purged/simulated | $p2e error(s)" `
                          -Level $(if ($p2e -gt 0) {'WARN'} elseif ($p2d -gt 0) {'INFO'} else {'SUCCESS'})
            }

            # Phase 3 : dossiers BACKUP
            if ($PurgeBackupFolders) {
                [array]$rows3 = @(Invoke-Phase -Block $PurgeBackupFoldersBlock -PhaseArgs @($UsersPath, $Script:IsWhatIf) `
                                      -PhaseName "BACKUP folders ($UsersPath)" -TargetHost $target -IsLocal $RunLocally)
                foreach ($r in $rows3) {
                    $GlobalStats.BackupFolders++
                    switch ($r.Status) {
                        'Deleted' { $GlobalStats.BackupDeleted++;  Write-Log "  DEL  $($r.Identifier)  --  $($r.Reason)" -Level SUCCESS }
                        'WhatIf'  { $GlobalStats.BackupWhatIf++;   Write-Log "  SIM  $($r.Identifier)  --  $($r.Reason)" -Level WARN    }
                        'Skipped' { $GlobalStats.BackupSkipped++;  Write-Log "  SKP  $($r.Identifier)  --  $($r.Reason)"               }
                        'Error'   { $GlobalStats.Errors++;           Write-Log "  ERR  $($r.Identifier)  --  $($r.Reason)" -Level ERROR; $serverOK = $false }
                    }
                    [void]$serverRows.Add($r)
                }
                $p3d = @($rows3 | Where-Object { $_.Status -eq 'Deleted' -or $_.Status -eq 'WhatIf' }).Count
                $p3e = @($rows3 | Where-Object { $_.Status -eq 'Error' }).Count
                Write-Log "  -> $($rows3.Count) BACKUP folders | $p3d purged/simulated | $p3e error(s)" `
                          -Level $(if ($p3e -gt 0) {'WARN'} elseif ($p3d -gt 0) {'INFO'} else {'SUCCESS'})
            }

            if ($serverOK) { $GlobalStats.ServersOK++ } else { $GlobalStats.ServersError++ }
            $allResults.AddRange($serverRows)
        }
    }

    # ─────────────────────────────────────────────────────────────────────────
    # RESUME FINAL (commun aux deux modes)
    # ─────────────────────────────────────────────────────────────────────────
    $Duration = New-TimeSpan -Start $Script:StartTime -End (Get-Date)
    if ($GlobalStats.Errors -gt 0) { $Script:ExitCode = [Math]::Max($Script:ExitCode, 2) }

    $summaryLines = @(
        "Mode              : $(if ($Script:IsWhatIf) {'SIMULATION (WhatIf) -- no changes made'} else {'Live execution'})"
        "Execution         : $(if ($useParallel) {"Parallel PS7 (ThrottleLimit: $ThrottleLimit)"} else {'Sequential'})"
        "Phases           : $($phases -join ' | ')"
        "Servers OK / KO  : $($GlobalStats.ServersOK) / $($GlobalStats.ServersError)"
        "Duration          : $("{0:hh\:mm\:ss}" -f $Duration)"
        "--- Profiles -------------------------------------------------"
        "  Scanned         : $($GlobalStats.ProfilesScanned)"
        "  Purged/simulated: $($GlobalStats.ProfilesDeleted + $GlobalStats.ProfilesWhatIf)"
        "  Kept            : $($GlobalStats.ProfilesKept)"
        "  Excluded        : $($GlobalStats.ProfilesExcluded)"
        "  Active sessions : $($GlobalStats.ProfilesSkipped)"
        "--- Domain Duplicates ----------------------------------------"
        "  Detected        : $($GlobalStats.DomainDupsFound)"
        "  Repaired/sim.   : $($GlobalStats.DomainDupsRepaired + $GlobalStats.DomainDupsWhatIf)"
        "--- *.bak Registry Keys --------------------------------------"
        "  Found           : $($GlobalStats.BakKeysFound)"
        "  Purged/simulated: $($GlobalStats.BakKeysDeleted + $GlobalStats.BakKeysWhatIf)"
        "--- BACKUP Folders -------------------------------------------"
        "  Found           : $($GlobalStats.BackupFolders)"
        "  Purged/simulated: $($GlobalStats.BackupDeleted + $GlobalStats.BackupWhatIf)"
        "  Skipped (active): $($GlobalStats.BackupSkipped)"
        "--------------------------------------------------------------"
        "Total errors     : $($GlobalStats.Errors)"
        "Exit code        : $Script:ExitCode"
    )

    Write-Sep '=' 70
    Write-Log "SUMMARY -- INVOKE-PROFILEPURGE v$($Script:Version)" -Level SECTION
    Write-Sep '=' 70
    foreach ($line in $summaryLines) {
        if ($line -match '\S') {
            $lvl = if ($line -match 'Total errors.*[^0\s]|Servers OK.*[^0\s/]|SIMULATION') {'WARN'} else {'INFO'}
            Write-Log $line -Level $lvl
        }
    }
    Write-Sep '=' 70
    Write-Log "Full log       : $LogFile"
    Write-Log "HTML report    : $ReportFile"

    # Event Log : synthese finale
    if ($WriteEventLog) {
        $evtBody = "Invoke-ProfilePurge v$($Script:Version)`n$($summaryLines -join "`n")`n`nLog : $LogFile`nHTML : $ReportFile"
        $evtType = switch ($Script:ExitCode) {
            0       { [System.Diagnostics.EventLogEntryType]::Information }
            1       { [System.Diagnostics.EventLogEntryType]::Error       }
            default { [System.Diagnostics.EventLogEntryType]::Warning     }
        }
        Write-EventLogEntry -Message $evtBody -EntryType $evtType -EventId 4100
    }

}
catch {
    Write-Log "CRITICAL ERROR: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack: $($_.ScriptStackTrace)"            -Level ERROR
    $Script:ExitCode = 1
    Write-EventLogEntry -Message "CRITICAL ERROR Invoke-ProfilePurge v$($Script:Version)`n`n$($_.Exception.Message)`n`n$($_.ScriptStackTrace)" `
                        -EntryType Error -EventId 4199
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region  RAPPORT HTML
# ─────────────────────────────────────────────────────────────────────────────

Add-Type -AssemblyName System.Web

# ── Helpers ──────────────────────────────────────────────────────────────────

function Get-StatusBadge {
    param([string]$Status)
    $cfg = @{
        'Deleted'  = @{ bg='#fee2e2'; fg='#ef4444'; label='Deleted'  }
        'WhatIf'   = @{ bg='#fef3c7'; fg='#d97706'; label='WhatIf'   }
        'Repaired' = @{ bg='#cffafe'; fg='#0891b2'; label='Repaired' }
        'Kept'     = @{ bg='#dcfce7'; fg='#16a34a'; label='Kept'     }
        'Excluded' = @{ bg='#ede9fe'; fg='#7c3aed'; label='Excluded' }
        'Skipped'  = @{ bg='#f1f5f9'; fg='#64748b'; label='Skipped'  }
        'Error'    = @{ bg='#fce7f3'; fg='#db2777'; label='Error'    }
        'Info'     = @{ bg='#f1f5f9'; fg='#64748b'; label='Info'     }
    }
    $e = if ($cfg.ContainsKey($Status)) { $cfg[$Status] } else { @{ bg='#f1f5f9'; fg='#64748b'; label=$Status } }
    "<span style='display:inline-block;background:$($e.bg);color:$($e.fg);padding:2px 9px;border-radius:4px;font-size:.72rem;font-family:monospace;font-weight:700;letter-spacing:.03em'>$($e.label)</span>"
}

function Get-PhasePill {
    param([string]$Phase)
    $cfg = @{
        'Profile'      = @{ label='Profile';   color='#3b82f6' }
        'BakKey'       = @{ label='.bak';      color='#8b5cf6' }
        'BackupFolder' = @{ label='Backup FS'; color='#f59e0b' }
    }
    $e = if ($cfg.ContainsKey($Phase)) { $cfg[$Phase] } else { @{ label=$Phase; color='#64748b' } }
    "<span style='display:inline-block;color:$($e.color);font-size:.7rem;font-family:monospace;font-weight:600;letter-spacing:.04em;text-transform:uppercase'>$($e.label)</span>"
}

function Get-SessionDot {
    param([bool]$Loaded, [string]$Phase)
    if ($Phase -ne 'Profile') { return '<span style="color:#334155">--</span>' }
    if ($Loaded) { return '<span style="color:#ef4444" title="Session active">&#9679;</span>' }
    return '<span style="color:#22c55e" title="Pas de session">&#9675;</span>'
}

function Get-TableRows {
    param([PSObject[]]$Rows)
    # Affiche uniquement les lignes actionnables (masque Kept, Excluded, Skipped generiques)
    $visible = @($Rows | Where-Object {
        $_.Status -eq 'Deleted'  -or $_.Status -eq 'WhatIf'  -or
        $_.Status -eq 'Repaired' -or $_.Status -eq 'Error'   -or
        $_.Status -eq 'Info'     -or
        ($_.Status -eq 'Skipped' -and $_.Reason -match 'Doublon|Date inconnue|reparation|active')
    })
    if ($visible.Count -eq 0) {
        return "<tr><td colspan='7' style='text-align:center;padding:1.5rem;color:#475569;font-style:italic'>Aucune action effectuee sur ce serveur</td></tr>"
    }
    $sb = [System.Text.StringBuilder]::new()
    foreach ($r in $visible) {
        $date  = if ($r.LastActivity) { $r.LastActivity.ToString('yyyy-MM-dd HH:mm') } else { '<span style="color:#334155">--</span>' }
        $days  = if ($null -ne $r.DaysInactive) { "$($r.DaysInactive)j" } else { '<span style="color:#334155">--</span>' }
        $badge = Get-StatusBadge  -Status $r.Status
        $phase = Get-PhasePill    -Phase  $r.Phase
        $sess  = Get-SessionDot   -Loaded $r.Loaded -Phase $r.Phase
        $name  = [System.Web.HttpUtility]::HtmlEncode($r.Identifier)
        $reason= [System.Web.HttpUtility]::HtmlEncode($r.Reason)
        [void]$sb.Append("<tr><td>$phase</td><td class='mono hi'>$name</td><td>$badge</td><td class='muted small' style='max-width:320px;word-break:break-word'>$reason</td><td class='muted'>$date</td><td class='center muted'>$days</td><td class='center'>$sess</td></tr>")
    }
    $sb.ToString()
}

# ── Construction du rapport ───────────────────────────────────────────────────

$purgedProfiles = $GlobalStats.ProfilesDeleted + $GlobalStats.ProfilesWhatIf
$purgedBak      = $GlobalStats.BakKeysDeleted  + $GlobalStats.BakKeysWhatIf
$purgedBackup   = $GlobalStats.BackupDeleted   + $GlobalStats.BackupWhatIf
$Duration       = New-TimeSpan -Start $Script:StartTime -End (Get-Date)
$reportDate     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

# KPI cards
function KpiCard { param($v,$l,$c,$s='')
    $sub = if ($s) {"<div class='kpi-sub'>$s</div>"} else {''}
    "<div class='kpi' style='--ac:$c'><div class='kpi-val'>$v</div><div class='kpi-lbl'>$l</div>$sub</div>"
}
$kpis = @(
    (KpiCard $GlobalStats.ProfilesScanned  'Profiles scanned'  '#60a5fa')
    (KpiCard $purgedProfiles               'Purged / simulated' '#ef4444' "kept: $($GlobalStats.ProfilesKept)")
    (KpiCard ($GlobalStats.DomainDupsRepaired + $GlobalStats.DomainDupsWhatIf) 'Domain dups' '#06b6d4' "found: $($GlobalStats.DomainDupsFound)")
    (KpiCard $purgedBak                    '.bak keys'         '#8b5cf6' "found: $($GlobalStats.BakKeysFound)")
    (KpiCard $purgedBackup                 'Backup folders'    '#f59e0b' "skipped: $($GlobalStats.BackupSkipped)")
    (KpiCard $GlobalStats.Errors           'Errors'            '#ec4899')
) -join ''

# Server blocks
$serverBlocks = [System.Text.StringBuilder]::new()
foreach ($computer in ($allResults | Select-Object -ExpandProperty ComputerName -Unique)) {
    $rows   = @($allResults | Where-Object { $_.ComputerName -eq $computer })
    $nDel   = @($rows | Where-Object { $_.Status -eq 'Deleted' -or $_.Status -eq 'WhatIf' -or $_.Status -eq 'Repaired' }).Count
    $nErr   = @($rows | Where-Object { $_.Status -eq 'Error' }).Count
    $dot    = if ($nErr -gt 0) {'#ef4444'} elseif ($nDel -gt 0) {'#f59e0b'} else {'#22c55e'}
    $tableRows = Get-TableRows -Rows $rows
    [void]$serverBlocks.Append("
<div class='server'>
  <div class='server-hd'>
    <span class='server-dot' style='background:$dot'></span>
    <span class='server-name'>$computer</span>
    <span class='server-meta'>$($rows.Count) profiles&nbsp;&nbsp;$nDel actions&nbsp;&nbsp;$nErr errors</span>
  </div>
  <div class='table-wrap'>
    <table>
      <thead><tr><th>Phase</th><th>Account</th><th>Status</th><th>Detail</th><th>Last activity</th><th>Inactive</th><th>Session</th></tr></thead>
      <tbody>$tableRows</tbody>
    </table>
  </div>
</div>")
}

$modeClass = if ($Script:IsWhatIf) {'sim'} else {'real'}
$modeText  = if ($Script:IsWhatIf) {'SIMULATION MODE'} else {'LIVE MODE'}
$exMode    = if ($useParallel) {"Parallel PS7 (x$ThrottleLimit)"} else {'Sequential'}
$phaseTxt  = $phases -join ' &middot; '
$durStr    = '{0:hh\:mm\:ss}' -f $Duration

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>ProfilePurge &mdash; $reportDate</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@300;400;500;600&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --bg:       #f8fafc;
  --surface:  #ffffff;
  --border:   #e2e8f0;
  --text:     #0f172a;
  --muted:    #64748b;
  --mono:     'IBM Plex Mono', monospace;
  --sans:     'Inter', sans-serif;
}

body {
  background: var(--bg);
  color: var(--text);
  font-family: var(--sans);
  font-size: 14px;
  line-height: 1.5;
  min-height: 100vh;
}

/* ── Top bar ── */
.topbar {
  background: #0f172a;
  padding: 0 2rem;
  height: 52px;
  display: flex;
  align-items: center;
  gap: 1rem;
  position: sticky;
  top: 0;
  z-index: 100;
}
.topbar-logo {
  font-family: var(--mono);
  font-size: .82rem;
  font-weight: 600;
  color: #94a3b8;
  letter-spacing: .06em;
}
.topbar-logo span { color: #3b82f6; }
.topbar-spacer { flex: 1; }
.mode-pill {
  font-family: var(--mono);
  font-size: .7rem;
  font-weight: 600;
  padding: .25rem .75rem;
  border-radius: 999px;
  letter-spacing: .06em;
}
.mode-pill.sim  { background: #451a03; color: #fcd34d; border: 1px solid #92400e; }
.mode-pill.real { background: #052e16; color: #4ade80; border: 1px solid #166534; }
.topbar-meta { font-size: .72rem; color: #475569; font-family: var(--mono); }

/* ── Header ── */
.header {
  padding: 2.5rem 2rem 0;
  max-width: 1200px;
  margin: 0 auto;
}
.header-title {
  font-size: 1.4rem;
  font-weight: 600;
  color: #0f172a;
  letter-spacing: -.02em;
}
.header-sub {
  margin-top: .35rem;
  font-size: .82rem;
  color: var(--muted);
  display: flex;
  flex-wrap: wrap;
  gap: .25rem .75rem;
}
.header-sub span::before { content: ''; }
.header-chip {
  display: inline-flex;
  align-items: center;
  gap: .3rem;
  background: #f1f5f9;
  border: 1px solid var(--border);
  border-radius: 4px;
  padding: .15rem .5rem;
  font-size: .72rem;
  font-family: var(--mono);
  color: #475569;
}

/* ── KPI strip ── */
.kpi-strip {
  display: flex;
  flex-wrap: wrap;
  gap: .75rem;
  padding: 1.5rem 2rem;
  max-width: 1200px;
  margin: 0 auto;
}
.kpi {
  flex: 1;
  min-width: 110px;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 1rem 1.25rem;
  border-top: 3px solid var(--ac);
}
.kpi-val {
  font-family: var(--mono);
  font-size: 1.75rem;
  font-weight: 600;
  color: var(--ac);
  line-height: 1;
}
.kpi-lbl {
  margin-top: .3rem;
  font-size: .72rem;
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: .06em;
  color: var(--muted);
}
.kpi-sub {
  margin-top: .2rem;
  font-size: .7rem;
  color: #94a3b8;
  font-family: var(--mono);
}

/* ── Content ── */
.content {
  max-width: 1200px;
  margin: 0 auto;
  padding: 0 2rem 3rem;
  display: flex;
  flex-direction: column;
  gap: 1rem;
}

/* ── Server block ── */
.server {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 8px;
  overflow: hidden;
}
.server-hd {
  display: flex;
  align-items: center;
  gap: .65rem;
  padding: .85rem 1.25rem;
  border-bottom: 1px solid var(--border);
  background: #f8fafc;
}
.server-dot {
  width: 8px;
  height: 8px;
  border-radius: 50%;
  flex-shrink: 0;
}
.server-name {
  font-family: var(--mono);
  font-weight: 600;
  font-size: .88rem;
  color: #0f172a;
}
.server-meta {
  font-size: .75rem;
  color: var(--muted);
  margin-left: auto;
}

/* ── Table ── */
.table-wrap { overflow-x: auto; }
table {
  width: 100%;
  border-collapse: collapse;
  font-size: .81rem;
}
thead tr { background: #f8fafc; }
th {
  padding: .55rem 1rem;
  text-align: left;
  font-family: var(--mono);
  font-size: .67rem;
  font-weight: 600;
  color: #94a3b8;
  text-transform: uppercase;
  letter-spacing: .07em;
  border-bottom: 1px solid var(--border);
  white-space: nowrap;
}
td {
  padding: .55rem 1rem;
  border-bottom: 1px solid #f1f5f9;
  vertical-align: middle;
}
tr:last-child td { border-bottom: none; }
tr:hover td { background: #f8fafc; }
.mono  { font-family: var(--mono); }
.hi    { color: #1e293b; font-weight: 500; }
.muted { color: var(--muted); }
.small { font-size: .75rem; }
.center { text-align: center; }

/* ── Footer ── */
footer {
  text-align: center;
  padding: 2rem;
  font-size: .72rem;
  color: #94a3b8;
  font-family: var(--mono);
  border-top: 1px solid var(--border);
}

@media (max-width: 768px) {
  .kpi-strip { gap: .5rem; }
  .kpi { min-width: 90px; }
}
@media print {
  .topbar { position: static; }
  .server { break-inside: avoid; }
}
</style>
</head>
<body>

<nav class="topbar">
  <div class="topbar-logo"><span>//</span> PROFILEPURGE</div>
  <div class="topbar-spacer"></div>
  <span class="topbar-meta">$reportDate &nbsp;&middot;&nbsp; $durStr &nbsp;&middot;&nbsp; exit $Script:ExitCode</span>
  <span class="mode-pill $modeClass">$modeText</span>
</nav>

<div class="header">
  <div class="header-title">Profile Purge Report</div>
  <div class="header-sub">
    <span class="header-chip">targets: $TargetStr</span>
    <span class="header-chip">phases: $phaseTxt</span>
    <span class="header-chip">threshold: ${DaysInactive}d inactive</span>
    <span class="header-chip">exec: $exMode</span>
    <span class="header-chip">log: $LogFile</span>
  </div>
</div>

<div class="kpi-strip">$kpis</div>

<div class="content">$serverBlocks</div>

<footer>9 Lives IT Solutions &nbsp;&middot;&nbsp; Invoke-ProfilePurge.ps1 v$($Script:Version) &nbsp;&middot;&nbsp; $reportDate</footer>
</body>
</html>
"@

try {
    $html | Out-File -FilePath $ReportFile -Encoding UTF8 -Force -WhatIf:$false
    Write-Log "HTML report generated : $ReportFile" -Level SUCCESS
}
catch { Write-Log "Failed to write HTML report : $_" -Level WARN }

#endregion

# Pipeline output: silent by default, enabled with -PassThru
# (prevents all PSCustomObject results from printing in interactive console)
if ($PassThru) {
    $allResults | Sort-Object ComputerName, Phase, Status, Identifier
}
exit $Script:ExitCode
