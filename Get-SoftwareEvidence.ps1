<#
.SYNOPSIS
  Software inventory evidence collector - PowerShell path.
  Covers Windows Vista / Server 2008 through the current release, and Windows XP
  SP2+ / Server 2003 SP1+ if PowerShell has been installed on them. Written against
  PowerShell 2.0 syntax (no CIM cmdlets, no [pscustomobject], no Get-ChildItem -Depth)
  so it runs unmodified on every PowerShell version from 2.0 through 7.x.

  Windows 2000 cannot run any version of PowerShell - use the companion
  Get-SoftwareEvidence-Legacy.vbs on that OS (or on XP/2003 boxes that never had
  PowerShell installed). Get-SoftwareEvidence.bat picks whichever of the two is
  correct for the machine it's run on automatically; run this script directly only
  if you already know PowerShell is present.

  Run ELEVATED (right-click -> Run as administrator, or an elevated PowerShell
  prompt) for a complete collection. It still runs unelevated, but the offline
  user-hive and some registry/WMI reads will be skipped or partial, and both the
  console output and Summary.txt will say so plainly.

.OUTPUTS
  One timestamped folder on the current user's Desktop containing exactly two files:
    SoftwareEvidence_<host>_<timestamp>.csv  - every record type in one file, keyed
                                                by a RecordType column (see README.md)
    Summary_<host>_<timestamp>.txt           - human-readable summary

  No network calls of any kind. Win32_Product is never queried (it silently
  triggers an MSI self-repair on every installed package).
#>
param(
  # Skip the loose-executable file system scan (section 5 below) - the one slow
  # step, worth skipping on a machine with a large or slow disk. Mirrors /noexe
  # on Get-SoftwareEvidence-Legacy.vbs; Get-SoftwareEvidence.bat forwards
  # whichever one applies automatically.
  [switch]$NoExe
)

# ============================================================
# 0. setup
# ============================================================
$cn    = $env:COMPUTERNAME
$now   = Get-Date
$ts    = $now.ToString('yyyy-MM-dd HH:mm:ss')
$stamp = $now.ToString('yyyyMMdd-HHmmss')
$outDir  = Join-Path $env:USERPROFILE "Desktop\SWEvidence_${cn}_$stamp"
New-Item $outDir -ItemType Directory -Force | Out-Null
$csvPath = Join-Path $outDir "SoftwareEvidence_${cn}_$stamp.csv"
$sumPath = Join-Path $outDir "Summary_${cn}_$stamp.txt"

$adm = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
       ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $adm) { Write-Warning 'NOT ELEVATED - collection will be incomplete' }

$psVer = '1.0 (or unrecognized)'
if (Test-Path Variable:PSVersionTable) { $psVer = $PSVersionTable.PSVersion.ToString() }

# Fixed column order shared with Get-SoftwareEvidence-Legacy.vbs so both collectors -
# whichever one a given machine needed - produce an identical schema. Palisade (and
# anything else reading this) keys off RecordType to know what the rest of the row means.
$AllColumns = @(
  'RecordType','ComputerName','Collected','Elevated','Name','Version','Publisher',
  'InstallDate','Scope','Owner','Arch','IsUpdate','InstallLocation','UninstallString',
  'RegistryKey','Path','Modified','Signature','Signer','State',
  'Model','Serial','OSCaption','OSVersion','OSBuild','OSRelease','OSArchitecture',
  'Domain','CollectorMethod','Notes'
)
$AllRows = New-Object System.Collections.ArrayList

# [pscustomobject] and Add-Member -NotePropertyName/-Value are PowerShell 3.0+ only;
# this is the PSv2-safe way to build an object with a guaranteed column order.
function New-Row {
  param([hashtable]$Fields)
  $o = New-Object PSObject
  foreach ($col in $AllColumns) {
    $v = $null
    if ($Fields.ContainsKey($col)) { $v = $Fields[$col] }
    $o | Add-Member -MemberType NoteProperty -Name $col -Value $v
  }
  [void]$AllRows.Add($o)
  return $o
}

try {

# ============================================================
# 1. system identification
# ============================================================
$os = Get-WmiObject Win32_OperatingSystem -EA 0
$cs = Get-WmiObject Win32_ComputerSystem  -EA 0
$bi = Get-WmiObject Win32_BIOS            -EA 0
$cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -EA 0

$osCaption = $os.Caption
if ($os.CSDVersion) { $osCaption = "$osCaption ($($os.CSDVersion))" }

$osBuild = $os.Version
if ($cv.UBR) { $osBuild = "$($os.Version).$($cv.UBR)" }   # UBR only exists on Windows 10+

$osRelease = $cv.DisplayVersion                            # Windows 10 2009+
if (-not $osRelease) { $osRelease = $cv.ReleaseId }         # Windows 10 pre-2009
if (-not $osRelease) { $osRelease = $cv.CurrentVersion }    # XP/2003/Vista/7/8.x fallback ("6.1" etc.)

$osArch = $os.OSArchitecture   # not present on XP/2003/2000 - Win32_OperatingSystem gained it in Vista
if (-not $osArch) {
  $osArch = if (Test-Path 'HKLM:\SOFTWARE\WOW6432Node') { '64-bit' } else { '32-bit' }
}

# Feature availability drives which optional sections below actually run. Checking
# for the cmdlet itself (rather than hardcoding a version-number cutoff) keeps this
# correct on Server SKUs that omit a feature client SKUs have, and on whatever
# ships after this script was written.
$hasAppx       = [bool](Get-Command Get-AppxPackage           -EA SilentlyContinue)
$hasOptFeature = [bool](Get-Command Get-WindowsOptionalFeature -EA SilentlyContinue)
$hasCapability = [bool](Get-Command Get-WindowsCapability      -EA SilentlyContinue)

$collectorNote = "PowerShell $psVer / WMI (Get-WmiObject)"
if ($hasAppx)       { $collectorNote += ' + Appx' }
if ($hasOptFeature) { $collectorNote += ' + OptionalFeature' }
if ($hasCapability) { $collectorNote += ' + Capability' }

$sysNotes = ''
if (-not $adm) { $sysNotes = 'NOT RUN ELEVATED - collection incomplete' }

New-Row @{
  RecordType = 'SystemInfo'; ComputerName = $cn; Collected = $ts; Elevated = $adm
  Model = "$($cs.Manufacturer) $($cs.Model)".Trim(); Serial = $bi.SerialNumber
  OSCaption = $osCaption; OSVersion = $os.Version; OSBuild = $osBuild; OSRelease = $osRelease
  OSArchitecture = $osArch; Domain = $cs.Domain
  CollectorMethod = $collectorNote; Notes = $sysNotes
} | Out-Null

# ============================================================
# 2. installed programs, every registry scope
# ============================================================
function Get-Un ($key, $scope, $arch, $owner) {
  Get-ChildItem $key -EA 0 | ForEach-Object {
    $p = Get-ItemProperty $_.PSPath -EA 0
    if ($p.DisplayName) {
      New-Object PSObject -Property @{
        Name = $p.DisplayName; Version = $p.DisplayVersion; Publisher = $p.Publisher
        InstallDate = $p.InstallDate; Scope = $scope; Owner = $owner; Arch = $arch
        IsUpdate = [bool]($p.SystemComponent -eq 1 -or $p.ParentKeyName)
        InstallLocation = $p.InstallLocation; UninstallString = $p.UninstallString
        RegistryKey = $_.Name
      }
    }
  }
}

$sw = @()
$sw += Get-Un 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' 'Machine' '64-bit' 'ALL USERS'
if (Test-Path 'HKLM:\SOFTWARE\WOW6432Node') {
  $sw += Get-Un 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' 'Machine' '32-bit' 'ALL USERS'
}

New-PSDrive HKU Registry HKEY_USERS -EA 0 | Out-Null
$loaded = @((Get-ChildItem HKU:\ -EA 0).PSChildName)

foreach ($s in ($loaded | Where-Object { $_ -match '^S-1-5-21-[\d-]+$' })) {
  $sw += Get-Un "HKU:\$s\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" 'User' '64-bit' $s
  if (Test-Path "HKU:\$s\SOFTWARE\WOW6432Node") {
    $sw += Get-Un "HKU:\$s\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" 'User' '32-bit' $s
  }
}

# Profiles that are not logged on: mount their hive, read it, unmount. Uses the
# ProfileList registry key (present on every version back to Windows 2000) rather
# than the Win32_UserProfile WMI class, which does not exist before Vista.
$profileListKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
Get-ChildItem $profileListKey -EA 0 | ForEach-Object {
  $sid = $_.PSChildName
  if ($sid -notmatch '^S-1-5-21-[\d-]+$' -or $loaded -contains $sid) { return }
  $pl = Get-ItemProperty $_.PSPath -EA 0
  if (-not $pl.ProfileImagePath) { return }
  $dat = Join-Path $pl.ProfileImagePath 'NTUSER.DAT'
  if (-not (Test-Path $dat)) { return }
  $t = 'TMP' + ($sid -replace '\W')
  reg load "HKU\$t" $dat 2>$null | Out-Null
  if ($LASTEXITCODE -eq 0) {
    $n = Split-Path $pl.ProfileImagePath -Leaf
    $sw += Get-Un "HKU:\$t\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" 'User-offline' '64-bit' $n
    if (Test-Path "HKU:\$t\SOFTWARE\WOW6432Node") {
      $sw += Get-Un "HKU:\$t\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" 'User-offline' '32-bit' $n
    }
    [gc]::Collect(); Start-Sleep -Milliseconds 250
    reg unload "HKU\$t" 2>$null | Out-Null
  } else { Write-Warning "Could not mount hive: $dat" }
}

$sw = @($sw | Sort-Object Name, Version, Owner -Unique)
foreach ($p in $sw) {
  New-Row @{
    RecordType = 'Program'; ComputerName = $cn; Collected = $ts
    Name = $p.Name; Version = $p.Version; Publisher = $p.Publisher
    InstallDate = $p.InstallDate; Scope = $p.Scope; Owner = $p.Owner; Arch = $p.Arch
    IsUpdate = $p.IsUpdate; InstallLocation = $p.InstallLocation
    UninstallString = $p.UninstallString; RegistryKey = $p.RegistryKey
  } | Out-Null
}
$apps = @($sw | Where-Object { -not $_.IsUpdate })

# ============================================================
# 3. store apps, features, capabilities (Windows 8 / Server 2012 and later only -
#    the cmdlets simply do not exist on older Windows, so these are skipped there)
# ============================================================
$appxCount = 0
if ($hasAppx) {
  try {
    Get-AppxPackage -AllUsers -EA 0 | ForEach-Object {
      New-Row @{
        RecordType = 'AppxPackage'; ComputerName = $cn; Collected = $ts
        Name = $_.Name; Version = $_.Version; Publisher = $_.Publisher
        Arch = $_.Architecture; Path = $_.InstallLocation
      } | Out-Null
      $appxCount++
    }
  } catch { Write-Warning "Appx enumeration failed: $($_.Exception.Message)" }
}

$featCount = 0
if ($hasOptFeature) {
  try {
    Get-WindowsOptionalFeature -Online -EA 0 | Where-Object { $_.State -eq 'Enabled' } | ForEach-Object {
      New-Row @{ RecordType='OptionalFeature'; ComputerName=$cn; Collected=$ts; Name=$_.FeatureName; State=$_.State } | Out-Null
      $featCount++
    }
  } catch { Write-Warning "Optional feature enumeration failed: $($_.Exception.Message)" }
}

$capCount = 0
if ($hasCapability) {
  try {
    Get-WindowsCapability -Online -EA 0 | Where-Object { $_.State -eq 'Installed' } | ForEach-Object {
      New-Row @{ RecordType='Capability'; ComputerName=$cn; Collected=$ts; Name=$_.Name; State=$_.State } | Out-Null
      $capCount++
    }
  } catch { Write-Warning "Capability enumeration failed: $($_.Exception.Message)" }
}

# ============================================================
# 4. hotfixes - Win32_QuickFixEngineering exists on every version this script
#    targets (Windows XP / Server 2003 and later)
# ============================================================
$hfCount = 0
try {
  Get-HotFix -EA 0 | Sort-Object InstalledOn -Descending | ForEach-Object {
    New-Row @{
      RecordType = 'Hotfix'; ComputerName = $cn; Collected = $ts
      Name = $_.HotFixID; Notes = $_.Description; InstallDate = $_.InstalledOn; Owner = $_.InstalledBy
    } | Out-Null
    $hfCount++
  }
} catch { Write-Warning "Hotfix enumeration failed: $($_.Exception.Message)" }

# ============================================================
# 5. loose / portable executables (slow - this is the one step worth skipping on
#    old or slow disks; run with -NoExe to skip it)
# ============================================================
$exCount = 0
$unsignedCount = 0
if ($NoExe) {
  Write-Host 'Skipping executable scan (-NoExe).'
} else {
Write-Host 'Scanning for unregistered executables...'

# Get-ChildItem -Depth needs PowerShell 5.0+; this manual walk is PSv2-safe and
# enforces the same depth cap explicitly.
function Get-ExeFilesBounded ($rootPath, $maxDepth) {
  $found = New-Object System.Collections.ArrayList
  function Walk ($path, $depth) {
    if ($depth -gt $maxDepth) { return }
    $items = Get-ChildItem -LiteralPath $path -Force -EA 0
    foreach ($it in $items) {
      if ($it.PSIsContainer) { Walk $it.FullName ($depth + 1) }
      elseif ($it.Extension -ieq '.exe') { [void]$found.Add($it) }
    }
  }
  Walk $rootPath 0
  return $found
}

$roots = @(
  "$env:SystemDrive\Users", "$env:SystemDrive\Documents and Settings",   # XP/2003-style profile root
  $env:ProgramData, $env:ProgramFiles, ${env:ProgramFiles(x86)}
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

foreach ($r in $roots) {
  foreach ($f in (Get-ExeFilesBounded $r 6)) {
    $sig = Get-AuthenticodeSignature $f.FullName -EA 0
    $signer = ''
    if ($sig.SignerCertificate) { $signer = $sig.SignerCertificate.Subject }
    New-Row @{
      RecordType = 'Executable'; ComputerName = $cn; Collected = $ts
      Name = $f.Name; Version = $f.VersionInfo.FileVersion; Publisher = $f.VersionInfo.CompanyName
      Path = $f.FullName; Modified = $f.LastWriteTime; Signature = $sig.Status; Signer = $signer
    } | Out-Null
    $exCount++
    if ($sig.Status -ne 'Valid') { $unsignedCount++ }
  }
}
}

} catch {
  Write-Warning "Collector error: $($_.Exception.Message)"
  Write-Warning 'Writing whatever was collected before the error, so partial evidence is not lost.'
} finally {

# ============================================================
# 6. export - exactly two files
# ============================================================

# Windows PowerShell's default Export-Csv encoding is plain ASCII and silently
# mangles anything outside it (accented publisher names, "µTorrent", curly quotes
# in game titles, etc.) into "?". Force UTF-8 so the evidence is byte-accurate.
$AllRows | Select-Object $AllColumns | Export-Csv $csvPath -NoTypeInformation -Encoding UTF8

$sum = New-Object System.Collections.ArrayList
[void]$sum.Add('SOFTWARE INVENTORY EVIDENCE')
[void]$sum.Add("Host      : $cn   Serial: $($bi.SerialNumber)")
[void]$sum.Add("Model     : $($cs.Manufacturer) $($cs.Model)".Trim())
[void]$sum.Add("OS        : $osCaption  $osBuild  (version $osRelease, $osArch)")
[void]$sum.Add("Collector : $env:USERDOMAIN\$env:USERNAME")
[void]$sum.Add("Collected : $ts")
[void]$sum.Add("Elevated  : $adm")
[void]$sum.Add("Method    : $collectorNote; registry uninstall keys (machine 64/32-bit + all user")
[void]$sum.Add("            hives incl. offline profiles), hotfix list$(if($NoExe){'.'}else{', file system scan.'})")
[void]$sum.Add('            No network calls. Win32_Product NOT queried.')
[void]$sum.Add('')
[void]$sum.Add("Products                 : $($apps.Count)")
[void]$sum.Add("Update/component entries : $($sw.Count - $apps.Count)")
[void]$sum.Add("Store (Appx) packages    : $appxCount$(if(-not $hasAppx){' (not applicable on this OS)'})")
[void]$sum.Add("Optional features on     : $featCount$(if(-not $hasOptFeature){' (not applicable on this OS)'})")
[void]$sum.Add("Capabilities installed   : $capCount$(if(-not $hasCapability){' (not applicable on this OS)'})")
[void]$sum.Add("Hotfixes                 : $hfCount")
[void]$sum.Add("Executables on disk      : $exCount$(if($NoExe){' (scan skipped: -NoExe)'})")
[void]$sum.Add("Unsigned executables     : $unsignedCount")
[void]$sum.Add('')
if (-not $adm) { [void]$sum.Add('*** COLLECTION INCOMPLETE - NOT RUN ELEVATED ***'); [void]$sum.Add('') }
[void]$sum.Add('--- INSTALLED PRODUCTS ---')
$apps | Sort-Object Name | ForEach-Object {
  [void]$sum.Add(('{0,-55} {1,-18} {2}' -f $_.Name, $_.Version, $_.Publisher))
}
$sum | Set-Content $sumPath -Encoding UTF8

Write-Host "`nDone. $($apps.Count) products. Output: $outDir"
$apps | Select-Object Name, Version, Publisher, Owner | Format-Table -AutoSize

}
