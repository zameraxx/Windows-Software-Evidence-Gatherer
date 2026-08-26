# SUPERSEDED by Get-SoftwareEvidence.bat / .ps1 / -Legacy.vbs (see README.md), which
# cover Windows 2000 through current and consolidate output to one CSV + one Summary.txt
# instead of the several CSVs below. Kept here unmodified for anyone with automation
# built against its old multi-file output; Palisade still ingests that output too.
#
# Software inventory evidence - standalone Windows 10 laptop. Run ELEVATED.
# Output: CSVs + Summary.txt on the Desktop. No network calls. No Win32_Product.

$cn  = $env:COMPUTERNAME
$ts  = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
$out = "$env:USERPROFILE\Desktop\SWEvidence_${cn}_$(Get-Date -Format yyyyMMdd-HHmmss)"
New-Item $out -ItemType Directory -Force | Out-Null

$adm = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
       ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $adm) { Write-Warning 'NOT ELEVATED - collection will be incomplete' }

# ---------- 1. system identification ----------
$os = Get-CimInstance Win32_OperatingSystem
$cs = Get-CimInstance Win32_ComputerSystem
$bi = Get-CimInstance Win32_BIOS
$cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$rel = $cv.DisplayVersion; if (-not $rel) { $rel = $cv.ReleaseId }

[pscustomobject]@{
  ComputerName = $cn; Collected = $ts; Elevated = $adm
  Model  = "$($cs.Manufacturer) $($cs.Model)"; Serial = $bi.SerialNumber
  OS     = $os.Caption; Build = "$($os.Version).$($cv.UBR)"; Release = $rel
  Arch   = $os.OSArchitecture; OSInstalled = $os.InstallDate; Workgroup = $cs.Domain
} | Export-Csv "$out\SystemInfo.csv" -NoTypeInformation

# ---------- 2. installed programs, every registry scope ----------
function Get-Un ($key, $scope, $arch, $owner) {
  Get-ChildItem $key -EA 0 | ForEach-Object {
    $p = Get-ItemProperty $_.PSPath -EA 0
    if ($p.DisplayName) {
      [pscustomobject]@{
        ComputerName = $cn; Collected = $ts
        DisplayName  = $p.DisplayName; Version = $p.DisplayVersion; Publisher = $p.Publisher
        InstallDate  = $p.InstallDate; Scope = $scope; Owner = $owner; Arch = $arch
        IsUpdate     = [bool]($p.SystemComponent -eq 1 -or $p.ParentKeyName)
        InstallLocation = $p.InstallLocation; UninstallString = $p.UninstallString; Key = $_.Name
      }
    }
  }
}

$sw  = @()
$sw += Get-Un 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'             'Machine' '64-bit' 'ALL USERS'
$sw += Get-Un 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' 'Machine' '32-bit' 'ALL USERS'

New-PSDrive HKU Registry HKEY_USERS -EA 0 | Out-Null
$loaded = @((Get-ChildItem HKU:\ -EA 0).PSChildName)

foreach ($s in ($loaded | Where-Object { $_ -match '^S-1-5-21-[\d-]+$' })) {
  $sw += Get-Un "HKU:\$s\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"             'User' '64-bit' $s
  $sw += Get-Un "HKU:\$s\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" 'User' '32-bit' $s
}

# profiles that are not logged on - mount their hive, read it, unmount
foreach ($u in (Get-CimInstance Win32_UserProfile -EA 0 |
                Where-Object { -not $_.Special -and $loaded -notcontains $_.SID })) {
  $dat = "$($u.LocalPath)\NTUSER.DAT"
  if (-not (Test-Path $dat)) { continue }
  $t = 'TMP' + ($u.SID -replace '\W')
  reg load "HKU\$t" $dat | Out-Null
  if ($LASTEXITCODE -eq 0) {
    $n = Split-Path $u.LocalPath -Leaf
    $sw += Get-Un "HKU:\$t\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"             'User-offline' '64-bit' $n
    $sw += Get-Un "HKU:\$t\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" 'User-offline' '32-bit' $n
    [gc]::Collect(); Start-Sleep -Milliseconds 250
    reg unload "HKU\$t" | Out-Null
  } else { Write-Warning "Could not mount hive: $dat" }
}

$sw = @($sw | Sort-Object DisplayName, Version, Owner -Unique)
$sw | Export-Csv "$out\InstalledPrograms.csv" -NoTypeInformation
$apps = @($sw | Where-Object { -not $_.IsUpdate })

# ---------- 3. store apps, features, patches ----------
$appx = @(Get-AppxPackage -AllUsers -EA 0 | ForEach-Object {
  [pscustomobject]@{ ComputerName = $cn; Collected = $ts; Name = $_.Name; Version = $_.Version
                     Publisher = $_.Publisher; Arch = $_.Architecture; Location = $_.InstallLocation }
})
$appx | Export-Csv "$out\AppxPackages.csv" -NoTypeInformation

Get-WindowsOptionalFeature -Online -EA 0 |
  Select-Object @{n='ComputerName';e={$cn}}, FeatureName, State |
  Export-Csv "$out\OptionalFeatures.csv" -NoTypeInformation

Get-WindowsCapability -Online -EA 0 | Where-Object State -eq Installed |
  Select-Object @{n='ComputerName';e={$cn}}, Name, State |
  Export-Csv "$out\Capabilities.csv" -NoTypeInformation

$hf = @(Get-HotFix -EA 0 | Sort-Object InstalledOn -Descending |
        Select-Object @{n='ComputerName';e={$cn}}, HotFixID, Description, InstalledOn, InstalledBy)
$hf | Export-Csv "$out\Hotfixes.csv" -NoTypeInformation

# ---------- 4. loose / portable executables (slow) ----------
Write-Host 'Scanning for unregistered executables...'
$ex = @(foreach ($r in "$env:SystemDrive\Users", $env:ProgramData, $env:ProgramFiles, ${env:ProgramFiles(x86)}) {
  if ($r -and (Test-Path $r)) {
    Get-ChildItem $r -Filter *.exe -File -Recurse -Depth 6 -Force -EA 0 | ForEach-Object {
      $sig = Get-AuthenticodeSignature $_.FullName -EA 0
      [pscustomobject]@{
        ComputerName = $cn; Collected = $ts; FileName = $_.Name
        Product = $_.VersionInfo.ProductName; FileVersion = $_.VersionInfo.FileVersion
        Company = $_.VersionInfo.CompanyName; Path = $_.FullName
        Modified = $_.LastWriteTime; Signature = $sig.Status; Signer = $sig.SignerCertificate.Subject
      }
    }
  }
})
$ex | Export-Csv "$out\Executables.csv" -NoTypeInformation

# ---------- 5. summary ----------
$sum = @(
  'SOFTWARE INVENTORY EVIDENCE'
  "Host      : $cn   Serial: $($bi.SerialNumber)"
  "Model     : $($cs.Manufacturer) $($cs.Model)"
  "OS        : $($os.Caption)  $($os.Version).$($cv.UBR)  (version $rel, $($os.OSArchitecture))"
  "Collector : $env:USERDOMAIN\$env:USERNAME"
  "Collected : $ts"
  "Elevated  : $adm"
  'Method    : registry uninstall keys (machine 64/32-bit + all user hives incl.'
  '            offline profiles), Appx, DISM features/capabilities, hotfix list,'
  '            file system scan. No network calls. Win32_Product NOT queried.'
  ''
  "Products                 : $($apps.Count)"
  "Update/component entries : $($sw.Count - $apps.Count)"
  "Store (Appx) packages    : $($appx.Count)"
  "Hotfixes                 : $($hf.Count)"
  "Executables on disk      : $($ex.Count)"
  "Unsigned executables     : $(@($ex | Where-Object { $_.Signature -ne 'Valid' }).Count)"
  ''
)
if (-not $adm) { $sum += '*** COLLECTION INCOMPLETE - NOT RUN ELEVATED ***'; $sum += '' }
$sum += '--- INSTALLED PRODUCTS ---'
$sum += @($apps | Sort-Object DisplayName |
          ForEach-Object { '{0,-55} {1,-18} {2}' -f $_.DisplayName, $_.Version, $_.Publisher })
$sum | Set-Content "$out\Summary.txt" -Encoding UTF8

Write-Host "`nDone. $($apps.Count) products. Output: $out"
$apps | Select-Object DisplayName, Version, Publisher, Owner | Format-Table -AutoSize
