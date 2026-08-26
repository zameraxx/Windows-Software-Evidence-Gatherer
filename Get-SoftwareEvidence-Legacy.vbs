' ================================================================
' Software inventory evidence collector - legacy (no PowerShell) path.
'
' For Windows 2000, and any Windows XP / Server 2003 box that never had
' PowerShell installed on it. Windows 2000 cannot run PowerShell at all - no
' version of PowerShell ever supported it - so this VBScript, running under
' cscript.exe (part of Windows Script Host, present by default since
' Windows 98), is the only zero-install way to collect evidence from it.
' Get-SoftwareEvidence.bat picks this automatically when powershell.exe is
' not found on the machine; run it directly with cscript only if you already
' know that's the case.
'
' Run ELEVATED (right-click -> Run as administrator on XP+, or logged on as
' an Administrator on Windows 2000, which has no UAC) for a complete
' collection. It still runs unelevated, but the collection will be partial,
' and both console output and Summary.txt say so plainly.
'
' Usage:  cscript //nologo Get-SoftwareEvidence-Legacy.vbs [/noexe]
'   /noexe  skip the loose-executable file system scan (the one slow step -
'           worth skipping on an old machine with a large, slow disk)
'
' Output: one timestamped folder on the current user's Desktop containing
' exactly two files, in the SAME column layout Get-SoftwareEvidence.ps1
' produces, so a mixed fleet of old and new machines all feed one format:
'   SoftwareEvidence_<host>_<timestamp>.csv
'   Summary_<host>_<timestamp>.txt
'
' No network calls of any kind. Win32_Product is never queried.
'
' Known gaps versus the PowerShell collector (documented, not silent):
'   - Offline (not logged on) user profiles cannot be enumerated: doing so
'     needs reg.exe to load/unload NTUSER.DAT, and Windows 2000 does not
'     ship reg.exe by default. Only currently logged-on users' hives
'     (visible under HKEY_USERS) are read.
'   - Authenticode signature status is not checked for executables found on
'     disk (WinVerifyTrust is not reachable from plain VBScript without an
'     external helper); the Signature column reads "N/A (legacy collector)".
'   - Store apps / optional features / Windows capabilities do not exist as
'     concepts before Windows 8, so those RecordTypes simply have zero rows.
' ================================================================

Option Explicit

Dim fso, shell, wmi, reg
Dim cn, tsDisplay, tsStamp, outDir, csvPath, sumPath, csvStream, sumStream
Dim adm, skipExe, arg
Dim model, serial, domain, osCaption, osVersion, osBuild, osRelease, osArch, csdVer
Dim collectorNote
Dim appCount, updCount, hfCount, exCount
Dim cols(29)
Dim seenProg, prodNames(), prodLines(), prodCount

Const HKLM        = &H80000002
Const HKEY_USERS  = &H80000003

' ---------------- 0. setup ----------------
skipExe = False
For Each arg In WScript.Arguments
  If LCase(arg) = "/noexe" Then skipExe = True
Next

Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")
Set wmi   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set reg   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

cn        = shell.ExpandEnvironmentStrings("%COMPUTERNAME%")
tsDisplay = FormatTimestamp(Now)
tsStamp   = FormatStamp(Now)

outDir = shell.SpecialFolders("Desktop") & "\SWEvidence_" & cn & "_" & tsStamp
If Not fso.FolderExists(outDir) Then fso.CreateFolder(outDir)
csvPath = outDir & "\SoftwareEvidence_" & cn & "_" & tsStamp & ".csv"
sumPath = outDir & "\Summary_" & cn & "_" & tsStamp & ".txt"

adm = IsElevated()
If Not adm Then WScript.Echo "WARNING: NOT ELEVATED - collection will be incomplete"

' Column order MUST match $AllColumns in Get-SoftwareEvidence.ps1 exactly -
' this is what lets both collectors' output be treated as one format downstream.
cols(0)="RecordType":       cols(1)="ComputerName":     cols(2)="Collected"
cols(3)="Elevated":         cols(4)="Name":              cols(5)="Version"
cols(6)="Publisher":        cols(7)="InstallDate":       cols(8)="Scope"
cols(9)="Owner":            cols(10)="Arch":             cols(11)="IsUpdate"
cols(12)="InstallLocation": cols(13)="UninstallString":  cols(14)="RegistryKey"
cols(15)="Path":            cols(16)="Modified":         cols(17)="Signature"
cols(18)="Signer":          cols(19)="State":            cols(20)="Model"
cols(21)="Serial":          cols(22)="OSCaption":        cols(23)="OSVersion"
cols(24)="OSBuild":         cols(25)="OSRelease":        cols(26)="OSArchitecture"
cols(27)="Domain":          cols(28)="CollectorMethod":  cols(29)="Notes"

appCount = 0 : updCount = 0 : hfCount = 0 : exCount = 0
prodCount = 0
ReDim prodNames(63)
ReDim prodLines(63)
Set seenProg = CreateObject("Scripting.Dictionary")

Set csvStream = fso.CreateTextFile(csvPath, True, False)   ' overwrite, ANSI (see WriteRow note)
WriteCsvHeader csvStream

' ---------------- 1. system identification ----------------
Dim osRow, csRow, biRow
For Each osRow In wmi.ExecQuery("SELECT Caption,Version,CSDVersion FROM Win32_OperatingSystem")
  osCaption = osRow.Caption
  osVersion = osRow.Version
  csdVer    = osRow.CSDVersion
Next
On Error Resume Next
osArch = ""
osArch = osRow.OSArchitecture   ' property does not exist before Vista - guarded
If Err.Number <> 0 Then osArch = "" : Err.Clear
On Error Goto 0

' WMI leaves an unset string property as Null, not "" - guard every one before
' it is used in Len()/comparison rather than only where it is merely stored.
If IsNull(osCaption) Then osCaption = ""
If IsNull(osVersion) Then osVersion = ""
If IsNull(csdVer)    Then csdVer    = ""
If Len(csdVer) > 0 Then osCaption = osCaption & " (" & csdVer & ")"
osBuild   = osVersion
osRelease = ""   ' DisplayVersion/ReleaseId/UBR do not exist before Windows 10
If Len(osArch) = 0 Then
  If fso.FolderExists(shell.ExpandEnvironmentStrings("%SystemRoot%") & "\SysWOW64") Then
    osArch = "64-bit"
  Else
    osArch = "32-bit"
  End If
End If

For Each csRow In wmi.ExecQuery("SELECT Manufacturer,Model,Domain FROM Win32_ComputerSystem")
  model  = Trim(csRow.Manufacturer & " " & csRow.Model)   ' "&" coalesces Null to "" on its own
  domain = csRow.Domain
Next
If IsNull(domain) Then domain = ""
For Each biRow In wmi.ExecQuery("SELECT SerialNumber FROM Win32_BIOS")
  serial = biRow.SerialNumber
Next
If IsNull(serial) Then serial = ""

collectorNote = "VBScript legacy collector (cscript) / WMI + StdRegProv - no PowerShell required"

Dim sysNotes
sysNotes = ""
If Not adm Then sysNotes = "NOT RUN ELEVATED - collection incomplete"

Dim rowSys : Set rowSys = NewRowDict()
rowSys.Item("RecordType")      = "SystemInfo"
rowSys.Item("ComputerName")    = cn
rowSys.Item("Collected")       = tsDisplay
rowSys.Item("Elevated")        = CStr(adm)
rowSys.Item("Model")           = model
rowSys.Item("Serial")          = serial
rowSys.Item("OSCaption")       = osCaption
rowSys.Item("OSVersion")       = osVersion
rowSys.Item("OSBuild")         = osBuild
rowSys.Item("OSRelease")       = osRelease
rowSys.Item("OSArchitecture")  = osArch
rowSys.Item("Domain")          = domain
rowSys.Item("CollectorMethod") = collectorNote
rowSys.Item("Notes")           = sysNotes
WriteRow csvStream, rowSys

' ---------------- 2. installed programs, every registry scope this host has ----------------
EnumUninstall HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", _
              "Machine", "64-bit", "ALL USERS", csvStream

Dim wowProbe
wowProbe = Null
reg.EnumKey HKLM, "SOFTWARE\WOW6432Node", wowProbe
If Not IsNull(wowProbe) Then
  EnumUninstall HKLM, "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", _
                "Machine", "32-bit", "ALL USERS", csvStream
End If

' Only currently logged-on users' hives are visible under HKEY_USERS. Offline
' profiles are a documented gap on this collector - see the header comment.
Dim userSids, usid, uwowProbe
reg.EnumKey HKEY_USERS, "", userSids
If Not IsNull(userSids) Then
  For Each usid In userSids
    If Left(usid, 9) = "S-1-5-21-" And InStr(usid, "_Classes") = 0 Then
      EnumUninstall HKEY_USERS, usid & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", _
                    "User", "64-bit", usid, csvStream
      uwowProbe = Null
      reg.EnumKey HKEY_USERS, usid & "\SOFTWARE\WOW6432Node", uwowProbe
      If Not IsNull(uwowProbe) Then
        EnumUninstall HKEY_USERS, usid & "\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", _
                      "User", "32-bit", usid, csvStream
      End If
    End If
  Next
End If

' ---------------- 3. hotfixes ----------------
Dim qfe
For Each qfe In wmi.ExecQuery("SELECT HotFixID,Description,InstalledOn,InstalledBy FROM Win32_QuickFixEngineering")
  Dim rowH : Set rowH = NewRowDict()
  rowH.Item("RecordType")   = "Hotfix"
  rowH.Item("ComputerName") = cn
  rowH.Item("Collected")    = tsDisplay
  rowH.Item("Name")         = qfe.HotFixID
  rowH.Item("Notes")        = qfe.Description
  rowH.Item("InstallDate")  = qfe.InstalledOn
  rowH.Item("Owner")        = qfe.InstalledBy
  WriteRow csvStream, rowH
  hfCount = hfCount + 1
Next

' ---------------- 4. loose / portable executables (slow - skip with /noexe) ----------------
If Not skipExe Then
  Dim rootsList(4), rr
  rootsList(0) = shell.ExpandEnvironmentStrings("%SystemDrive%") & "\Users"
  rootsList(1) = shell.ExpandEnvironmentStrings("%SystemDrive%") & "\Documents and Settings"
  rootsList(2) = shell.ExpandEnvironmentStrings("%ProgramData%")
  rootsList(3) = shell.ExpandEnvironmentStrings("%ProgramFiles%")
  rootsList(4) = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
  For Each rr In rootsList
    If Len(rr) > 0 And InStr(rr, "%") = 0 And fso.FolderExists(rr) Then
      WScript.Echo "Scanning " & rr & " for executables..."
      ScanFolderForExe fso.GetFolder(rr), 0, 6, csvStream
    End If
  Next
Else
  WScript.Echo "Skipping executable scan (/noexe)."
End If

' ---------------- 5. Summary.txt ----------------
Set sumStream = fso.CreateTextFile(sumPath, True, False)
sumStream.WriteLine "SOFTWARE INVENTORY EVIDENCE"
sumStream.WriteLine "Host      : " & cn & "   Serial: " & serial
sumStream.WriteLine "Model     : " & model
sumStream.WriteLine "OS        : " & osCaption & "  " & osBuild & "  (" & osArch & ")"
sumStream.WriteLine "Collector : " & shell.ExpandEnvironmentStrings("%USERDOMAIN%") & "\" & shell.ExpandEnvironmentStrings("%USERNAME%")
sumStream.WriteLine "Collected : " & tsDisplay
sumStream.WriteLine "Elevated  : " & CStr(adm)
sumStream.WriteLine "Method    : " & collectorNote & ";"
sumStream.WriteLine "            registry uninstall keys (machine 64/32-bit + logged-on user hives),"
sumStream.WriteLine "            hotfix list" & IIfStr(Not skipExe, ", file system scan.", " (executable scan skipped: /noexe).")
sumStream.WriteLine "            No network calls. Win32_Product NOT queried. Offline (not logged on)"
sumStream.WriteLine "            user profiles are not enumerable without reg.exe - see script header."
sumStream.WriteLine ""
sumStream.WriteLine "Products                 : " & appCount
sumStream.WriteLine "Update/component entries : " & updCount
sumStream.WriteLine "Store (Appx) packages    : 0 (not applicable before Windows 8)"
sumStream.WriteLine "Optional features on     : 0 (not applicable before Windows 8)"
sumStream.WriteLine "Capabilities installed   : 0 (not applicable before Windows 10)"
sumStream.WriteLine "Hotfixes                 : " & hfCount
sumStream.WriteLine "Executables on disk      : " & exCount & IIfStr(skipExe, " (scan skipped: /noexe)", "")
sumStream.WriteLine "Unsigned executables     : not checked (legacy collector - see script header)"
sumStream.WriteLine ""
If Not adm Then
  sumStream.WriteLine "*** COLLECTION INCOMPLETE - NOT RUN ELEVATED ***"
  sumStream.WriteLine ""
End If
sumStream.WriteLine "--- INSTALLED PRODUCTS ---"
Dim pi
For pi = 0 To prodCount - 1
  sumStream.WriteLine prodLines(pi)
Next
sumStream.Close
csvStream.Close

WScript.Echo ""
WScript.Echo "Done. " & appCount & " products. Output: " & outDir

' ================================================================
' functions / subs
' ================================================================

Function FormatTimestamp(d)
  FormatTimestamp = Year(d) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d)) & " " & _
                     Pad2(Hour(d)) & ":" & Pad2(Minute(d)) & ":" & Pad2(Second(d))
End Function

Function FormatStamp(d)
  FormatStamp = Year(d) & Pad2(Month(d)) & Pad2(Day(d)) & "-" & _
                Pad2(Hour(d)) & Pad2(Minute(d)) & Pad2(Second(d))
End Function

Function Pad2(n)
  If n < 10 Then Pad2 = "0" & CStr(n) Else Pad2 = CStr(n)
End Function

Function IsElevated()
  On Error Resume Next
  Err.Clear
  shell.RegWrite "HKLM\SOFTWARE\__SWEvidenceElevTest", "1", "REG_SZ"
  If Err.Number = 0 Then
    IsElevated = True
    Err.Clear
    shell.RegDelete "HKLM\SOFTWARE\__SWEvidenceElevTest"
  Else
    IsElevated = False
  End If
  Err.Clear
  On Error Goto 0
End Function

Function IIfStr(cond, whenTrue, whenFalse)
  If cond Then
    IIfStr = whenTrue
  Else
    IIfStr = whenFalse
  End If
End Function

Function NewRowDict()
  Dim d, i
  Set d = CreateObject("Scripting.Dictionary")
  For i = 0 To UBound(cols)
    d.Add cols(i), ""
  Next
  Set NewRowDict = d
End Function

Sub WriteCsvHeader(stream)
  Dim i, line
  line = ""
  For i = 0 To UBound(cols)
    If i > 0 Then line = line & ","
    line = line & cols(i)
  Next
  stream.WriteLine line
End Sub

' Plain-ANSI CSV output (CreateTextFile's default) rather than Unicode: this is
' the one collector expected to run on Windows 2000, where a non-ANSI text file
' is more likely to trip up whatever the evidence gets opened in downstream.
' A publisher/product name outside the system codepage renders as its closest
' ANSI equivalent or "?" - a limitation worth knowing about, not one silently
' hidden; Get-SoftwareEvidence.ps1 (used everywhere PowerShell is available)
' does not have this limitation.
Sub WriteRow(stream, d)
  Dim i, line, v
  line = ""
  For i = 0 To UBound(cols)
    If i > 0 Then line = line & ","
    v = d.Item(cols(i))
    ' StdRegProv and some WMI providers set an [out] parameter to Null rather than
    ' leaving it at its pre-initialized "" when a value is absent - CStr(Null)
    ' raises "Invalid use of Null", so every field is coalesced here as a backstop
    ' regardless of whether the value came from the registry, WMI, or a literal.
    If IsNull(v) Then v = ""
    line = line & CsvEscape(CStr(v))
  Next
  stream.WriteLine line
End Sub

Function CsvEscape(s)
  If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
    CsvEscape = """" & Replace(s, """", """""") & """"
  Else
    CsvEscape = s
  End If
End Function

Function PadRight(s, n)
  s = CStr(s)
  If Len(s) >= n Then
    PadRight = Left(s, n)
  Else
    PadRight = s & Space(n - Len(s))
  End If
End Function

Sub AddProductLine(nm, line)
  If prodCount > UBound(prodNames) Then
    ReDim Preserve prodNames(UBound(prodNames) + 64)
    ReDim Preserve prodLines(UBound(prodLines) + 64)
  End If
  Dim i
  i = prodCount - 1
  Do While i >= 0
    If StrComp(prodNames(i), nm, vbTextCompare) > 0 Then
      prodNames(i + 1) = prodNames(i)
      prodLines(i + 1) = prodLines(i)
      i = i - 1
    Else
      Exit Do
    End If
  Loop
  prodNames(i + 1) = nm
  prodLines(i + 1) = line
  prodCount = prodCount + 1
End Sub

' One machine-or-user Uninstall key -> Program rows, via StdRegProv rather than
' reg.exe (which Windows 2000 does not ship), de-duplicated by name+version+owner
' the same way Get-SoftwareEvidence.ps1 de-duplicates $sw.
Sub EnumUninstall(hive, rootKey, scope, arch, owner, stream)
  Dim subkeys, sk, keyPath
  Dim dn, dv, pub, instDate, instLoc, uninstStr, parentKey, sysCompDw, isUpd, dedupKey

  reg.EnumKey hive, rootKey, subkeys
  If IsNull(subkeys) Then Exit Sub

  For Each sk In subkeys
    keyPath = rootKey & "\" & sk
    dn = "" : reg.GetStringValue hive, keyPath, "DisplayName", dn
    If IsNull(dn) Then dn = ""
    If Len(dn) > 0 Then
      ' StdRegProv leaves each [out] parameter as Null (not the "" pre-set here)
      ' when the named value does not exist under this key, which is the normal
      ' case for most of these - most Uninstall entries have no ParentKeyName,
      ' for instance. Every one is coalesced back to a safe default immediately.
      dv = ""        : reg.GetStringValue hive, keyPath, "DisplayVersion", dv
      If IsNull(dv) Then dv = ""
      pub = ""       : reg.GetStringValue hive, keyPath, "Publisher", pub
      If IsNull(pub) Then pub = ""
      instDate = ""  : reg.GetStringValue hive, keyPath, "InstallDate", instDate
      If IsNull(instDate) Then instDate = ""
      instLoc = ""   : reg.GetStringValue hive, keyPath, "InstallLocation", instLoc
      If IsNull(instLoc) Then instLoc = ""
      uninstStr = "" : reg.GetStringValue hive, keyPath, "UninstallString", uninstStr
      If IsNull(uninstStr) Then uninstStr = ""
      parentKey = "" : reg.GetStringValue hive, keyPath, "ParentKeyName", parentKey
      If IsNull(parentKey) Then parentKey = ""
      sysCompDw = 0  : reg.GetDWORDValue hive, keyPath, "SystemComponent", sysCompDw
      If IsNull(sysCompDw) Then sysCompDw = 0
      isUpd = (sysCompDw = 1) Or (Len(parentKey) > 0)

      dedupKey = LCase(dn) & "|" & LCase(dv) & "|" & LCase(owner)
      If Not seenProg.Exists(dedupKey) Then
        seenProg.Add dedupKey, True

        Dim row : Set row = NewRowDict()
        row.Item("RecordType")      = "Program"
        row.Item("ComputerName")    = cn
        row.Item("Collected")       = tsDisplay
        row.Item("Name")            = dn
        row.Item("Version")         = dv
        row.Item("Publisher")       = pub
        row.Item("InstallDate")     = instDate
        row.Item("Scope")           = scope
        row.Item("Owner")           = owner
        row.Item("Arch")            = arch
        row.Item("IsUpdate")        = CStr(isUpd)
        row.Item("InstallLocation") = instLoc
        row.Item("UninstallString") = uninstStr
        row.Item("RegistryKey")     = keyPath
        WriteRow stream, row

        If isUpd Then
          updCount = updCount + 1
        Else
          appCount = appCount + 1
          AddProductLine dn, PadRight(dn, 55) & " " & PadRight(dv, 18) & " " & pub
        End If
      End If
    End If
  Next
End Sub

' Depth-bounded recursive .exe scan. On Error Resume Next around the whole walk
' so one access-denied folder (common under other users' profiles when not
' elevated) does not abort the rest of the scan.
Sub ScanFolderForExe(folder, depth, maxDepth, stream)
  If depth > maxDepth Then Exit Sub
  On Error Resume Next
  Dim f, subf
  For Each f In folder.Files
    If LCase(fso.GetExtensionName(f.Name)) = "exe" Then
      Dim rowE : Set rowE = NewRowDict()
      rowE.Item("RecordType")   = "Executable"
      rowE.Item("ComputerName") = cn
      rowE.Item("Collected")    = tsDisplay
      rowE.Item("Name")         = f.Name
      rowE.Item("Path")         = f.Path
      rowE.Item("Modified")     = CStr(f.DateLastModified)
      rowE.Item("Signature")    = "N/A (legacy collector)"
      WriteRow stream, rowE
      exCount = exCount + 1
    End If
  Next
  For Each subf In folder.SubFolders
    ScanFolderForExe subf, depth + 1, maxDepth, stream
  Next
  On Error Goto 0
End Sub
