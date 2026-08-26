@echo off
setlocal

rem ================================================================
rem Get-SoftwareEvidence.bat - picks the right collector for this
rem machine and runs it. This is the file to double-click / hand to
rem someone; they should not need to know which script underneath
rem applies to their OS.
rem
rem   PowerShell present -> Get-SoftwareEvidence.ps1
rem                         (Windows Vista/Server 2008 and later; also
rem                          Windows XP SP2+/Server 2003 SP1+ if PowerShell
rem                          was installed on them)
rem   PowerShell absent  -> Get-SoftwareEvidence-Legacy.vbs
rem                         (Windows 2000, which cannot run any version of
rem                          PowerShell; or XP/2003 that never had it installed)
rem
rem For a complete collection, right-click this file and choose "Run as
rem administrator" (Windows XP and later). Windows 2000 has no such option -
rem just be logged on as a user in the Administrators group. It still runs
rem without that, but the collection will be partial and both the console
rem and Summary.txt will say so.
rem
rem Usage:
rem   Get-SoftwareEvidence.bat            normal run
rem   Get-SoftwareEvidence.bat /noexe     skip the (slow) loose-executable
rem                                       file system scan
rem ================================================================

set "HERE=%~dp0"
set "PS1=%HERE%Get-SoftwareEvidence.ps1"
set "VBS=%HERE%Get-SoftwareEvidence-Legacy.vbs"

rem translate the one shared flag into each collector's own syntax
set "NOEXE_PS="
set "NOEXE_VBS="
if /I "%~1"=="/noexe" (
  set "NOEXE_PS=-NoExe"
  set "NOEXE_VBS=/noexe"
)

rem locate powershell.exe without relying on PATH (some old/locked-down
rem systems do not have it on PATH even when it is installed)
set "PSEXE="
if exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" (
  set "PSEXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
) else if exist "%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" (
  set "PSEXE=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
)

net session >nul 2>&1
if not %errorlevel%==0 (
  echo ------------------------------------------------------------------
  echo NOT running elevated. Right-click this file and choose "Run as
  echo administrator" for a complete collection ^(Windows 2000: just be
  echo logged on as an Administrator - there is no "Run as" for that OS^).
  echo Continuing anyway; the collector will flag the gap in its output.
  echo ------------------------------------------------------------------
  echo.
)

if defined PSEXE (
  if not exist "%PS1%" (
    echo ERROR: PowerShell was found, but Get-SoftwareEvidence.ps1 is not next
    echo to this .bat file. Put both files in the same folder and try again.
    goto :end
  )
  echo PowerShell detected - running Get-SoftwareEvidence.ps1
  echo.
  "%PSEXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %NOEXE_PS%
) else (
  if not exist "%VBS%" (
    echo ERROR: PowerShell was not found on this machine, and
    echo Get-SoftwareEvidence-Legacy.vbs is not next to this .bat file. Put
    echo both files in the same folder and try again.
    goto :end
  )
  echo No PowerShell on this machine ^(expected on Windows 2000 / older XP^) -
  echo running the legacy VBScript collector instead.
  echo.
  cscript.exe //nologo "%VBS%" %NOEXE_VBS%
)

:end
echo.
pause
