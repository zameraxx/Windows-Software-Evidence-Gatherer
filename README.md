# Windows Software Evidence Gatherer

Offline software-inventory evidence collector for a single Windows host. Run it
elevated on a machine and it produces exactly **two files** — one CSV with every
record, one plain-text summary — ready to hand to an assessor or drop straight
into [Palisade](#palisade-ingestion). No network calls, ever.

Covers **Windows 2000 through the current release** by picking one of two
collectors depending on what the machine actually has available:

| Machine has...                                   | Collector used                        |
|----------------------------------------------------|----------------------------------------|
| PowerShell (Vista/Server 2008+, or XP/2003 with PowerShell installed) | `Get-SoftwareEvidence.ps1` |
| No PowerShell (Windows 2000, or XP/2003 that never had it installed)  | `Get-SoftwareEvidence-Legacy.vbs` |

You don't need to know which one applies — **run `Get-SoftwareEvidence.bat`**
and it detects this automatically and calls the right one.

## Quick start

1. Copy all three files (`Get-SoftwareEvidence.bat`, `Get-SoftwareEvidence.ps1`,
   `Get-SoftwareEvidence-Legacy.vbs`) to the target machine, in the same folder.
2. Right-click `Get-SoftwareEvidence.bat` → **Run as administrator**.
   (Windows 2000 has no "Run as administrator" — just be logged on as a user in
   the Administrators group.) It still runs unelevated, but the collection will
   be partial and both the console and `Summary.txt` say so plainly.
3. Find the output on the Desktop, in a folder named
   `SWEvidence_<hostname>_<timestamp>`.

To skip the (slow) loose-executable file system scan — worth doing on an old
machine with a large or slow disk — run `Get-SoftwareEvidence.bat /noexe`.

You can also run either collector directly if you already know which one
applies: `powershell -ExecutionPolicy Bypass -File Get-SoftwareEvidence.ps1`
(add `-NoExe` to skip the file scan) or
`cscript //nologo Get-SoftwareEvidence-Legacy.vbs` (add `/noexe`).

## Output

Exactly two files per run, both inside the timestamped folder:

- **`SoftwareEvidence_<host>_<timestamp>.csv`** — every record the collector
  found, one row each, tagged by a `RecordType` column so a single wide CSV can
  hold several shapes of data without seven separate files. Both collectors
  emit the identical 30-column schema, so a fleet mixing old and new Windows
  produces one consistent format:

  | RecordType        | What it is                                              | Windows 2000/XP | Windows 8+/10+ |
  |--------------------|----------------------------------------------------------|:---:|:---:|
  | `SystemInfo`       | one row: host name, model, serial, OS, elevation, collector method | yes | yes |
  | `Program`          | one row per installed product (registry Uninstall keys, all scopes) | yes | yes |
  | `AppxPackage`      | Windows Store apps                                        | — | yes (8+) |
  | `OptionalFeature`  | enabled Windows optional features                         | — | yes (8+) |
  | `Capability`       | installed Windows capabilities                             | — | yes (10+) |
  | `Hotfix`           | installed hotfixes/KBs                                     | yes | yes |
  | `Executable`       | loose/portable `.exe` files found under user profiles and Program Files (depth-limited to 6) | yes | yes |

  RecordTypes that don't apply to a given OS simply produce zero rows — the
  column layout never changes, only which rows show up.

- **`Summary_<host>_<timestamp>.txt`** — plain-text summary: host/OS identity,
  collection method, counts per category, and the full installed-product list.

### Column reference (CSV)

`RecordType, ComputerName, Collected, Elevated, Name, Version, Publisher,
InstallDate, Scope, Owner, Arch, IsUpdate, InstallLocation, UninstallString,
RegistryKey, Path, Modified, Signature, Signer, State, Model, Serial,
OSCaption, OSVersion, OSBuild, OSRelease, OSArchitecture, Domain,
CollectorMethod, Notes`

Not every column applies to every `RecordType` — a `Program` row leaves the
`OSCaption`/`Model`/`Serial` columns blank (those live on the one `SystemInfo`
row), an `Executable` row leaves `Scope`/`Owner`/`InstallDate` blank, and so on.

## What each collector does

### `Get-SoftwareEvidence.ps1`

Written against PowerShell **2.0** syntax on purpose (`Get-WmiObject` instead
of `Get-CimInstance`, no `[pscustomobject]`, no `Get-ChildItem -Depth`) so the
one script runs unmodified from Windows Vista/Server 2008 through whatever
ships after this was written, and on Windows XP SP2+/Server 2003 SP1+ if
PowerShell was installed on them. Feature-gates itself by checking whether a
cmdlet exists (`Get-AppxPackage`, `Get-WindowsOptionalFeature`,
`Get-WindowsCapability`) rather than hardcoding version numbers, so it also
does the right thing on Server SKUs and future Windows releases without
editing. Offline (not logged on) user profiles are read by mounting their
`NTUSER.DAT` via `reg load`/`reg unload` against every SID under
`HKLM\...\ProfileList` — a method that works on every version back to Windows
2000, unlike the `Win32_UserProfile` WMI class it replaces (Vista+ only).

### `Get-SoftwareEvidence-Legacy.vbs`

For Windows 2000, which cannot run **any** version of PowerShell, and for
XP/2003 machines that never had PowerShell installed. Runs under `cscript.exe`
(Windows Script Host — present by default since Windows 98), using WMI and the
`StdRegProv` registry provider instead of `reg.exe` (which Windows 2000 does
not ship by default). Documented, known gaps versus the PowerShell collector:

- **Offline user profiles are not enumerated.** Doing so needs `reg.exe` to
  load/unload `NTUSER.DAT`, which Windows 2000 doesn't ship. Only currently
  logged-on users' hives (visible under `HKEY_USERS`) are read.
- **No Authenticode signature check** on executables found on disk — plain
  VBScript has no path to `WinVerifyTrust` without an external helper. The
  `Signature` column reads `N/A (legacy collector)` instead.
- **CSV is written in the system ANSI codepage, not UTF-8** — the safer choice
  for whatever Windows 2000 ends up viewing it in. A product or publisher name
  with characters outside that codepage renders as its closest equivalent or
  `?`. `Get-SoftwareEvidence.ps1` (used everywhere PowerShell is available)
  does not have this limitation — its CSV is UTF-8.
- Store apps / optional features / Windows capabilities don't exist as
  concepts before Windows 8/8/10 respectively, so those `RecordType`s are
  simply empty on older OS — not a bug, just nothing to report.

### `Get-SoftwareEvidence.bat`

Thin dispatcher: finds `powershell.exe` without relying on `PATH` (checks
`%SystemRoot%\System32\WindowsPowerShell\v1.0\` and the SysWOW64 equivalent
directly), runs the PowerShell collector if found, otherwise falls back to the
VBScript collector. Also checks elevation (`net session`) up front and warns
if it's missing, before either collector gets a chance to. Forwards `/noexe`
to whichever collector it ends up calling, translating it to that collector's
own flag syntax.

## Design notes

- **No network calls, ever.** Every collector only reads local registry, WMI,
  and file system state.
- **`Win32_Product` is never queried** — enumerating it silently triggers an
  MSI consistency/repair pass on every installed package, which is not
  something an evidence collector should be doing as a side effect.
- **Partial evidence beats none.** `Get-SoftwareEvidence.ps1` wraps its
  collection in `try/finally` so a mid-run error still writes whatever it
  collected up to that point rather than losing the run.
- **De-duplication**: the same product can appear more than once across
  registry scopes (machine 64-bit, machine 32-bit under WOW6432Node, per-user);
  both collectors de-duplicate `Program` rows by name + version + owner.

## Palisade ingestion

[Palisade](https://github.com/) recognizes `SoftwareEvidence_<host>_<timestamp>.csv`
automatically — drop it on the Overview tab (or drag it in) alongside your
Tenable/ACAS exports. It's detected by its `RecordType`/`ComputerName`/
`Collected`/`OSArchitecture` header signature, and is treated as **credentialed
evidence** exactly the way a credentialed Nessus scan is: the `SystemInfo` row
sets the host's OS, last-seen date, and scan-health fields; non-update
`Program`/`AppxPackage` rows feed the Software Listing and drive STIG
applicability detection. `Hotfix`/`OptionalFeature`/`Capability`/`Executable`
rows still register the host (so it shows up even if nothing else did) but are
deliberately left out of the Software Listing itself — they're a different
kind of evidence (patches, loose binaries) than "installed product," the same
reason Palisade already filters KB/patch lines out of its `.nessus` ingestion.

Palisade still loads the older two-file `Get-SoftwareEvidence-Short` format
(`SystemInfo.csv` + `InstalledPrograms.csv`) if you have evidence collected
before this consolidation — nothing has to be re-collected.
