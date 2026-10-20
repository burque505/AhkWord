# Makes a release download of AhkWord: ..\AhkWord-releases\AhkWord-NN.zip (beside this folder), to attach to a GitHub
# release. It holds the program as last committed (git archive of HEAD: uncommitted changes are left out), without the
# files only a developer needs, and AhkWord.exe: a copy of AutoHotkey v2's AutoHotkey64.exe under AhkWord's name, so
# that AhkWord starts with a double-click on a computer without AutoHotkey (an AutoHotkey started without a script runs
# the .ahk of its own name beside it: AhkWord.ahk). AutoHotkey is under the GNU GPL v2: its license and where its
# source is come with it (licenses\AutoHotkey-GPL-2.0.txt, AhkWord.exe.txt).
#   .\make_release.ps1                         from the installed AutoHotkey (C:\Program Files\AutoHotkey\v2)
#   .\make_release.ps1 -AutoHotkey <path>      from another AutoHotkey64.exe (its license.txt beside it or above it)
param([string]$AutoHotkey = "$env:ProgramFiles\AutoHotkey\v2\AutoHotkey64.exe")
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$releases = Join-Path (Split-Path $root) "AhkWord-releases"

# the iteration (AhkWord.cs's Iteration, as committed)
$cs = git -C $root show HEAD:AhkWord.cs
$iteration = [regex]::Match(($cs -join "`n"), 'public const int Iteration = (\d+);').Groups[1].Value
if (-not $iteration) { throw "AhkWord.cs (as committed) has no Iteration." }
$changes = git -C $root status --porcelain
if ($changes) { "Note: there are uncommitted changes; the release is made from the last commit without them." }

# AutoHotkey: the exe, its version and its license
if (-not (Test-Path $AutoHotkey)) { throw "AutoHotkey64.exe was not found at $AutoHotkey (pass -AutoHotkey <path>)." }
$version = (Get-Item $AutoHotkey).VersionInfo.FileVersion.Trim()
$license = @((Join-Path (Split-Path $AutoHotkey) "license.txt"), (Join-Path (Split-Path (Split-Path $AutoHotkey)) "license.txt")) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $license) { throw "AutoHotkey's license.txt was not found beside $AutoHotkey or in the folder above it." }

New-Item -ItemType Directory -Force $releases | Out-Null
$zip = Join-Path $releases "AhkWord-$iteration.zip"
if (Test-Path $zip) { throw "$zip is there already: remove it to make it again." }
$work = Join-Path ([IO.Path]::GetTempPath()) ("ahkword-release-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
$folder = Join-Path $work "AhkWord"
try {
	# the committed files, without the developer's
	$archive = Join-Path $work "head.zip"
	New-Item -ItemType Directory -Force $work | Out-Null
	git -C $root archive --format=zip -o $archive HEAD
	Expand-Archive $archive $folder
	foreach ($only in "make_version.ps1", "make_release.ps1", ".gitignore", "make_docs.py", "mkdocs.yml") {
		$path = Join-Path $folder $only
		if (Test-Path $path) { [IO.File]::Delete($path) }
	}
	# (the documentation's sources and its workflow: the documentation is on the web, https://burque505.github.io/AhkWord/)
	foreach ($only in "docs", ".github") {
		$path = Join-Path $folder $only
		if (Test-Path $path) { [IO.Directory]::Delete($path, $true) }
	}
	# AhkWord.exe, AutoHotkey's license, and what AhkWord.exe is
	Copy-Item $AutoHotkey (Join-Path $folder "AhkWord.exe")
	Copy-Item $license (Join-Path $folder "licenses\AutoHotkey-GPL-2.0.txt")
	Set-Content (Join-Path $folder "AhkWord.exe.txt") -Encoding UTF8 @(
		"AhkWord.exe is AutoHotkey $version (AutoHotkey64.exe, 64-bit), unchanged, under AhkWord's name.",
		"",
		"AutoHotkey is free software under the GNU General Public License, version 2: see licenses\AutoHotkey-GPL-2.0.txt.",
		"Its source code, for this version: https://github.com/AutoHotkey/AutoHotkey/tree/v$version",
		"AutoHotkey's site: https://www.autohotkey.com/",
		"",
		"Started without a script, AutoHotkey runs the .ahk file of its own name beside it: AhkWord.exe runs AhkWord.ahk.",
		"AhkWord's own files are under the MIT license (LICENSE); AhkWord.exe is only bundled with them.")
	Set-Content (Join-Path $folder "START HERE.txt") -Encoding UTF8 @(
		"AhkWord $iteration",
		"",
		"1. AhkWord needs the .NET 10 Desktop Runtime (once per computer). In a terminal:",
		"      winget install Microsoft.DotNet.DesktopRuntime.10",
		"   or download it from https://dotnet.microsoft.com/download/dotnet/10.0 (.NET Desktop Runtime, x64).",
		"2. Unzip this folder anywhere (not into Program Files), and double-click AhkWord.exe.",
		"   The first start takes a minute or two: it downloads the components AhkWord uses and compiles it. Later starts",
		"   are quick.",
		"3. Pin it to the taskbar if you like (right-click its button > Pin to taskbar).",
		"",
		"If you have AutoHotkey v2 installed, AhkWord.ahk runs AhkWord just the same.",
		"Your settings are kept in %APPDATA%\AhkWord. To keep everything in this folder instead (to carry it on a USB",
		"stick), put an empty file named portable beside AhkWord.exe.",
		"",
		"README.md tells all about it; https://github.com/burque505/AhkWord is its home.")
	Compress-Archive -Path $folder -DestinationPath $zip
} finally {
	if (Test-Path $work) { [IO.Directory]::Delete($work, $true) }
}
$size = "{0:N1} MB" -f ((Get-Item $zip).Length / 1MB)
"AhkWord $iteration with AutoHotkey $version -> $zip ($size)"
