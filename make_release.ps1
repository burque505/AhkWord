# Makes the release downloads of AhkWord, beside this folder in ..\AhkWord-releases, to attach to a GitHub release:
#   AhkWord-NN.zip            everything AhkWord needs but .NET: the program as last committed (git archive of HEAD:
#                             uncommitted changes are left out) without the files only a developer needs; AhkWord.exe;
#                             and (iteration 54) the libraries, the C# compiler's two DLLs and the C# host built
#                             beforehand (their licenses: THIRD-PARTY-NOTICES.md, licenses\), so that nothing is downloaded at the first start and nothing is written
#                             outside AhkWord's folder. Without .NET 10 on the computer, AhkWord offers to download it
#                             for itself (AhkWordSetup.ahk).
#   AhkWord-NN-complete.zip   the same with .NET 10 (its Desktop Runtime) inside: nothing downloaded at all.
# AhkWord.exe is a copy of AutoHotkey v2's AutoHotkey64.exe under AhkWord's name, so that AhkWord starts with a
# double-click on a computer without AutoHotkey (an AutoHotkey started without a script runs the .ahk of its own name
# beside it: AhkWord.ahk). AutoHotkey is under the GNU GPL v2: its license and where its source is come with it.
#   .\make_release.ps1                         from the installed AutoHotkey (C:\Program Files\AutoHotkey\v2)
#   .\make_release.ps1 -AutoHotkey <path>      from another AutoHotkey64.exe (its license.txt beside it or above it)
#   .\make_release.ps1 -Working                from the working tree as it is (for trying a release out; the zips are
#                                              named ...-working)
# The libraries come from Csharp-in-AHK's last NuGet restore for AhkWord's packages (start AhkWord once after a package
# changes), the compiler from the package Csharp-in-AHK downloads (%LOCALAPPDATA%\Csharp-in-AHK\roslyn), .NET from
# Microsoft (fetched once, as AhkWordSetup.ahk does, into ..\AhkWord-releases\dotnet-<version>).
param([string]$AutoHotkey = "$env:ProgramFiles\AutoHotkey\v2\AutoHotkey64.exe", [switch]$Working)
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$releases = Join-Path (Split-Path $root) "AhkWord-releases"

# the iteration (AhkWord.cs's Iteration, as committed, or as it is with -Working)
$cs = if ($Working) { Get-Content (Join-Path $root "AhkWord.cs") } else { git -C $root show HEAD:AhkWord.cs }
$iteration = [regex]::Match(($cs -join "`n"), 'public const int Iteration = (\d+);').Groups[1].Value
if (-not $iteration) { throw "AhkWord.cs has no Iteration." }
$changes = git -C $root status --porcelain
if ($changes -and -not $Working) { "Note: there are uncommitted changes; the release is made from the last commit without them." }
$name = if ($Working) { "AhkWord-$iteration-working" } else { "AhkWord-$iteration" }

# AutoHotkey: the exe, its version and its license
if (-not (Test-Path $AutoHotkey)) { throw "AutoHotkey64.exe was not found at $AutoHotkey (pass -AutoHotkey <path>)." }
$version = (Get-Item $AutoHotkey).VersionInfo.FileVersion.Trim()
$license = @((Join-Path (Split-Path $AutoHotkey) "license.txt"), (Join-Path (Split-Path (Split-Path $AutoHotkey)) "license.txt")) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $license) { throw "AutoHotkey's license.txt was not found beside $AutoHotkey or in the folder above it." }

# the compiler: Microsoft.Net.Compilers.Toolset's tasks\net472 (csc.exe to build the host; its two DLLs go with AhkWord)
$roslyn = Join-Path $env:LOCALAPPDATA "Csharp-in-AHK\roslyn\5.9.0\tasks\net472"
if (-not (Test-Path (Join-Path $roslyn "csc.exe"))) { throw "The C# compiler package isn't in $roslyn (Csharp-in-AHK downloads it on a computer without Visual Studio)." }

# the libraries: the restore whose files include every package AhkWord.cs and its rte_*.cs files name
$wanted = Get-ChildItem $root -Filter "*.cs" | Where-Object { $_.Name -eq "AhkWord.cs" -or $_.Name -like "rte_*.cs" } |
	ForEach-Object { Select-String -Path $_.FullName -Pattern '^#:package\s+([^@\s]+)' } | ForEach-Object { $_.Matches[0].Groups[1].Value.ToLowerInvariant() } | Sort-Object -Unique
$nuget = if ($env:NUGET_PACKAGES) { $env:NUGET_PACKAGES } else { Join-Path $env:USERPROFILE ".nuget\packages" }
$restore = Get-ChildItem (Join-Path $env:LOCALAPPDATA "Csharp-in-AHK\nuget\restore") -Directory | Sort-Object LastWriteTime -Descending |
	Where-Object { $f = Join-Path $_.FullName "assets.txt"; (Test-Path $f) -and -not ($wanted | Where-Object { -not (Select-String -Path $f -SimpleMatch "\$_\" -Quiet) }) } | Select-Object -First 1
if (-not $restore) { throw "No NuGet restore has all of AhkWord's packages ($($wanted -join ', ')): start AhkWord once, then make the release." }
$assets = Get-Content (Join-Path $restore.FullName "assets.txt") | Where-Object { $_ -notmatch '\.pdb$' }

# .NET for the complete download: the latest .NET 10 Desktop Runtime, fetched (and checked) once
$dotnetCache = Get-ChildItem $releases -Directory -Filter "dotnet-10.*" -ErrorAction SilentlyContinue | Sort-Object { [version]($_.Name -replace '^dotnet-', '') } -Descending | Select-Object -First 1
# (iteration 59) fetched again when Microsoft has released a newer one (a security fix, most months)
try {
	$latest = (Invoke-RestMethod "https://builds.dotnet.microsoft.com/dotnet/release-metadata/10.0/releases.json" -TimeoutSec 30)."latest-runtime"
	if ($dotnetCache -and $latest -and [version]$latest -gt [version]($dotnetCache.Name -replace '^dotnet-', '')) {
		"Microsoft has released .NET $latest (the cache has $($dotnetCache.Name -replace '^dotnet-', '')): fetching it."
		$dotnetCache = $null
	}
} catch { "Note: Microsoft's list of .NET releases couldn't be read; the cached .NET is used." }
New-Item -ItemType Directory -Force $releases | Out-Null
$work = Join-Path ([IO.Path]::GetTempPath()) ("ahkword-release-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Force $work | Out-Null
try {
	if (-not $dotnetCache) {
		"Fetching .NET 10 from Microsoft (once) ..."
		$fetch = Join-Path $work "fetch.ahk"
		$into = Join-Path $work "dotnet"
		Set-Content $fetch -Encoding UTF8 @("#Include `"$root\Csharp-in-AHK.ahk`"", "#Include `"$root\AhkWordSetup.ahk`"", "_AhkWordFetchDotnet(A_Args[1])")
		& $AutoHotkey /ErrorStdOut $fetch $into | Out-Host
		$desktop = Get-ChildItem (Join-Path $into "shared\Microsoft.WindowsDesktop.App") -Directory | Select-Object -First 1
		if (-not $desktop) { throw ".NET could not be fetched." }
		$dotnetCache = Join-Path $releases ("dotnet-" + $desktop.Name)
		Move-Item $into $dotnetCache
		$dotnetCache = Get-Item $dotnetCache
	}
	$folder = Join-Path $work "AhkWord"
	# the program's files, without the developer's
	if ($Working) {
		New-Item -ItemType Directory -Force $folder | Out-Null
		foreach ($file in git -C $root ls-files -co --exclude-standard) {
			$from = Join-Path $root $file
			if (-not (Test-Path $from -PathType Leaf)) { continue }
			$to = Join-Path $folder $file
			New-Item -ItemType Directory -Force (Split-Path $to) | Out-Null
			Copy-Item $from $to
		}
	} else {
		$archive = Join-Path $work "head.zip"
		git -C $root archive --format=zip -o $archive HEAD
		Expand-Archive $archive $folder
	}
	foreach ($only in "make_version.ps1", "make_release.ps1", ".gitignore", "make_docs.py", "mkdocs.yml", "CLAUDE.local.md", "Csharp-in-AHK.ini") {
		$path = Join-Path $folder $only
		if (Test-Path $path) { [IO.File]::Delete($path) }
	}
	# (the documentation's sources and its workflow: the documentation is on the web, https://burque505.github.io/AhkWord/)
	foreach ($only in "docs", ".github") {
		$path = Join-Path $folder $only
		if (Test-Path $path) { [IO.Directory]::Delete($path, $true) }
	}

	# (iteration 54) the libraries in lib\packages, as NuGet lays them out, and their list for Csharp-in-AHK
	$lib = Join-Path $folder "lib"
	$manifest = New-Object System.Collections.Generic.List[string]
	$packages = @{}
	foreach ($line in $assets) {
		$kind, $path = $line -split '\|', 2
		if (-not $path.StartsWith($nuget, [StringComparison]::OrdinalIgnoreCase)) { throw "A library outside the NuGet folder: $path" }
		$relative = $path.Substring($nuget.Length).TrimStart('\')
		$to = Join-Path $lib "packages\$relative"
		if (-not (Test-Path $to)) { New-Item -ItemType Directory -Force (Split-Path $to) | Out-Null; Copy-Item $path $to }
		$manifest.Add("$kind|packages\$relative")
		$id, $ver = ($relative -split '\\')[0, 1]
		$packages["$id|$ver"] = $true
	}
	Set-Content (Join-Path $lib "packages.txt") -Encoding UTF8 $manifest
	# the compiler's DLLs (the scripts are compiled at each start), and the host, built here with that compiler
	New-Item -ItemType Directory -Force (Join-Path $lib "roslyn"), (Join-Path $lib "host") | Out-Null
	foreach ($dll in "Microsoft.CodeAnalysis.dll", "Microsoft.CodeAnalysis.CSharp.dll") { Copy-Item (Join-Path $roslyn $dll) (Join-Path $lib "roslyn\$dll") }
	$build = Join-Path $work "build-host.ahk"
	Set-Content $build -Encoding UTF8 @("#Include `"$root\Csharp-in-AHK.ahk`"", "global CSRoslynDir := `"$roslyn`"", "CSharpBuildHost(A_Args[1], `"CSHost-AhkWord`")")
	& $AutoHotkey /ErrorStdOut $build (Join-Path $lib "host") | Out-Host
	if (-not (Test-Path (Join-Path $lib "host\CSHost-AhkWord.dll"))) { throw "The C# host could not be built." }
	Set-Content (Join-Path $folder "Csharp-in-AHK.ini") -Encoding UTF8 @(
		"; AhkWord's own: its libraries, the compiler's DLLs and the C# host are in lib, so that nothing is downloaded and",
		"; nothing written outside this folder (see Csharp-in-AHK.ahk). The .NET it runs on: AhkWordSetup.ahk.",
		"[Paths]", "RoslynDir=lib\roslyn", "HostDll=lib\host\CSHost-AhkWord.dll", "PackagesManifest=lib\packages.txt", "CacheDir=lib\cache")

	# AhkWord.exe, the licenses, what AhkWord.exe is, and the other parts' notices
	Copy-Item $AutoHotkey (Join-Path $folder "AhkWord.exe")
	New-Item -ItemType Directory -Force (Join-Path $folder "licenses") | Out-Null
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
	# (the other parts' licenses: THIRD-PARTY-NOTICES.md and licenses\, which must name every package carried)
	$named = Get-Content (Join-Path $folder "THIRD-PARTY-NOTICES.md") -Raw
	foreach ($key in $packages.Keys | Sort-Object) {
		$id = ($key -split '\|')[0]
		$spec = Get-ChildItem (Join-Path $nuget "$id\$(($key -split '\|')[1])") -Filter "*.nuspec" | Select-Object -First 1
		$real = ([xml](Get-Content $spec.FullName -Raw)).package.metadata.id
		if ($named -notmatch [regex]::Escape($real)) { throw "THIRD-PARTY-NOTICES.md doesn't name $real, which the release carries." }
	}
	Set-Content (Join-Path $folder "START HERE.txt") -Encoding UTF8 @(
		"AhkWord $iteration",
		"",
		"1. Unzip this folder anywhere you can write to (your Documents, a USB stick; not Program Files).",
		"2. Double-click AhkWord.exe.",
		"   - Windows may say ""Windows protected your PC"" the first time: AhkWord isn't signed (signing costs money).",
		"     Click ""More info"", then ""Run anyway"".",
		"   - AhkWord runs on Microsoft's .NET 10 Desktop Runtime. If the computer hasn't got it, AhkWord offers to",
		"     download it from Microsoft for itself only (about 70 MB, into this folder: nothing is installed in Windows,",
		"     and deleting the folder removes it), or to open Microsoft's page to install it for all of Windows.",
		"     The ""complete"" download (AhkWord-$iteration-complete.zip) has it inside already.",
		"   - The first start takes a little longer than later ones.",
		"3. Pin it to the taskbar if you like (right-click its button > Pin to taskbar).",
		"",
		"Just trying AhkWord out? Make it portable first: put an empty file named portable (or portable.txt) beside",
		"AhkWord.exe. Then everything it keeps (settings, recent files, templates, recovery copies) stays in a data",
		"folder here, and deleting this folder leaves no trace. Without it, settings are kept for your Windows user in",
		"%APPDATA%\AhkWord, shared by every copy of AhkWord, so that a newer release you unzip elsewhere keeps them.",
		"",
		"Everything AhkWord needs is in this folder; nothing is downloaded but .NET, and only if you say so.",
		"To remove AhkWord, delete this folder (and %APPDATA%\AhkWord and %LOCALAPPDATA%\AhkWord, if not portable).",
		"",
		"If you have AutoHotkey v2 installed, AhkWord.ahk runs AhkWord just the same.",
		"README.md tells all about it; https://github.com/burque505/AhkWord is its home.")

	# the two downloads
	foreach ($zipName in "$name.zip", "$name-complete.zip") {
		$zip = Join-Path $releases $zipName
		if (Test-Path $zip) { throw "$zip is there already: remove it to make it again." }
	}
	Compress-Archive -Path $folder -DestinationPath (Join-Path $releases "$name.zip")
	Copy-Item $dotnetCache.FullName (Join-Path $folder "dotnet") -Recurse
	Compress-Archive -Path $folder -DestinationPath (Join-Path $releases "$name-complete.zip")
} finally {
	if (Test-Path $work) { [IO.Directory]::Delete($work, $true) }
}
foreach ($zipName in "$name.zip", "$name-complete.zip") {
	"{0} ({1:N1} MB)" -f (Join-Path $releases $zipName), ((Get-Item (Join-Path $releases $zipName)).Length / 1MB)
}
"AhkWord $iteration with AutoHotkey $version and .NET $($dotnetCache.Name -replace '^dotnet-', '')"
