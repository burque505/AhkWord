# Iteration 88: makes AhkWord's installer, ..\AhkWord-releases\AhkWord-NN-setup.exe (Inno Setup 6, installer\AhkWord.iss),
# and with -Publish puts it on GitHub as the pre-release "AhkWord NN (installer preview)" (tag installer-NN).
# After iteration 90: the same files also as a portable zip, ..\AhkWord-releases\AhkWord-NN-portable.zip (a "portable" file
# beside AhkWord.exe: settings and the rest in its data folder, nothing in the registry), published beside the installer.
# What it installs: AhkWord.exe built by vs\AhkWord.csproj from the same sources AhkWord.ahk runs, published
# self-contained (.NET 10 inside: nothing to download); AhkWord's own files beside it (as in AhkWord-NN.zip, without the
# developer's); AutoHotkey64.exe beside it (AhkWordCompanion.ahk: hotkeys, snippets' hotstrings, macros).
#   .\make_installer.ps1                       from the last commit (uncommitted changes are left out)
#   .\make_installer.ps1 -Working              from the working tree as it is (named ...-working; for trying it out)
#   .\make_installer.ps1 -Publish              then create or update the GitHub pre-release (needs a clean, pushed main)
#   -AutoHotkey <path>                         another AutoHotkey64.exe (its license.txt beside it or above it)
# Neither AhkWord.exe nor the installer is signed (yet): the release notes say what Windows will show.
param([string]$AutoHotkey = "$env:ProgramFiles\AutoHotkey\v2\AutoHotkey64.exe", [switch]$Working, [switch]$Publish)
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$releases = Join-Path (Split-Path $root) "AhkWord-releases"
if ($Working -and $Publish) { throw "-Publish makes the installer from the last commit: leave out -Working." }

# the iteration (AhkWord.cs's Iteration, as committed, or as it is with -Working)
$cs = if ($Working) { Get-Content (Join-Path $root "AhkWord.cs") } else { git -C $root show HEAD:AhkWord.cs }
$iteration = [regex]::Match(($cs -join "`n"), 'public const int Iteration = (\d+);').Groups[1].Value
if (-not $iteration) { throw "AhkWord.cs has no Iteration." }
$changes = git -C $root status --porcelain
if ($Publish) {
	if ($changes) { throw "-Publish: commit first (there are uncommitted changes)." }
	git -C $root fetch -q origin main
	if ((git -C $root rev-parse HEAD) -ne (git -C $root rev-parse origin/main)) { throw "-Publish: push main first (HEAD isn't origin/main)." }
} elseif ($changes -and -not $Working) { "Note: there are uncommitted changes; the installer is made from the last commit without them." }
$name = if ($Working) { "AhkWord-$iteration-working-setup" } else { "AhkWord-$iteration-setup" }
$portableName = if ($Working) { "AhkWord-$iteration-working-portable" } else { "AhkWord-$iteration-portable" }

# the tools: Inno Setup's compiler, AutoHotkey and its license
$iscc = @("$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe", "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "$env:ProgramFiles\Inno Setup 6\ISCC.exe") |
	Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw "Inno Setup 6 isn't installed (winget install JRSoftware.InnoSetup)." }
if (-not (Test-Path $AutoHotkey)) { throw "AutoHotkey64.exe was not found at $AutoHotkey (pass -AutoHotkey <path>)." }
$ahkVersion = (Get-Item $AutoHotkey).VersionInfo.FileVersion.Trim()
$license = @((Join-Path (Split-Path $AutoHotkey) "license.txt"), (Join-Path (Split-Path (Split-Path $AutoHotkey)) "license.txt")) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $license) { throw "AutoHotkey's license.txt was not found beside $AutoHotkey or in the folder above it." }

New-Item -ItemType Directory -Force $releases | Out-Null
$work = Join-Path ([IO.Path]::GetTempPath()) ("ahkword-installer-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Force $work | Out-Null
try {
	# the sources: the last commit (git archive), or the working tree's files
	$src = Join-Path $work "src"
	if ($Working) {
		foreach ($file in git -C $root ls-files -co --exclude-standard) {
			$from = Join-Path $root $file
			if (-not (Test-Path $from -PathType Leaf)) { continue }
			$to = Join-Path $src $file
			New-Item -ItemType Directory -Force (Split-Path $to) | Out-Null
			Copy-Item $from $to
		}
	} else {
		$archive = Join-Path $work "head.zip"
		git -C $root archive --format=zip -o $archive HEAD
		Expand-Archive $archive $src
	}

	# AhkWord.exe, self-contained (no .pdb: nothing of this computer's paths goes with it)
	"Building AhkWord.exe (dotnet publish, self-contained) ..."
	# the .NET inside: the latest .NET 10 (a security fix most months), as make_release.ps1's complete zip, pinned by
	# RuntimeFrameworkVersion (else the SDK's own, often older; both runtime packs follow it, fetched by the restore)
	try { $latest = (Invoke-RestMethod "https://builds.dotnet.microsoft.com/dotnet/release-metadata/10.0/releases.json" -TimeoutSec 30)."latest-runtime" }
	catch { $latest = $null; "Note: Microsoft's list of .NET releases couldn't be read; the SDK's .NET is used." }
	$pin = if ($latest) { "-p:RuntimeFrameworkVersion=$latest" } else { "-p:AhkWordNoPin=1" }
	$pubDir = Join-Path $work "publish"
	dotnet publish (Join-Path $src "vs\AhkWord.csproj") -c Release -r win-x64 --self-contained true -o $pubDir $pin `
		-p:DebugType=none -p:DebugSymbols=false -p:Deterministic=true -p:ContinuousIntegrationBuild=true -nologo -v:q | Out-Host
	if ($LASTEXITCODE -ne 0 -or -not (Test-Path (Join-Path $pubDir "AhkWord.exe"))) { throw "dotnet publish failed." }
	$included = (Get-Content (Join-Path $pubDir "AhkWord.runtimeconfig.json") -Raw | ConvertFrom-Json).runtimeOptions.includedFrameworks
	$dotnetVersion = ($included | Where-Object name -eq "Microsoft.NETCore.App").version
	$desktopVersion = ($included | Where-Object name -eq "Microsoft.WindowsDesktop.App").version
	if ($latest -and ([version]$dotnetVersion -ne [version]$latest -or [version]$desktopVersion -ne [version]$latest)) {
		throw ".NET $dotnetVersion / Windows Desktop $desktopVersion came out, not $latest."
	}
	"AhkWord.exe with .NET $dotnetVersion (Windows Desktop $desktopVersion)"

	# the program's files, without the developer's, and AhkWord.exe's beside them
	$folder = Join-Path $work "AhkWord"
	Copy-Item $src $folder -Recurse
	foreach ($only in "make_version.ps1", "make_release.ps1", "make_installer.ps1", ".gitignore", "make_docs.py", "mkdocs.yml", "CLAUDE.local.md", "Csharp-in-AHK.ini") {
		$path = Join-Path $folder $only
		if (Test-Path $path) { [IO.File]::Delete($path) }
	}
	foreach ($only in "docs", ".github", "vs", "installer") {
		$path = Join-Path $folder $only
		if (Test-Path $path) { [IO.Directory]::Delete($path, $true) }
	}
	foreach ($item in Get-ChildItem $pubDir -Recurse -File) {
		$relative = $item.FullName.Substring($pubDir.Length).TrimStart('\')
		$to = Join-Path $folder $relative
		if (Test-Path $to) { throw "AhkWord.exe's file $relative has the name of one of AhkWord's own files." }
		New-Item -ItemType Directory -Force (Split-Path $to) | Out-Null
		Copy-Item $item.FullName $to
	}

	# AutoHotkey beside AhkWord.exe, its license and what it is there for
	Copy-Item $AutoHotkey (Join-Path $folder "AutoHotkey64.exe")
	New-Item -ItemType Directory -Force (Join-Path $folder "licenses") | Out-Null
	Copy-Item $license (Join-Path $folder "licenses\AutoHotkey-GPL-2.0.txt")
	# .NET's license and notices, from the runtime packs the publish used
	$nuget = if ($env:NUGET_PACKAGES) { $env:NUGET_PACKAGES } else { Join-Path $env:USERPROFILE ".nuget\packages" }
	$dotnetLicenses = Join-Path $folder "licenses\dotnet-runtime"
	New-Item -ItemType Directory -Force $dotnetLicenses | Out-Null
	$core = Join-Path $nuget "microsoft.netcore.app.runtime.win-x64\$dotnetVersion"
	$desk = Join-Path $nuget "microsoft.windowsdesktop.app.runtime.win-x64\$desktopVersion"
	Copy-Item (Join-Path $core "LICENSE.TXT") (Join-Path $dotnetLicenses "LICENSE.txt")
	Copy-Item (Join-Path $core "THIRD-PARTY-NOTICES.TXT") (Join-Path $dotnetLicenses "ThirdPartyNotices.txt")
	Copy-Item (Join-Path $desk "LICENSE") (Join-Path $dotnetLicenses "WindowsDesktop-LICENSE.txt")
	Set-Content (Join-Path $folder "AutoHotkey64.exe.txt") -Encoding UTF8 @(
		"AutoHotkey64.exe is AutoHotkey $ahkVersion (64-bit), unchanged. AhkWord.exe starts it with AhkWordCompanion.ahk",
		"for AhkWord's hotkeys, the snippets' hotstrings and the macros.",
		"",
		"AutoHotkey is free software under the GNU General Public License, version 2: see licenses\AutoHotkey-GPL-2.0.txt.",
		"Its source code, for this version: https://github.com/AutoHotkey/AutoHotkey/tree/v$ahkVersion",
		"AutoHotkey's site: https://www.autohotkey.com/",
		"",
		"AhkWord's own files are under the MIT license (LICENSE); AutoHotkey64.exe is only bundled with them.")

	# every package AhkWord.exe carries must be named in THIRD-PARTY-NOTICES.md
	$named = Get-Content (Join-Path $folder "THIRD-PARTY-NOTICES.md") -Raw
	$unnamed = @()
	$assetsJson = Get-Content (Join-Path $src "vs\obj\project.assets.json") -Raw | ConvertFrom-Json
	foreach ($lib in $assetsJson.libraries.PSObject.Properties) {
		if ($lib.Value.type -ne "package") { continue }
		$id = ($lib.Name -split '/')[0]
		if ($id -like "Microsoft.NETCore.App.*" -or $id -like "Microsoft.WindowsDesktop.App.*" -or $id -like "Microsoft.AspNetCore.App.*") { continue }
		if ($id -like "*.NativeAssets.*" -and $id -notlike "*.NativeAssets.Win32") { continue }    # (other systems' native libraries: not carried on win-x64)
		if ($named -notmatch [regex]::Escape($id)) { $unnamed += @("$($lib.Name) ($(@($lib.Value.files | Where-Object { $_ -like 'lib/*.dll' }).Count) lib dlls)") }
	}
	if ($unnamed) { throw "THIRD-PARTY-NOTICES.md doesn't name these packages, which AhkWord.exe carries:`n$($unnamed -join "`n")" }

	# the privacy check over what is shipped: the text files, and the built AhkWord.dll/.exe for this computer's paths
	# (the pattern is the owner's, kept outside the repository: ..\AhkWord-tests\privacy_pattern.txt)
	$patternFile = Join-Path (Split-Path $root) "AhkWord-tests\privacy_pattern.txt"
	if (-not (Test-Path $patternFile)) { throw "The privacy check's pattern isn't at $patternFile." }
	$pattern = (Get-Content $patternFile -Raw).Trim()
	$userPart = [regex]::Escape($env:USERNAME)
	$texts = Get-ChildItem $folder -Recurse -File | Where-Object { $_.Extension -in ".ahk", ".cs", ".txt", ".md", ".json", ".xml", ".html", ".css", ".ini", ".rtf" }
	$hits = $texts | Select-String -Pattern $pattern -List
	foreach ($built in "AhkWord.dll", "AhkWord.exe") {
		$bytes = [IO.File]::ReadAllBytes((Join-Path $folder $built))
		foreach ($enc in [Text.Encoding]::Latin1, [Text.Encoding]::Unicode) {
			if ($enc.GetString($bytes) -match ('\\Users\\' + $userPart)) { $hits += "$built has this computer's user folder in it" }
		}
	}
	if ($hits) { $hits | Out-Host; throw "The privacy check found something in what the installer would carry." }

	# the installer
	"Making the installer (Inno Setup) ..."
	$iss = Join-Path $src "installer\AhkWord.iss"
	& $iscc /Q "/DIteration=$iteration" "/DSourceDir=$folder" "/DOutputDir=$releases" "/DRepo=$src" "/F$name" $iss | Out-Host
	if ($LASTEXITCODE -ne 0) { throw "Inno Setup failed." }

	# the portable zip: the same files in a folder AhkWord-NN, with the "portable" file (AppFolders.Portable)
	"Making the portable zip ..."
	New-Item -ItemType File -Force (Join-Path $folder "portable") | Out-Null
	Set-Content (Join-Path $folder "README-portable.txt") -Encoding UTF8 @(
		"AhkWord $iteration, portable",
		"",
		"1. Unzip this folder anywhere you can write to (your Documents, a USB stick; not Program Files).",
		"   Tip: before unzipping, right-click the zip > Properties > Unblock, so Windows asks only once.",
		"2. Double-click AhkWord.exe. Nothing needs installing: .NET 10 and AutoHotkey are inside.",
		"   - Windows may say ""Windows protected your PC"": AhkWord isn't signed yet. Click ""More info"", then ""Run anyway"".",
		"   - Smart App Control, if it is on, blocks unsigned programs and offers no way past it.",
		"",
		"The empty file named portable beside AhkWord.exe keeps everything AhkWord remembers (settings, recent files,",
		"templates, snippets, recovery copies) in the data folder here. Nothing goes into the registry or %APPDATA%, and",
		"deleting this folder removes AhkWord completely. The data folder appears the first time AhkWord saves something",
		"(when you close it, or change a setting), not at its first start. Delete the portable file to share the settings in",
		"%APPDATA%\AhkWord with the installed AhkWord instead.",
		"",
		"A newer release: unzip it beside this one and move this folder's data folder into it.",
		"",
		"AutoHotkey64.exe here is AhkWord's own private copy (see AutoHotkey64.exe.txt); an AutoHotkey you have installed",
		"is not touched. README.md tells all about AhkWord; https://github.com/burque505/AhkWord is its home.")
	$portableFolder = Join-Path $work "AhkWord-$iteration"
	Rename-Item $folder $portableFolder
	$portableZip = Join-Path $releases "$portableName.zip"
	if (Test-Path $portableZip) { Remove-Item $portableZip }
	Add-Type -AssemblyName System.IO.Compression.FileSystem
	[IO.Compression.ZipFile]::CreateFromDirectory($portableFolder, $portableZip, [IO.Compression.CompressionLevel]::Optimal, $true)
} finally {
	if (Test-Path $work) { [IO.Directory]::Delete($work, $true) }
}
$setup = Get-Item (Join-Path $releases "$name.exe")
$hash = (Get-FileHash $setup.FullName -Algorithm SHA256).Hash
"{0} ({1:N1} MB)" -f $setup.FullName, ($setup.Length / 1MB)
"SHA-256 $hash"
$portable = Get-Item (Join-Path $releases "$portableName.zip")
$portableHash = (Get-FileHash $portable.FullName -Algorithm SHA256).Hash
"{0} ({1:N1} MB)" -f $portable.FullName, ($portable.Length / 1MB)
"SHA-256 $portableHash"

if (-not $Publish) { return }

# the GitHub pre-release "AhkWord NN (installer preview)", tag installer-NN at this commit (moved there when the
# installer of the same iteration is made again after a later commit)
$tag = "installer-$iteration"
$title = "AhkWord $iteration (installer preview)"
$commit = git -C $root rev-parse --short HEAD
$notes = Join-Path ([IO.Path]::GetTempPath()) "ahkword-installer-notes.md"
Set-Content $notes -Encoding UTF8 @(
	"**A preview of AhkWord with an installer**, built from commit $commit. AhkWord.exe here is AhkWord compiled with Visual Studio's",
	"tools from the same sources as the AutoHotkey-based zips: it starts faster and needs neither AutoHotkey nor .NET installed.",
	"",
	"**Download** ``$($setup.Name)`` below and run it. It installs for your Windows user only (no administrator rights),",
	"in ``%LOCALAPPDATA%\Programs\AhkWord``, with a Start menu entry; Settings > Apps uninstalls it. Your settings, templates",
	"and recovery copies stay in ``%APPDATA%\AhkWord``, shared with the zips' AhkWord, and are kept when it is uninstalled.",
	"",
	"**Portable instead?** Download ``$($portable.Name)``, unzip it anywhere you can write to (a USB stick, your",
	"Documents) and run ``AhkWord.exe`` in it: nothing is installed. Its settings, recent files, templates and recovery",
	"copies stay in a ``data`` folder beside it (the empty ``portable`` file there says so), nothing goes into the registry",
	"or ``%APPDATA%``, and deleting the folder removes it completely. (The ``data`` folder appears the first time AhkWord",
	"saves something, e.g. when you close it, not at its first start.) Tip: right-click the zip > Properties > **Unblock**",
	"before unzipping. README-portable.txt inside tells more.",
	"",
	"**Already have AutoHotkey?** It is left alone. AhkWord brings its own private copy of AutoHotkey64.exe, kept in",
	"AhkWord's folder and used only by AhkWord: your installed AutoHotkey, its version, its ``.ahk`` file association and",
	"your scripts are not touched, and uninstalling AhkWord doesn't remove them. (While AhkWord runs, its own hotkeys",
	"and snippets are active, so a script of yours using the same hotkey may compete with it.)",
	"",
	"**Not signed yet.** Windows will warn about it:",
	"- *Windows protected your PC* (SmartScreen): click **More info**, then **Run anyway**.",
	"- **Smart App Control**, if it is on, blocks unsigned programs and offers no way past it. Use the zips of the latest",
	"  ""AhkWord NN"" pre-release instead, or wait for a signed installer.",
	"- Some antivirus programs distrust new unsigned installers.",
	"",
	"To check the downloads (PowerShell: ``Get-FileHash <file>``), SHA-256:",
	"- ``$($setup.Name)``: ``$hash``",
	"- ``$($portable.Name)``: ``$portableHash``",
	"",
	"Inside: AhkWord $iteration, .NET 10 (self-contained), AutoHotkey $ahkVersion (GPL v2, for the hotkeys, snippets and macros).")
try {
	git -C $root tag -f $tag HEAD | Out-Null
	git -C $root push -f origin "refs/tags/$tag" 2>&1 | Out-Host
	$exists = $true
	gh release view $tag --repo burque505/AhkWord *> $null
	if ($LASTEXITCODE -ne 0) { $exists = $false }
	if ($exists) {
		gh release upload $tag $setup.FullName $portable.FullName --clobber --repo burque505/AhkWord | Out-Host
		gh release edit $tag --title $title --notes-file $notes --prerelease --repo burque505/AhkWord | Out-Host
	} else {
		gh release create $tag $setup.FullName $portable.FullName --title $title --notes-file $notes --prerelease --latest=false --repo burque505/AhkWord | Out-Host
		# (the earlier installer previews: superseded)
		foreach ($old in gh release list --repo burque505/AhkWord --limit 100 --json tagName --jq '.[].tagName' | Where-Object { $_ -like "installer-*" -and $_ -ne $tag }) {
			$body = gh release view $old --repo burque505/AhkWord --json body --jq .body
			if ($body -notmatch '^Superseded by') {
				Set-Content $notes -Encoding UTF8 (@("Superseded by AhkWord $iteration (installer preview).", "") + $body)
				gh release edit $old --notes-file $notes --repo burque505/AhkWord | Out-Host
			}
		}
	}
	if ($LASTEXITCODE -ne 0) { throw "gh failed." }
	"Published: $title (https://github.com/burque505/AhkWord/releases/tag/$tag)"
} finally {
	Remove-Item $notes -ErrorAction SilentlyContinue
}
