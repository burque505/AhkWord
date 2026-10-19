# Makes a runnable copy of an iteration of RTE Writer, to run it again later, in ..\rte-versions (beside this folder,
# not inside it: nothing done to this folder or its git repository can touch the copies).
#   .\make_version.ps1 33          iteration 33 as it was committed ("Iteration 33: ..."): rte-versions\iteration-33
#   .\make_version.ps1 -All        every committed iteration from 24 on (those already there are left as they are)
#   .\make_version.ps1 -Working    the files as they are now, committed or not: rte-versions\iteration-NN-working-<date>
#                                  (a new one each time; the ones before stay)
#   .\make_version.ps1 -List       the copies there are
# Each copy is complete and stands alone: the iteration's own rte_writerNN.ahk/.cs, its modules (they are edited in
# place from iteration to iteration, so an old rte_writerNN.cs runs only with the modules of its own time), the
# ribbon, macros, templates and examples, exported from git (no link to the repository is left). Its files are made
# read-only, so that they can't be changed by mistake; a "portable" file makes it keep its settings, recovery copies
# and converters' state in a data folder of its own, so that it never changes the current version's (in %APPDATA% and
# %LOCALAPPDATA%). VERSION.txt says what it is. Run a copy with its rte_writerNN.ahk.
param([int]$Iteration = 0, [switch]$All, [switch]$Working, [switch]$List)
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$versions = Join-Path (Split-Path $root) "rte-versions"
New-Item -ItemType Directory -Force $versions | Out-Null

function Commits {
	# the commits named "Iteration NN: ...", newest first: NN -> hash
	$map = @{}
	foreach ($line in git -C $root log --format="%H %s") {
		if ($line -match '^([0-9a-f]{40}) Iteration (\d+):' -and -not $map.ContainsKey([int]$Matches[2])) { $map[[int]$Matches[2]] = $Matches[1] }
	}
	$map
}

# The copy finished: what it is, its own data folder, its program files read-only.
function Finish($folder, [string]$what) {
	$script = Get-ChildItem $folder -Filter "rte_writer*.ahk" | Sort-Object { [int]($_.BaseName -replace '\D', '') } | Select-Object -Last 1
	Set-Content (Join-Path $folder "VERSION.txt") @(
		$what,
		"Made $(Get-Date -Format 'yyyy-MM-dd HH:mm') from $root.",
		"Run it with $($script.Name). It keeps its settings in the data folder here (portable mode), not yours.")
	New-Item -ItemType File -Force (Join-Path $folder "portable") | Out-Null
	New-Item -ItemType Directory -Force (Join-Path $folder "data") | Out-Null
	Get-ChildItem $folder -Recurse -File | Where-Object { $_.FullName -notlike (Join-Path $folder "data\*") } | ForEach-Object { $_.IsReadOnly = $true }
	"$what -> $folder\$($script.Name)"
}

function MakeCommitted([int]$n, [string]$hash) {
	$folder = Join-Path $versions ("iteration-{0:D2}" -f $n)
	if (Test-Path $folder) { "iteration $n is there already: $folder"; return }
	$zip = Join-Path $env:TEMP "rte-version-$n.zip"
	git -C $root archive --format=zip -o $zip $hash
	Expand-Archive $zip $folder
	Remove-Item $zip
	$subject = git -C $root log -1 --format="%s (%ad, commit %h)" --date=short $hash
	Finish $folder $subject
}

if ($List) {
	Get-ChildItem $versions -Directory | Sort-Object Name | ForEach-Object {
		$about = Join-Path $_.FullName "VERSION.txt"
		"{0}: {1}" -f $_.Name, ($(if (Test-Path $about) { (Get-Content $about -TotalCount 1) } else { "?" }))
	}
	return
}

if ($Working) {
	$script = Get-ChildItem $root -Filter "rte_writer*.ahk" | Sort-Object { [int]($_.BaseName -replace '\D', '') } | Select-Object -Last 1
	$n = [int]($script.BaseName -replace '\D', '')
	$stamp = Get-Date -Format "yyyyMMdd-HHmm"
	$folder = Join-Path $versions ("iteration-{0:D2}-working-{1}" -f $n, $stamp)
	if (Test-Path $folder) { "There is a copy from this minute already: $folder"; return }
	# the files git has or would have (tracked, and new ones not ignored), as they are now
	$files = @(git -C $root ls-files) + @(git -C $root ls-files --others --exclude-standard) | Sort-Object -Unique
	foreach ($file in $files) {
		$from = Join-Path $root $file
		if (-not (Test-Path $from -PathType Leaf)) { continue }    # (deleted, not yet committed)
		$to = Join-Path $folder $file
		New-Item -ItemType Directory -Force (Split-Path $to) | Out-Null
		Copy-Item $from $to
	}
	$head = git -C $root log -1 --format="%h"
	Finish $folder "Iteration $n as it was on $(Get-Date -Format 'yyyy-MM-dd HH:mm'), not committed (after commit $head)"
	return
}

$commits = Commits
if ($All) {
	foreach ($n in ($commits.Keys | Where-Object { $_ -ge 24 } | Sort-Object)) { MakeCommitted $n $commits[$n] }
} elseif ($Iteration -gt 0) {
	if (-not $commits.ContainsKey($Iteration)) { throw "No commit named ""Iteration $($Iteration): ..."" (committed ones: $(($commits.Keys | Sort-Object) -join ', '))." }
	MakeCommitted $Iteration $commits[$Iteration]
} else {
	"Use: .\make_version.ps1 33 | -All | -Working | -List"
}
