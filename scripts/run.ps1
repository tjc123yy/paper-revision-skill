# paper-revision skill - Python interpreter resolver and runner.
#
# This environment has no python on PATH (only the WindowsApps Store alias) and
# lxml exists only inside standalone runtimes, so this helper probes a candidate
# list, verifies that `import lxml` works, then runs the target script.
#
# ASCII-ONLY ON PURPOSE: Windows PowerShell 5.1 reads a BOM-less .ps1 file using
# the system ANSI code page (GBK on a Chinese Windows). Non-ASCII text here can
# then swallow quote characters and break parsing. Keep Chinese documentation in
# SKILL.md instead of in this file.
#
# Usage (no param block, so flags pass through verbatim):
#   .\run.ps1 inspect_docx.py "..\..\student\paper.docx"
#   .\run.ps1 verify_revision.py "paper_fixed.docx" --config task.json
#
# Set DSH_PAPER_PYTHON to force a specific interpreter.

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

if ($args.Count -lt 1) {
    Write-Host 'Usage: run.ps1 <script.py> [args...]'
    Write-Host 'Example: run.ps1 verify_revision.py paper.docx --config task.json'
    exit 2
}

$scriptPath = $args[0]
if (-not (Test-Path -LiteralPath $scriptPath)) {
    $cand = Join-Path $here $scriptPath
    if (Test-Path -LiteralPath $cand) { $scriptPath = $cand }
}
if (-not (Test-Path -LiteralPath $scriptPath)) {
    Write-Host ("Script not found: " + $args[0])
    exit 2
}
$scriptPath = (Resolve-Path -LiteralPath $scriptPath).Path

$rest = @()
if ($args.Count -gt 1) { $rest = $args[1..($args.Count - 1)] }

# Third-party site-packages that ship a matching lxml (cp312 / cp313 build).
$extraSites = @()
$runtimeSite = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\Lib\site-packages'
$workbuddySite = Join-Path $env:USERPROFILE '.workbuddy\binaries\python\envs\default\Lib\site-packages'
foreach ($s in @($runtimeSite, $workbuddySite)) {
    if (Test-Path -LiteralPath $s) { $extraSites += $s }
}

$candidates = @()

function Add-Candidate([string]$exe, [string]$pyPath) {
    if ($exe -and (Test-Path -LiteralPath $exe)) {
        $script:candidates += [pscustomobject]@{ Exe = $exe; PyPath = $pyPath }
    }
}

# 1. explicit override
if ($env:DSH_PAPER_PYTHON) { Add-Candidate $env:DSH_PAPER_PYTHON '' }

# 2. python on PATH, if it actually works
$onPath = Get-Command python -ErrorAction SilentlyContinue
if ($onPath) { Add-Candidate $onPath.Source '' }

# 3. standalone runtime that already ships lxml (preferred: no PYTHONPATH needed)
$runtimePy = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
Add-Candidate $runtimePy ''

# 4. uv-installed interpreters, with a neighbouring lxml site-packages
$uvRoot = Join-Path $env:APPDATA 'uv\python'
if (Test-Path -LiteralPath $uvRoot) {
    $uv = Get-ChildItem -LiteralPath $uvRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending |
        ForEach-Object { Join-Path $_.FullName 'python.exe' }
    foreach ($exe in $uv) {
        Add-Candidate $exe ''
        foreach ($site in $extraSites) { Add-Candidate $exe $site }
    }
}

# 5. interpreters bundled with other tools
$wb = Join-Path $env:USERPROFILE '.workbuddy\binaries\python\envs\default\Scripts\python.exe'
Add-Candidate $wb ''
Add-Candidate 'C:\Program Files\LibreOffice\program\python.exe' ''

if ($candidates.Count -eq 0) {
    Write-Host 'No candidate interpreter found. Set DSH_PAPER_PYTHON to one.'
    exit 3
}

$chosen = $null
foreach ($c in $candidates) {
    $env:PYTHONPATH = $c.PyPath
    try {
        & $c.Exe -c 'import lxml, sys; sys.exit(0)' 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { $chosen = $c; break }
    } catch {
        continue
    }
}

if (-not $chosen) {
    Write-Host 'Found interpreters but none can import lxml. Candidates:'
    foreach ($c in $candidates) {
        Write-Host ("  " + $c.Exe + "  PYTHONPATH=" + $c.PyPath)
    }
    exit 3
}

Write-Host ("[run.ps1] python = " + $chosen.Exe)
if ($chosen.PyPath) { Write-Host ("[run.ps1] PYTHONPATH = " + $chosen.PyPath) }

$env:PYTHONPATH = $chosen.PyPath
& $chosen.Exe -X utf8 $scriptPath @rest
exit $LASTEXITCODE
