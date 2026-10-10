param(
    # Falls back to the product name in version.py. Override only if you need
    # a differently-named artifact (e.g. a one-off test build).
    [string]$AppName
)

$ErrorActionPreference = 'Stop'

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path

# Resolve a real Python interpreter. Two traps this avoids:
#   * The Windows Store alias stub (...\WindowsApps\python.exe) is a shim that
#     opens the Store instead of running, so it is never selected.
#   * The original build used a Codex-bundled 3.12 interpreter under
#     %USERPROFILE%\.cache\codex-runtimes\. That cache is being retired, so it
#     is no longer referenced. Set $env:UNITRANS_PYTHON to pin an interpreter.
$PythonExe = $null
$PythonPrefix = @()

if ($env:UNITRANS_PYTHON) {
    $PythonExe = $env:UNITRANS_PYTHON
} else {
    $launcher = Get-Command py -ErrorAction SilentlyContinue
    if ($launcher) {
        $PythonExe = $launcher.Source
        $PythonPrefix = @('-3')
    } else {
        foreach ($name in @('python3', 'python')) {
            $candidate = Get-Command $name -ErrorAction SilentlyContinue
            if ($candidate -and $candidate.Source -notmatch '\\WindowsApps\\') {
                $PythonExe = $candidate.Source
                break
            }
        }
        if (-not $PythonExe) {
            $localPython = Join-Path $env:LOCALAPPDATA 'Python\bin\python.exe'
            if (Test-Path -LiteralPath $localPython) { $PythonExe = $localPython }
        }
    }
}

if (-not $PythonExe) {
    throw 'No usable Python found. Install Python 3.12+ from python.org, or set $env:UNITRANS_PYTHON to the full path of python.exe.'
}

$PythonVersion = & $PythonExe @PythonPrefix -c "import sys; print('%d.%d.%d' % sys.version_info[:3])"
if ($LASTEXITCODE -ne 0 -or -not $PythonVersion) {
    throw "Python at '$PythonExe' did not run. Set `$env:UNITRANS_PYTHON to a working interpreter."
}
Write-Host "Using Python $PythonVersion ($PythonExe)"

$Venv = Join-Path $Root '.venv'
$VenvPython = Join-Path $Venv 'Scripts\python.exe'
$VenvOk = Test-Path $VenvPython
if ($VenvOk) {
    & $VenvPython -c "import sys" *> $null
    $VenvOk = $LASTEXITCODE -eq 0
}
if (-not $VenvOk) {
    if (Test-Path $Venv) {
        Remove-Item -LiteralPath $Venv -Recurse -Force
    }
    & $PythonExe @PythonPrefix -m venv $Venv
}

& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r (Join-Path $Root 'requirements.txt')

# ---------------------------------------------------------------------------
# Version and product identity come from version.py - the single source of
# truth. Nothing here re-declares a version number, so the exe metadata, the
# artifact name and the in-app footer cannot drift apart.
# ---------------------------------------------------------------------------
$BuildInfoJson = & $VenvPython (Join-Path $Root 'version.py')
if ($LASTEXITCODE -ne 0) { throw 'version.py did not run.' }
$BuildInfo = $BuildInfoJson | ConvertFrom-Json

$Version = $BuildInfo.version
$FileVersion = $BuildInfo.file_version
$Publisher = $BuildInfo.publisher
$ProductName = $BuildInfo.app_name
if (-not $AppName) { $AppName = $ProductName }

Write-Host "Building $ProductName $Version (file version $FileVersion, publisher $Publisher)"

$BuildDir = Join-Path $Root 'build'
if (-not (Test-Path $BuildDir)) { New-Item -ItemType Directory -Path $BuildDir | Out-Null }

# --- icon ------------------------------------------------------------------
$IconPath = Join-Path $Root 'assets\app.ico'
$AssetScript = Join-Path $Root 'make_assets.py'
if (Test-Path $AssetScript) {
    & $VenvPython $AssetScript
    if ($LASTEXITCODE -ne 0) { throw 'make_assets.py failed - the icon could not be generated.' }
}
if (-not (Test-Path $IconPath)) {
    throw "Icon not found at '$IconPath'. Run make_assets.py to generate it from the logo."
}

# --- Windows version resource ---------------------------------------------
# Without this the exe has blank FileVersion/ProductName, so Explorer cannot
# tell two builds apart.
$VersionFile = Join-Path $BuildDir 'version_info.txt'
& $VenvPython (Join-Path $Root 'version.py') --pyinstaller-version-file |
    Set-Content -LiteralPath $VersionFile -Encoding utf8
if (-not (Test-Path $VersionFile)) { throw 'Failed to generate the version resource file.' }

$PyInstallerArgs = @(
    '-m', 'PyInstaller',
    '--noconfirm',
    '--clean',
    '--onefile',
    '--windowed',
    '--name', $AppName,
    '--icon', $IconPath,
    '--version-file', $VersionFile,
    '--additional-hooks-dir', (Join-Path $Root 'pyinstaller_hooks'),
    '--hidden-import', 'tkinter',
    '--hidden-import', '_tkinter'
)

$SitePackages = & $VenvPython -c "import site; print(site.getsitepackages()[0])"
$RuntimeHook = Join-Path $SitePackages 'PyInstaller\hooks\rthooks\pyi_rth__tkinter.py'
if (Test-Path $RuntimeHook) {
    $PyInstallerArgs += @('--runtime-hook', $RuntimeHook)
}

$PythonBase = & $VenvPython -c "import sys; print(sys.base_prefix)"
$TclData = Join-Path $PythonBase 'tcl\tcl8.6'
$TkData = Join-Path $PythonBase 'tcl\tk8.6'
$TclDll = Join-Path $PythonBase 'DLLs\tcl86t.dll'
$TkDll = Join-Path $PythonBase 'DLLs\tk86t.dll'
if ((Test-Path $TclData) -and (Test-Path $TkData)) {
    $PyInstallerArgs += @('--add-data', "$TclData;_tcl_data", '--add-data', "$TkData;_tk_data")
}
if ((Test-Path $TclDll) -and (Test-Path $TkDll)) {
    $PyInstallerArgs += @('--add-binary', "$TclDll;.", '--add-binary', "$TkDll;.")
}

$Logo = Join-Path $Root 'UniTrans Logo.png'
if (Test-Path $Logo) {
    $PyInstallerArgs += @('--add-data', "$Logo;.")
}

# Release notes travel inside the exe so the "What's New" dialog works with no
# network access.
$ChangelogJson = Join-Path $Root 'changelog.json'
if (Test-Path $ChangelogJson) {
    $PyInstallerArgs += @('--add-data', "$ChangelogJson;.")
}

# version.py is imported by waybill_sorter.py, so it must be importable inside
# the frozen bundle. It is a top-level module next to the entry script, which
# PyInstaller picks up automatically, but listing it makes that explicit.
$PyInstallerArgs += @('--hidden-import', 'version')

$PyInstallerArgs += (Join-Path $Root 'waybill_sorter.py')
& $VenvPython @PyInstallerArgs
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller failed.' }

$BuiltExe = Join-Path $Root "dist\$AppName.exe"
if (-not (Test-Path $BuiltExe)) { throw "Build reported success but '$BuiltExe' does not exist." }

# ---------------------------------------------------------------------------
# Name the release artifact with its version. Two builds with different
# versions must never share a filename, or "which one is current?" becomes a
# question about file timestamps again.
# ---------------------------------------------------------------------------
$ReleaseDir = Join-Path $Root 'releases'
if (-not (Test-Path $ReleaseDir)) { New-Item -ItemType Directory -Path $ReleaseDir | Out-Null }

$SafeName = ($ProductName -replace '\s+', '-')
$ReleaseExe = Join-Path $ReleaseDir "$SafeName-$Version.exe"
Copy-Item -LiteralPath $BuiltExe -Destination $ReleaseExe -Force

$Hash = Get-FileHash -LiteralPath $ReleaseExe -Algorithm SHA256
$SizeBytes = (Get-Item -LiteralPath $ReleaseExe).Length
$SizeMb = [math]::Round($SizeBytes / 1MB, 1)

Write-Host ''
Write-Host 'Build complete' -ForegroundColor Green
Write-Host "  product      : $ProductName $Version"
Write-Host "  publisher    : $Publisher"
Write-Host "  dist         : $BuiltExe"
Write-Host "  release      : $ReleaseExe"
Write-Host "  size         : $SizeMb MB ($SizeBytes bytes)"
Write-Host "  sha256       : $($Hash.Hash)"
Write-Host ''
Write-Host 'Verify the exe metadata before publishing:'
Write-Host "  (Get-Item '$ReleaseExe').VersionInfo | Format-List FileVersion,ProductName,CompanyName"
