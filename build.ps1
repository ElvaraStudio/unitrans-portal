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
$InstallerName = "$SafeName-$Version-Setup.exe"
$ReleaseInstaller = Join-Path $ReleaseDir $InstallerName

# A running copy of the app holds its own .exe open, so the copy below fails
# with a bare "being used by another process". Fail early with something
# actionable instead. Note the PyInstaller onefile child process shows up under
# the artifact's name, not under the bootloader's, so match on the app name.
$Running = Get-Process -Name $SafeName -ErrorAction SilentlyContinue
if ($Running) {
    $ids = ($Running | Select-Object -ExpandProperty Id) -join ', '
    throw "Unitrans Portal is currently running (PID $ids). Close it before building - a running copy locks its own executable."
}

Copy-Item -LiteralPath $BuiltExe -Destination $ReleaseExe -Force

# Keep this folder unambiguous: it should hold exactly one build, the one just
# produced. Stale artifacts are what caused the mess this release cleans up -
# two exes, neither self-describing, and no way to tell which was current.
# Published builds live on the GitHub Release page, not here.
$Stale = Get-ChildItem -LiteralPath $ReleaseDir -Filter '*.exe' -File |
    Where-Object { $_.FullName -ne $ReleaseExe }
foreach ($old in $Stale) {
    # Recycle Bin rather than a hard delete, so a mistaken cleanup is recoverable.
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
            $old.FullName,
            [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
            [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
        Write-Host "Removed superseded build: $($old.Name)" -ForegroundColor Yellow
    } catch {
        Write-Warning "Could not remove stale build '$($old.Name)': $($_.Exception.Message)"
    }
}

# Build a per-user installer when Inno Setup is installed. The portable exe
# remains available as a separate artifact for users who prefer no installer.
$InnoCommand = Get-Command iscc.exe -ErrorAction SilentlyContinue
$InnoCompilerPath = if ($InnoCommand) { $InnoCommand.Source } else { $null }
if (-not $InnoCompilerPath) {
    foreach ($candidate in @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'),
        (Join-Path $env:ProgramFiles 'Inno Setup 6\ISCC.exe')
    )) {
        if (Test-Path -LiteralPath $candidate) { $InnoCompilerPath = $candidate; break }
    }
}
$ReleaseInstaller = $null
$DownloadsInstaller = $null
if ($InnoCompilerPath) {
    Push-Location $Root
    try {
        & $InnoCompilerPath "/DAppVersion=$Version" (Join-Path $Root 'installer.iss')
        if ($LASTEXITCODE -ne 0) { throw 'Inno Setup failed to build the installer.' }
    } finally {
        Pop-Location
    }
    if (-not (Test-Path -LiteralPath (Join-Path $ReleaseDir $InstallerName))) {
        throw 'Inno Setup reported success but the installer was not created.'
    }
    $ReleaseInstaller = Join-Path $ReleaseDir $InstallerName
} else {
    Write-Warning 'Inno Setup was not found; the portable exe was built, but the installer was skipped. Install Inno Setup 6 and rerun build.ps1.'
}

$Hash = Get-FileHash -LiteralPath $ReleaseExe -Algorithm SHA256
$SizeBytes = (Get-Item -LiteralPath $ReleaseExe).Length
$SizeMb = [math]::Round($SizeBytes / 1MB, 1)
$InstallerHash = $null
$InstallerSizeMb = $null
if ($ReleaseInstaller) {
    $InstallerHash = Get-FileHash -LiteralPath $ReleaseInstaller -Algorithm SHA256
    $InstallerSizeMb = [math]::Round((Get-Item -LiteralPath $ReleaseInstaller).Length / 1MB, 1)
}

# ---------------------------------------------------------------------------
# Drop a copy in the Downloads folder for testing.
#
# Publishing to GitHub and then downloading the result just to try a build is a
# waste of time, so the built exe is left where it is easiest to grab. This
# copy is NOT the artifact of record - releases\ is - and it is never committed
# (the whole releases/ binary set is gitignored). Delete it any time.
# ---------------------------------------------------------------------------
$DownloadsCopy = $null
$DownloadsInstaller = $null
$DownloadsDir = Join-Path $env:USERPROFILE 'Downloads'
if (Test-Path -LiteralPath $DownloadsDir) {
    $DownloadsCopy = Join-Path $DownloadsDir "$SafeName-$Version.exe"
    try {
        Copy-Item -LiteralPath $ReleaseExe -Destination $DownloadsCopy -Force
    } catch {
        Write-Warning "Could not copy to Downloads: $($_.Exception.Message)"
        $DownloadsCopy = $null
    }
    if ($ReleaseInstaller) {
        $DownloadsInstaller = Join-Path $DownloadsDir $InstallerName
        try {
            Copy-Item -LiteralPath $ReleaseInstaller -Destination $DownloadsInstaller -Force
        } catch {
            Write-Warning "Could not copy installer to Downloads: $($_.Exception.Message)"
            $DownloadsInstaller = $null
        }
    }
}

Write-Host ''
Write-Host 'Build complete' -ForegroundColor Green
Write-Host "  product      : $ProductName $Version"
Write-Host "  publisher    : $Publisher"
Write-Host "  dist         : $BuiltExe"
Write-Host "  release      : $ReleaseExe"
if ($DownloadsCopy) {
    Write-Host "  for testing  : $DownloadsCopy" -ForegroundColor Cyan
}
if ($ReleaseInstaller) {
    Write-Host "  installer    : $ReleaseInstaller"
    if ($DownloadsInstaller) { Write-Host "  share copy   : $DownloadsInstaller" -ForegroundColor Cyan }
    Write-Host "  setup size   : $InstallerSizeMb MB"
    Write-Host "  setup sha256 : $($InstallerHash.Hash)"
}
Write-Host "  size         : $SizeMb MB ($SizeBytes bytes)"
Write-Host "  sha256       : $($Hash.Hash)"
Write-Host ''
Write-Host 'Verify the exe metadata before publishing:'
Write-Host "  (Get-Item '$ReleaseExe').VersionInfo | Format-List FileVersion,ProductName,CompanyName"
