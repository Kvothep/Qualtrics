# Builds Int_NY_PCS_ETL_V02.py into a standalone Windows .exe with PyInstaller.
# Run from the project folder:  .\build.ps1

$ErrorActionPreference = "Stop"

# Prefer the author's conda env if present (so behavior on this machine is
# unchanged); otherwise fall back to whatever Python is on PATH, so this
# script also works on other machines (e.g. a corporate laptop) that just
# have a plain Python install. Override by setting $env:BUILD_PYTHON.
$pythonCandidates = @(
    $env:BUILD_PYTHON,
    "C:\Users\aleja\anaconda3\envs\automate\python.exe",
    $(try { (Get-Command python -ErrorAction Stop).Source } catch { $null }),
    $(try { (Get-Command python3 -ErrorAction Stop).Source } catch { $null })
) | Where-Object { $_ -and (Test-Path $_) }

if (-not $pythonCandidates) {
    throw "No Python interpreter found. Install Python 3.x (https://python.org) so 'python' is on PATH, or set `$env:BUILD_PYTHON to a python.exe path before running this script."
}
$python = $pythonCandidates[0]
Write-Host "Using Python: $python"

# The project folder may live inside OneDrive (or another sync client) on
# some machines. Sync agents lock the many small intermediate files
# PyInstaller writes, causing random PermissionError failures mid-build.
# Build in a local (non-synced) temp folder instead, then copy just the
# final .exe back into dist\.
$buildRoot = "$env:LOCALAPPDATA\PyInstallerBuilds\NY_PCS_ETL"
New-Item -ItemType Directory -Force -Path $buildRoot | Out-Null
$workPath = Join-Path $buildRoot "build"
$distPath = Join-Path $buildRoot "dist"
$specPath = $buildRoot

Write-Host "Installing/updating build dependencies..."
& $python -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw "pip install failed (exit $LASTEXITCODE)" }

Write-Host "Building executable (working outside OneDrive at $buildRoot)..."
# The 'automate' conda env is a full data-science env (scipy, matplotlib, Qt,
# Jupyter, dask, etc.) that this script never imports. PyInstaller's hooks pull
# a lot of that in automatically unless explicitly excluded, bloating a ~635MB
# exe down to a fraction of that.
$exclude = @(
    "matplotlib", "scipy", "PyQt5", "PySide2", "PySide6", "tkinter",
    "IPython", "ipykernel", "jupyter", "jupyterlab", "jupyter_client", "jupyter_core",
    "notebook", "nbformat", "nbconvert", "qtconsole",
    "sphinx", "docutils", "babel", "pygments",
    "dask", "distributed", "h5py", "tables", "bokeh", "xyzservices", "panel",
    "xarray", "patsy", "statsmodels", "numba", "llvmlite",
    "sqlalchemy", "pyarrow", "fsspec", "botocore", "lxml",
    "zmq", "nacl", "argon2", "anyio", "rich",
    "pytest", "py", "astroid", "jedi", "parso", "black", "blib2to3", "yapf_third_party",
    "win32com", "pythoncom"
)
$excludeArgs = $exclude | ForEach-Object { "--exclude-module=$_" }

& $python -m PyInstaller `
    --onefile `
    --noconsole `
    --name "NY_PCS_ETL" `
    --clean `
    --workpath $workPath `
    --distpath $distPath `
    --specpath $specPath `
    @excludeArgs `
    Int_NY_PCS_ETL_V02.py

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed (exit $LASTEXITCODE) -- see output above."
}

New-Item -ItemType Directory -Force -Path "dist" | Out-Null
Copy-Item (Join-Path $distPath "NY_PCS_ETL.exe") "dist\NY_PCS_ETL.exe" -Force

Write-Host ""
Write-Host "Done. Executable is at dist\NY_PCS_ETL.exe"
$sizeMB = [math]::Round((Get-Item "dist\NY_PCS_ETL.exe").Length / 1MB, 1)
Write-Host "Size: $sizeMB MB"
Write-Host "Credentials are hardcoded in Int_NY_PCS_ETL_V02.py and now baked into this exe."
Write-Host "No .env file needed -- the exe runs standalone. See HOW_TO_RUN.txt for details."
