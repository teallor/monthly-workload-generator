$ErrorActionPreference = "Stop"
$project = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPython = Join-Path $project ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $venvPython)) {
    python -m venv (Join-Path $project ".venv")
}

& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r (Join-Path $project "requirements-build.txt")
& $venvPython -m PyInstaller --noconfirm --clean (Join-Path $project "gui_app.spec")

$exe = Get-ChildItem -LiteralPath (Join-Path $project "dist") -Filter "*.exe" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1 -ExpandProperty FullName
if (-not $exe) { throw "Build failed: executable not found" }
Write-Host "Build completed: $exe"
