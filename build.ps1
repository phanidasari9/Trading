# Create/update .venv and install dependencies (pyproject + requirements stay in sync).
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
Set-Location $root

$venvPy = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $venvPy)) {
    Write-Host "Creating .venv ..."
    python -m venv .venv
}
& $venvPy -m pip install --upgrade pip
& $venvPy -m pip install -e .
Write-Host "Build finished. Run: .\run.ps1   or   .\.venv\Scripts\python.exe -m streamlit run dashboard.py"
