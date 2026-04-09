# Start Streamlit dashboard (expects build.ps1 or manual venv + pip install -e .).
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
Set-Location $root

$venvPy = Join-Path $root ".venv\Scripts\python.exe"
if (Test-Path $venvPy) {
    & $venvPy -m streamlit run (Join-Path $root "dashboard.py")
} else {
    Write-Host "No .venv found. Run .\build.ps1 first, or: python -m streamlit run dashboard.py"
    python -m streamlit run (Join-Path $root "dashboard.py")
}
