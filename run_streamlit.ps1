$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$venvActivate = Join-Path $root ".venv\Scripts\Activate.ps1"
if (-not (Test-Path $venvActivate)) {
    Write-Error "Missing .venv. Create it first (python -m venv .venv)."
}

& $venvActivate

python -m pip install -r requirements.txt
python -m playwright install chromium

streamlit run app\streamlit_app.py
