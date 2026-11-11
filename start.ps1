# Starts the D4U SAM chatbot. Run from this folder: ./start.ps1
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path "config.json")) {
    Write-Host "Creating config.json from example..." -ForegroundColor Yellow
    Copy-Item "config.example.json" "config.json"
}

Write-Host "Installing dependencies (first run only)..." -ForegroundColor Cyan
pip install -q -r requirements.txt

Write-Host "Starting chatbot at http://localhost:8766 ..." -ForegroundColor Green
python app.py
