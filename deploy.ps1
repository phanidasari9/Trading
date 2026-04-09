# Build Docker image (and optionally start with docker compose).
param(
    [switch]$Up
)
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "Building trading-dashboard image..."
docker build -t trading-dashboard:latest .

if ($Up) {
    Write-Host "Starting stack..."
    docker compose up -d
    Write-Host "Open http://localhost:8501"
} else {
    Write-Host "Run: docker compose up -d   OR   docker run --rm -p 8501:8501 trading-dashboard:latest"
}
