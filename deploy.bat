@echo off
cd /d "%~dp0"
echo Building trading-dashboard image...
docker build -t trading-dashboard:latest .
if errorlevel 1 exit /b 1
echo.
echo Done. Run: docker compose up -d
echo    or:  docker run --rm -p 8501:8501 trading-dashboard:latest
echo Open: http://localhost:8501
