# Build and deploy — Trading dashboard

Streamlit app: **Market movers**, **Options** (`Options.py`), **Swing options** (`OptionSwing.py`). No API keys required (yfinance).

## Build locally (no Docker)

```powershell
.\build.ps1
.\run.ps1
```

## Build with Docker

```powershell
.\deploy.ps1
```

Run the container:

```powershell
docker run --rm -p 8501:8501 trading-dashboard:latest
```

Or Compose (same image, port 8501):

```powershell
.\deploy.ps1 -Up
# or: docker compose up --build -d
```

Open **http://localhost:8501**.

Health check: `GET http://localhost:8501/_stcore/health`

### Notes

- **Watchlists** (`watchlists.json`) live next to the app. In a plain container they are **ephemeral** (reset when the container is recreated). Mount a volume or file if you need persistence.
- **Playwright / PNG** from the Options tab is not in the default image; install optional deps in a custom image if needed.

## Streamlit Community Cloud

1. Push this folder to a GitHub repo (keep `Options.py` / `OptionSwing.py` casing for Linux).
2. [share.streamlit.io](https://share.streamlit.io) → New app → select repo.
3. **Main file:** `dashboard.py`
4. **Python:** 3.12 (optional: `runtime.txt` is provided).

## Heroku / Railway-style (Procfile)

- **Start command** is in `Procfile` (`$PORT` is set by the platform).
- Ensure **Python buildpack** installs `requirements.txt`, then run the Procfile `web` process.

## Render / Fly.io / cloud VM

- **Docker:** use the included `Dockerfile`; set the service **port** to **8501** (or map host port to 8501).
- **HTTPS** is usually terminated by the platform reverse proxy in front of the container.
