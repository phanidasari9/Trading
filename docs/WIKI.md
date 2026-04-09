# Trading dashboard — Wiki

> **Using this on GitHub:** In your repo on GitHub, open **Wiki** → create or edit the **Home** page and paste the sections below (or keep this file in-repo as `docs/WIKI.md`).

---

## Overview

Streamlit app: **Market movers**, **Options** (`Options.py`), **Swing options** (`OptionSwing.py`), **BestOptions** (`BestWiningOptionsv1.py`), **Flow** (`Flow.py`). Data via **yfinance** — no API keys required for basic use.

| Item | Value |
|------|--------|
| **Main file (Streamlit Cloud)** | `dashboard.py` |
| **Python** | `3.12` (see `runtime.txt`) |
| **Dependencies** | `requirements.txt` (primary on Cloud; `pyproject.toml` may show a duplicate warning — safe to ignore) |

---

## Run locally

```powershell
cd c:\Phani\Cursor\Trading
.\build.ps1          # optional: .venv + pip install -e .
.\run.ps1            # or: python -m streamlit run dashboard.py
```

With a fixed port:

```powershell
python -m streamlit run dashboard.py --server.headless true --server.port 8501
```

---

## Deploy with Docker

```powershell
.\deploy.ps1         # build image
.\deploy.ps1 -Up     # build + docker compose up
```

Open **http://localhost:8501**. Health: `GET http://localhost:8501/_stcore/health`

**Note:** `watchlists.json` in a plain container is ephemeral unless you mount a volume.

---

## Streamlit Community Cloud

### Connect the app

1. Push this repository to **GitHub**.
2. Go to [share.streamlit.io](https://share.streamlit.io) → sign in with GitHub.
3. **New app** → select repo and branch (**`main`**).
4. **Main file path:** `dashboard.py`
5. Deploy.

Cloud installs from **`requirements.txt`** and uses **`runtime.txt`** for Python 3.12.

### After you push code

- A **`git push`** to the tracked branch usually triggers a **new deploy** within a minute or two.
- Use **Manage app → Reboot** if the app is stuck, didn’t pick up the latest commit, or after changing **Secrets** / settings.

### Configuration in this repo (Cloud fixes)

| File | Purpose |
|------|--------|
| `.streamlit/credentials.toml` | Empty `email` under `[general]` — skips the interactive “Welcome to Streamlit!” prompt so the server starts and **health checks** (`/healthz`) succeed. |
| `.streamlit/config.toml` | `server.headless = true`, theme, `browser.gatherUsageStats = false`. |
| `watchlists_store.py` | On **read-only** app dirs (Cloud), reads bundled `watchlists.json` from the repo and **writes** saves under **system temp** (not persistent across full redeploys). |
| `requirements.txt` | **`pandas>=2.0.0,<3.0.0`** for compatibility with Streamlit stacks. |

### Watchlists on Cloud

- **Default universe** loads from **`watchlists.json`** in the repo (read-only).
- **Save / create list** writes to temp in Cloud — treat as **session-like** unless you add external storage later.

### Build log messages

- **“More than one requirements file”** — informational; Cloud still uses **`requirements.txt`** with uv.
- **“Streamlit is already installed”** — normal.

### If the app shows “Oh no” or health check fails

1. Open **Manage app → Logs** and read the **traceback**.
2. Typical causes we addressed:
   - **Onboarding / email prompt** blocking startup → fixed with **`.streamlit/credentials.toml`** and **headless** server.
   - **Permission denied** writing `watchlists.json` next to the app → fixed with **writable temp** in `watchlists_store.py`.
   - **pandas 3.x** issues → pinned **pandas &lt; 3** in `requirements.txt`.

---

## Git: remote and check-in

```powershell
cd c:\Phani\Cursor\Trading
git remote add origin https://github.com/USERNAME/Trading.git   # once
git branch -M main
git push -u origin main
```

Routine updates:

```powershell
git add -A
git status
git commit -m "Your message"
git push origin main
```

Set identity if Git asks:

```powershell
git config user.name "Your Name"
git config user.email "you@example.com"
```

For HTTPS push to GitHub, use a **Personal Access Token** as the password (not your GitHub account password).

---

## Project layout (high level)

| Path | Role |
|------|------|
| `dashboard.py` | Streamlit UI, tabs, session state |
| `market_analysis.py` | Equity movers, Excel export helpers |
| `Options.py` | Options flow analyzer + HTML report |
| `OptionSwing.py` / `BestWiningOptionsv1.py` | Swing / best-options scans |
| `Flow.py` | Money-flow scan + Excel bytes for dashboard |
| `watchlists_store.py` / `watchlists.json` | Named ticker lists |
| `trading_ui.py` | Plotly defaults + light primary-button CSS |
| `.streamlit/` | `config.toml`, `credentials.toml` (no secrets in repo) |

**Secrets:** Put API keys only in **Streamlit Cloud → Secrets** or local `.streamlit/secrets.toml` (gitignored). This app does not require them for yfinance-only use.

---

## Optional: PNG export (Options tab)

Playwright is **not** in the default `requirements.txt`. For local PNG export:

```powershell
pip install playwright
python -m playwright install chromium
```

---

*Last aligned with repo fixes: Streamlit Cloud onboarding, read-only watchlists, pandas pin, `credentials.toml`.*
