# Outlet Info Scraper

Automated CLI tool for extracting and managing merchant information from **Grab Merchant Portal** and **Shopee Food Partner**. Built with Python, Playwright, and Pandas.

## 🚀 Features

- **Grab Merchant Scraper**:
  - Headless/Automated Login handling.
  - Extracts essential outlet data (Portal, Nama, Group ID, Store ID, Status, Alamat, Bank Account) across multiple credentials.
  - Rapid data extraction using direct API queries (`merchant-selector` and `store/search`).
  - Automated data completeness validation (compares extracted count vs Grab's system total).
- **Shopee Food Scraper**:
  - Automated extraction of Shopee Food merchant lists.
- **Interactive CLI (`cli.py`)**:
  - Easy-to-use terminal interface to select specific platforms (Grab/Shopee) or individual outlets.
  - Built-in utilities for Excel data combination and duplicate checking.

## 🛠️ Prerequisites

- Python 3.9+
- `uv` (Python package manager)
- Playwright Chromium

## 📦 Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/SuperfoodTech/OutletInfo.git
   cd OutletInfo
   ```

2. Setup environment and install Playwright browsers:
   ```bash
   # Make sure dependencies in pyproject.toml / uv are installed
   playwright install chromium
   ```

## 🎮 Usage

Run the main CLI interface:

```bash
uv run cli.py
```

### Main Menu Options
1. **🟢 Grab**: Run Grab scraper, combine excel files, find/remove duplicates.
2. **🟠 Shopee**: Run Shopee scraper, check output status.
3. **🔴 GoFood**: Run GoFood scraper, check output status.
4. **🚀 Jalankan Semua**: Run Grab, Shopee, and GoFood scrapers sequentially.

## 📁 Project Structure

- `cli.py`: Main interactive command-line interface.
- `GRAB/`: Contains Grab-specific scripts (`grab_merchant_scraper.py`, utilities for deduplication and combination).
- `SHOPEE/`: Contains Shopee-specific scripts (`shopee_scraper.py`).
- Output directories (`GRAB/hasil_custom/` & `SHOPEE/data/`) are used for generated `.xlsx` reports.

## 🔒 Security Note
Do not commit sensitive files containing credentials. They are properly handled by `.gitignore`.

## 🌿 Branching & Versioning Strategy

This repository follows a structured branching and versioning model:

### Branches
- `main` : Production-ready code.
- `staging` : Code being tested before production.
- `feature/xxx` : New feature or task (e.g., `feature/grabfood-menu-data`).
- `fix/xxx` : Bug fix (e.g., `fix/grabfood-menu-cron`).

*Feature and fix branches should be merged into `staging` for testing, and then into `main`.*

### Versioning (SemVer)
Format: `vMAJOR.FEATURE.PATCH/BUGFIX`
- `v0.0.0` : First stable production version.
- `v0.1.0` : New feature added (e.g., add GrabFood Get Menu Data).
- `v0.1.1` : Small bug fix (e.g., fix GrabFood menu cron issue).
