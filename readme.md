# Work Diary Assistant

> AI-powered desktop tool that automatically analyzes on-screen activity, tracks focus time,
> and generates daily / weekly / monthly work reports — so you can stop writing overtime reports by hand.

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![GUI](https://img.shields.io/badge/GUI-PyQt5%20%7C%20Fluent%20Design-41CD52)
![AI](https://img.shields.io/badge/AI-GLM%20%7C%20Ollama-4FC3F7)
![Storage](https://img.shields.io/badge/Storage-CSV-FFB300)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6)
![License](https://img.shields.io/badge/License-Proprietary-red)

---

## Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Architecture](#architecture)
- [Getting Started](#getting-started)
- [User Guide](#user-guide)
- [Data Model](#data-model)
- [Configuration](#configuration)
- [Project Structure](#project-structure)
- [Development](#development)
- [FAQ](#faq)
- [Roadmap](#roadmap)
- [Contributing](#contributing)
- [License](#license)

---

## Overview

**Work Diary Assistant** is a desktop productivity application for Windows built with
PyQt5 / PyQt-Fluent-Widgets. It periodically captures the screen, uses a vision LLM
(GLM-4.6V-Flash or a local Ollama MiniCPM-V model) to recognize what you are working on,
and accumulates the results into a local, privacy-first dataset. From that data it derives:

- **Real-time statistics** — 24-hour heatmaps, focus duration, active days and work-type breakdown.
- **Automated reports** — one-click daily / weekly / monthly reports generated from customizable templates.
- **Scheduled monitoring** — unattended, interval-based analysis with a background worker queue.

Everything is stored locally as UTF-8 CSV files that open directly in Excel, and the app
optionally syncs records to a companion server for cross-device statistics.

### Problems it solves

| Pain point | Solution |
|---|---|
| Forgetting what you worked on each day | Automatic screen recognition with descriptions |
| Spending hours writing daily reports | Template-driven AI report generation |
| No insight into how time is actually spent | Per-type / per-hour statistics and heatmaps |
| Privacy concerns with cloud-only trackers | Local-first storage, auto-desensitization |

---

## Key Features

- **Smart Recognition** — AI vision models (GLM / Ollama) classify each screenshot into one of 14 work types with a natural-language description.
- **Privacy First** — sensitive content (contacts, accounts, passwords) is auto-desensitized; data never leaves your machine unless you choose to sync.
- **Local Storage** — all records are kept in UTF-8 (BOM) CSV files under `data/`, fully readable in Excel without any export step.
- **Real-time Heatmap** — period (date × 24 h) and annual (week × 7) heatmap views rendered natively with QPainter.
- **Table Export** — the heatmap page exports the underlying dataset as a dependency-free `.xlsx` workbook: summary statistics, hourly distribution matrix, work-type / monthly / weekday statistics and the full detailed records.
- **Scheduled Monitoring** — configurable interval (minutes) with a multi-threaded worker; each tick records exactly one interval of focus time.
- **AI Report Generation** — daily / weekly / monthly reports generated from user-defined templates with streaming output and automatic retry on rate limits.
- **Template Management** — create, edit, import and export report templates (local CSV).
- **Account & Sync (optional)** — email login, record / summary / report synchronization to a companion API server.
- **Modern UI** — Fluent Design style, responsive layout, DPI scaling, system tray, and auto-update checks.

### Work Categories

14 automatic classifications:

`Development` `Communication` `Life` `Learning` `Design` `Management` `Documentation` `Entertainment` `Product` `Meeting` `Operations` `Testing` `Data Analysis` `Other`

---

## Architecture

### Module Layout

```
┌──────────────────────────────────────────────────────────────┐
│                       UI  (PyQt5 + Fluent)                  │
│   Today · Timeline · Reports · History · Heatmap · Monitor  │
│   Records · Screenshot · Settings                            │
└───────────────┬──────────────────────────────┬──────────────┘
                │                              │
        ┌───────▼────────┐            ┌────────▼────────┐
        │   screenshot.py │            │  api_sync.py   │
        │ recognition +   │◄──────────►│  optional      │
        │ monitoring core │            │  server sync   │
        └───────┬────────┘            └─────────────────┘
                │
        ┌───────▼────────┐
        │    store.py    │   CSV persistence + report I/O
        └───────┬────────┘
                │
        ┌───────▼────────────────────────────┐
        │  data/  (CSV, reports, screenshots)│
        └────────────────────────────────────┘
```

### Tech Stack

| Layer | Technology | Purpose |
|---|---|---|
| Language | Python 3.10+ | Application runtime |
| GUI | PyQt5 ≥ 5.15, PyQt-Fluent-Widgets ≥ 1.10 | Fluent Design interface |
| Vision AI | GLM-4.6V-Flash (zhipuai) / Ollama (MiniCPM-V) | Screen content recognition |
| Report AI | GLM / Ollama chat with streaming | Report generation |
| Image | OpenCV, numpy, Pillow, mss | Capture & pre-processing |
| Storage | CSV (utf-8-sig) | Records, summaries, templates |
| Security | cryptography | Credential encryption |
| Theme | darkdetect | OS theme detection |
| Packaging | PyInstaller | Standalone EXE |

---

## Getting Started

### Prerequisites

- Windows 10 / 11
- Python 3.10+ (a conda environment is recommended)
- An AI backend: either a **GLM API key** (set in `screenshot.py`) or a reachable **Ollama** server

### Installation

```bash
# 1. Create and activate an environment (recommended)
conda activate daily-AI          # or: conda create -n daily-AI python=3.10

# 2. Install dependencies
pip install -r requirements.txt

# 3. Optional: install the GLM SDK for the cloud recognition backend
pip install zhipuai
```

### Running

```bash
python UI/main_fluent.py
```

Or double-click `run.bat` (uses the pre-configured conda interpreter).

### First-run workflow

1. Log in with your email (optional — required only for server sync).
2. Open **今日工作** and click the capture button to take the first screenshot; the app records the recognized work type and description.
3. (Optional) Go to **管理监控** to start interval-based monitoring.
4. After a day of usage, open **生成报告**, pick a template and a date range, and generate the report.
5. Review **热力图** for period / annual distribution, and export the underlying data as `.xlsx`.

---

## User Guide

| Page | Icon | Description |
|---|---|---|
| 今日工作 | Home | Today's overview, focus duration, recent records |
| 工作时间线 | Pie | Per-type and per-hour statistics |
| 生成报告 | Document | AI report generation with templates (日/周/月报) |
| 历史报告 | History | Browse, view, export and delete generated reports |
| 热力图 | Calendar | Period / annual heatmaps + table export (`.xlsx`) |
| 管理监控 | Play | Scheduled screenshot monitoring configuration |
| 工作记录（内测） | Document | Raw record list (beta) |
| 截图分析（内测） | Camera | Manual screenshot analysis (beta) |
| 设置 | Settings | Account, DPI scale, recognition model, Ollama host |

---

## Data Model

All files live under `data/` and are UTF-8 with BOM so Excel renders Chinese correctly.

### records.csv — detail records

| Column | Type | Description |
|---|---|---|
| ID | int | Auto-incrementing primary key |
| 日期 | date | `YYYY-MM-DD` |
| 时间 | time | `HH:MM:SS` |
| 工作类型 | string | One of the 14 work categories |
| 工作描述 | string | AI-generated description of the activity |
| 持续时长(分钟) | float | Duration attributed to this record |

### daily_summary.csv — per-day summary

| Column group | Description |
|---|---|
| 日期 / 记录条数 / 使用时长(小时) | Core daily metrics |
| 主要工作 | Work type with the longest accumulated duration |
| 最早/最晚使用时间 | Active window of the day |
| `{类型}时长(小时)` × 14 | Duration per work type |
| `{HH}:00记录数` × 24 | Record count per hour (heatmap source) |

### report_templates.csv — report templates

`name`, `intro`, `desc`, `is_cloud`, `prompt` — the `prompt` supports placeholders
(`{record_count}`, `{duration_hours}`, `{main_work}`, `{hour_data}`, `{records}`) that are filled with live statistics before generation.

---

## Configuration

| Item | Location | Notes |
|---|---|---|
| UI scale factor | `data/config.txt` (`scale_factor=…`) | Applied at startup |
| GLM API key / model | `screenshot.py` (`GLM_API_KEY`, `GLM_MODEL`) | Replace with your own key |
| Ollama host & model | 设置 page | Default `http://192.168.31.23:11434`, `minicpm-v4.6` |
| Sync server | `api_sync.py` (`API_BASE_URL`) | Optional; skipped when not logged in |
| Report templates | `data/report_templates.csv` | Manageable from the UI |
| Generated reports | `data/report/*.md` | Timestamped markdown files |

---

## Project Structure

```
├── main.py                  # Legacy entry point (kept for compatibility)
├── screenshot.py            # Recognition core: capture, AI classify, monitor, report prompts
├── store.py                 # CSV persistence, summaries, templates, report file I/O
├── template.py              # Default report template definitions
├── api_sync.py              # Optional server sync (login / records / summaries / reports)
├── crypto_utils.py          # Credential encryption helpers
├── requirements.txt         # Python dependencies
├── run.bat                  # One-click launcher (pre-configured interpreter)
├── report_templates.csv     # Default template dataset
├── UI/
│   ├── main_fluent.py       # Fluent Design main application (all pages)
│   ├── main.py              # Legacy UI
│   ├── main_window.ui       # Qt Designer layout (legacy)
│   ├── styles.qss           # QSS stylesheet
│   └── *.spec               # PyInstaller build specs
└── data/
    ├── config.txt           # Runtime configuration (scale factor)
    ├── records.csv          # Detail records
    ├── daily_summary.csv    # Per-day summaries
    ├── report_templates.csv # Report templates
    ├── settings.csv         # UI settings
    ├── login_state.json     # Cached login state
    ├── secret.json          # Encrypted credentials
    ├── report/              # Generated reports (*.md)
    └── photo/               # Screenshots captured in test mode
```

---

## Development

### Adding a new page

1. Define a `QWidget` subclass inside `main_fluent.py` (pages are declared inside `main()` so they share its imports and closures).
2. Register it in `MainWindow` via `addSubInterface(page, FluentIcon.XXX, "标题")`.

### Extending recognition

- Add or change work categories in `store.WORK_TYPES` and the corresponding prompt in `screenshot.py` (keep both in sync).
- Recognition failures caused by rate limits are retried automatically (`retry_on_overload`, configurable `MAX_RETRIES` / `RETRY_DELAY`).

### Validation

```bash
python -m py_compile UI/main_fluent.py        # syntax check
python UI/main_fluent.py                      # run the app
```

### Packaging

PyInstaller specs are included under `UI/` (`WorkDiary.spec`, `工作日报助手.spec`).

---

## FAQ

**Q: Do my screenshots get uploaded anywhere?**
A: No. Analysis runs against your chosen backend (GLM API or local Ollama); records and summaries stay in `data/` locally. Server sync happens only when you are logged in and is entirely optional.

**Q: Excel shows garbled characters?**
A: All CSVs are written with UTF-8 BOM (`utf-8-sig`), which Excel detects automatically. If it still misdetects, import the file manually with UTF-8 encoding.

**Q: The heatmap page has a single "生成热力图" button — what does it export?**
A: It exports the heatmap's underlying data as a multi-sheet `.xlsx` workbook (summary, distribution matrix, type / monthly / weekday statistics, and detailed records). No third-party library is required to generate the file.

**Q: How do I switch between GLM and Ollama?**
A: Open 设置 → 识别模型, choose the backend, configure the Ollama host/model if needed, and click 应用设置.

**Q: The report generation fails with a rate-limit message.**
A: The client retries automatically up to `MAX_RETRIES` times. If it still fails, wait a moment or switch to the Ollama backend.

---

## Roadmap

- [ ] Week / month heatmap aggregation views
- [ ] Export reports to DOCX / PDF
- [ ] Multi-language UI (i18n)
- [ ] macOS / Linux support

---

## Contributing

1. Fork the repository and create a feature branch.
2. Keep changes focused; run `py_compile` before submitting.
3. Update this README and `README_ZN.md` when public behavior changes.

---

## License

Proprietary. All rights reserved. This software is provided for internal use;
redistribution or commercial use requires prior written permission.
