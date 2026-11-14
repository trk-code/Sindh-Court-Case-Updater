# Sindh Courts Case Diary Updater

A desktop automation tool built with **Python**, **Tkinter**, **Playwright**, and **OpenPyXL** that updates an advocate's Excel case diary directly from the official **High Court of Sindh (SHC)** and **Sindh District Courts** web portals.

---

## Features

- **Multi-Portal Lookup**
  - **High Court of Sindh**: Queries the SHC Hyderabad Circuit Bench (`cases.shc.gov.pk/hyd`) and Karachi Principal Seat (`cases.shc.gov.pk/cfms-hc-search`), with automatic Karachi fallback when a Hyderabad lookup returns no match.
  - **Sindh District Courts**: Queries `cases.districtcourtssindh.gos.pk` (currently configured for Hyderabad District and the Anti-Encroachment Tribunal) and opens the `#cprofile` diary modal to extract proceedings history and under-section details.
- **Non-Destructive Column Enrichment**
  - Columns **D–H** (`Filed by`, `Respondnets`, `Name of Court`, `Nature of Case`, `Present Status of the Case`) are populated **only if the cell is completely blank**—existing advocate notes are never overwritten.
- **Smart Hearing Date Progression**
  - Updates `Next Hearing` (**Column J**) whenever the portal reports a newer hearing date, and records the previous hearing date in `Last Hearing` (**Column I**).
- **Color-Coded Status Highlighting**
  - **Green**: Updated hearing date or newly populated `Last Hearing`.
  - **Red**: Urgent hearing scheduled within the next **0–3 days**.
  - **Yellow**: Overdue/past hearing date, Crime/FIR manual-check rows, unparsed case numbers, or unmatched cases.
- **Automatic `Upcoming Hearings` Sheet**
  - Generates or refreshes a print-ready, landscape-formatted **`Upcoming Hearings`** tab sorted chronologically for all cases due within the next 3 days.
- **Fuzzy Matching & Typo Tolerance**
  - Parses common diary formatting quirks (e.g., `0f`, `Np.`, missing spaces, inline parenthetical party titles) and ranks District Court search results by party names, case category keywords, and court abbreviations (`ADJ`, `SCJ`, `JAM`).
- **Safety-First Workflow**
  - Detects if the workbook is currently open/locked in Microsoft Excel before starting.
  - Automatically saves a timestamped copy inside a `Backups/` folder next to your Excel file before modifying anything.
- **Visible Browser Mode**
  - Toggle **"Show Browser Window (Visible Mode)"** to observe queries live or handle any manual browser check.
- **Local Chamber Personalization**
  - Optional `"advocate_name"` and `"advocate_subtitle"` keys in a local `settings.json` customize the window title and header without modifying source code.

---

## How Advocates Can Use the App

### 1. Prepare Your Excel Case Diary
Your worksheet should follow the 10-column layout shown in [`sample_cases.xlsx`](sample_cases.xlsx):

| Col | Header | Purpose |
| --- | --- | --- |
| A | `Sr. #` | Serial number |
| B | `Name of Council` | Client / council / department name |
| C | `Constitution Petition No.` | Case / Petition / Suit / Appeal number & year |
| D | `Filed by` | Petitioner / Plaintiff / Appellant *(auto-filled only if blank)* |
| E | `Respondnets` | Respondent / Defendant *(auto-filled only if blank)* |
| F | `Name of Court` | Court name *(helps route SHC vs District & auto-fills if blank)* |
| G | `Nature of Case` | Case category / under-section *(auto-filled only if blank)* |
| H | `Present Status of the Case` | Pending stage or disposal status *(auto-filled only if blank)* |
| I | `Last Hearing` | Previous hearing date *(auto-populated)* |
| J | `Next Hearing` | Upcoming hearing date *(updated & color-coded)* |

> The parser auto-detects the header row within the first 15 rows and skips an optional `1..10` column-number row beneath it.

### 2. Run an Update
1. **Close Microsoft Excel** if the target workbook is currently open.
2. Launch **`SindhCourtUpdater.exe`** (or run `python app.py` from source).
3. Click **Browse Excel File...** and select your `.xlsx` file (try [`sample_cases.xlsx`](sample_cases.xlsx) first).
4. Choose the target **Worksheet Tab** from the dropdown.
5. *(Optional)* Enable **Show Browser Window** if you want to watch the browser live.
6. Click **▶ Start Updating Cases** and monitor the live progress log.
7. Open the updated Excel file afterward and check the newly generated **`Upcoming Hearings`** tab.

### 3. Optional Local Personalization (`settings.json`)
Create a `settings.json` file in the same folder as `app.py` (or next to the `.exe`) to display your name in the header:

```json
{
  "advocate_name": "Advocate Your Name",
  "advocate_subtitle": "Your Designation / Chamber"
}
```

`settings.json` is automatically excluded by `.gitignore` so personal details stay local.

---

## How Developers Can Run from Source

### Prerequisites
- **Python 3.10+** (with Tkinter included)
- **Google Chrome** or **Microsoft Edge** installed on Windows (the scraper uses `channel="chrome"` with an automatic `channel="msedge"` fallback)

### Installation & Launch
```bash
git clone https://github.com/<your-username>/Sindh-Court-Case-Updater.git
cd Sindh-Court-Case-Updater

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
python app.py
```

*(If Playwright asks for browser binaries in an environment without Chrome/Edge, run `playwright install chromium`.)*

### Building a Standalone Windows Executable
```bash
pyinstaller --onefile --windowed --name SindhCourtUpdater --icon app_icon.ico app.py
```

The compiled executable will be generated in `dist/SindhCourtUpdater.exe`.

---

## Project Structure

```text
├── app.py               # Tkinter GUI, Excel processor, and Playwright portal scrapers
├── app_icon.ico         # Multi-resolution application icon
├── sample_cases.xlsx    # Fictional 3-row diary template for safe testing
├── requirements.txt     # Python dependencies (openpyxl, playwright, pillow, pyinstaller)
├── CONTRIBUTING.md      # Guide for extending districts, SHC benches, and categories
├── LICENSE              # MIT License
└── .gitignore           # Excludes private .xlsx workbooks, Backups/, settings.json, build/, dist/
```

---

## Roadmap for Contributors

- [ ] **All Sindh Districts Support**: Map `<select id="district">` IDs for Karachi (East, West, South, Central, Malir, Korangi), Sukkur, Larkana, Jamshoro, Dadu, Mirpurkhas, Shaheed Benazirabad, and others via Column F detection or a UI selector.
- [ ] **SHC Sukkur & Larkana Benches**: Add dedicated search URLs and routing for SHC Sukkur and Larkana circuit benches.
- [ ] **Expanded Case Category & Tribunal Mappings**: Support additional civil, criminal, banking, labour, and special tribunal categories.
- [ ] **Unit Tests**: Add pytest coverage for `parse_case_string()`, `parse_date_flexible()`, `split_case_title_parties()`, and `score_district_row()`.
- [ ] **Exportable Cause List**: One-click PDF or WhatsApp-formatted summary export of the `Upcoming Hearings` sheet.

See [CONTRIBUTING.md](CONTRIBUTING.md) for implementation details.

---

## Disclaimer & License

This utility automates lookup of publicly available case information from official Sindh judicial portals to assist legal practitioners in maintaining personal case diaries. Always verify critical dates against official court cause lists and orders.

Licensed under the [MIT License](LICENSE).
