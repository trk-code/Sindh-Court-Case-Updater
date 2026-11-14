# Contributing to Sindh Courts Case Diary Updater

Thank you for helping improve **Sindh Courts Case Diary Updater** for advocates and legal teams across Sindh!

Before submitting any pull request, please make sure you **never commit real client workbooks, `Backups/`, or your local `settings.json`**. Use [`sample_cases.xlsx`](sample_cases.xlsx) for local testing.

---

## Development Setup

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

When testing against [`sample_cases.xlsx`](sample_cases.xlsx), enable **"Show Browser Window (Visible Mode)"** in the UI to watch Playwright fill forms and inspect portal HTML selectors in real time.

---

## 1. Adding Support for Other Sindh Districts

Currently, `query_district_portal()` in [`app.py`](app.py) hardcodes Hyderabad District:

```python
# District = Hyderabad ("7")
page.select_option("select#district", "7")
```

To add support for additional districts (e.g., Karachi East, Karachi South, Jamshoro, Dadu, Sukkur, Larkana, Mirpurkhas, Shaheed Benazirabad):

1. Open `https://cases.districtcourtssindh.gos.pk/` (`DIST_URL`) and inspect the `<select id="district">` `<option value="...">` IDs for each district.
2. Add a district lookup dictionary in `app.py` (for example, `DISTRICT_CODES = {"hyderabad": "7", "jamshoro": "...", ...}`).
3. Infer the target district from Column F (`Name of Court`)—falling back to an optional `"default_district"` key in `settings.json` or a UI dropdown—and pass that value into `page.select_option("select#district", district_code)`.
4. Check `score_district_row()` to add any local court abbreviations used in that district (currently it expands `adj`, `jam`, and `scj`).

---

## 2. Adding SHC Sukkur & Larkana Circuit Benches

`extract_shc_detail_view()` in `app.py` already normalizes bench names for `"hyderabad"`, `"karachi"`, `"sukkur"`, and `"larkana"`, while `parse_case_string()` and `process_worksheet()` currently route between:

- `SHC_HYD_URL`: `https://cases.shc.gov.pk/hyd/web/index.php?r=cases%2Fsearch`
- `SHC_KHI_URL`: `https://cases.shc.gov.pk/cfms-hc-search/web/index.php?r=cases%2Fsearch`

To add **SHC Circuit Court Sukkur** and **SHC Circuit Court Larkana**:

1. Define the portal URL constants (`SHC_SUK_URL`, `SHC_LRK_URL`) alongside `SHC_HYD_URL` and `SHC_KHI_URL` at the top of `app.py`.
2. Update the portal selection block inside `parse_case_string()` when `is_shc` is `True`:
   - Check `"sukkur" in court_low` -> `"shc_suk"`
   - Check `"larkana" in court_low` -> `"shc_lrk"`
   - Check `"karachi" in court_low` -> `"shc_khi"`
   - Default to `"shc_hyd"` (or add a configurable default bench).
3. Add matching `elif portal == ...` branches inside `process_worksheet()` calling `query_shc_portal()`.
4. Verify that the CFMS form selectors (`#casessearch-casenamecode`, `#casessearch-caseno`, `#casessearch-caseyear`, `#casessearch-bench`, `#submit_search`, `#crud-datatable`, and `table#w1`) match those bench portals.

---

## 3. Adding New Case Category Mappings

### High Court of Sindh (`shc_category_code`)
In `parse_case_string()`, SHC case prefixes are mapped to the portal's `#casessearch-casenamecode` dropdown values (defaulting to `"21"` for Constitutional Petitions):

- `"7"` — Civil Revision
- `"40"` — Second / IInd / 2nd Appeal
- `"17"` — First / Ist / 1st Appeal
- `"16"` — Misc Appeal
- `"14"` — H.C.A. / High Court Appeal
- `"10"` — F.R.A.
- `"11"` — Criminal Appeal
- `"8"` — Criminal Bail
- `"25"` — Cr. Misc
- `"24"` — Cr. Rev
- `"45"` — Suit

To add a new SHC category:
1. Inspect the `<select id="casessearch-casenamecode">` options on the SHC search page to obtain its numeric value.
2. Add an `elif` keyword check in `parse_case_string()` (and update `is_shc` regex if the prefix does not explicitly say `"High Court"` in Column F).

### Sindh District Courts (`dist_cat_keywords`)
District portals return multiple case categories for the same case number and year. `parse_case_string()` populates `dist_cat_keywords` so `score_district_row()` gives `+2.0` score weight to matching categories.

To add a new District category (for example, `Succession`, `Guardian & Wards`, `Banking`, `Labour`, or `Anti-Terrorism`):
1. Add a clean `elif` pattern in `parse_case_string()` returning lowercase keywords matching the District portal's result title prefix.
2. If the category requires selecting a specific `<select id="courttype">` filter before submitting (similar to `"encroachment"` selecting `"22"` for the Removal of Anti-Encroachment Tribunal), add that handling in `query_district_portal()`.

---

## Pull Request Checklist

- [ ] Code parses [`sample_cases.xlsx`](sample_cases.xlsx) cleanly (`python -m py_compile app.py` passes with zero errors).
- [ ] No hardcoded personal names, local file paths, or real client case details in code, comments, or docstrings.
- [ ] `git status --short` contains no `.xlsx` files other than `sample_cases.xlsx`, and no `settings.json`, `Backups/`, `build/`, or `dist/`.
