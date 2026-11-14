import datetime
import json
import os
import re
import shutil
import sys
import threading
import tkinter as tk
from copy import copy
from difflib import SequenceMatcher
from tkinter import filedialog, messagebox, scrolledtext, ttk

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

# Highlight fills
FILL_UPDATED = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Light green
FILL_URGENT = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")   # Light red (within 3 days)
FILL_WARN = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")     # Light yellow (overdue/unfound)

SHC_HYD_URL = "https://cases.shc.gov.pk/hyd/web/index.php?r=cases%2Fsearch"
SHC_KHI_URL = "https://cases.shc.gov.pk/cfms-hc-search/web/index.php?r=cases%2Fsearch"
DIST_URL = "https://cases.districtcourtssindh.gos.pk/"

GENERIC_NON_IDENTIFYING = {
    "the", "and", "ors", "others", "other", "anr", "another", "through", "thr",
    "lrs", "legal", "heirs", "mst", "syed", "mir", "sardar", "state", "province",
    "sindh", "pakistan", "federation", "govt", "government", "secretary", "chief",
    "district", "deputy", "commissioner", "board", "revenue", "sho", "ps", "po",
    "tmc", "shah", "latifabad", "hyderabad", "karachi", "private", "ltd", "pvt",
    "m/s", "ms", "re", "vs", "versus",
}

COMMON_NAME_TOKENS = {
    "muhammad", "mohammad", "mohd", "ahmed", "ahmad", "ali", "khan", "begum",
}


def get_app_dir() -> str:
    """Return the directory where the script or frozen .exe resides."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


SETTINGS_PATH = os.path.join(get_app_dir(), "settings.json")


def load_settings() -> dict:
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_settings(data: dict) -> None:
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


def is_excel_locked(filepath: str) -> bool:
    """Check if the Excel file is currently open/locked in MS Excel."""
    folder = os.path.dirname(os.path.abspath(filepath))
    basename = os.path.basename(filepath)
    lock_file = os.path.join(folder, f"~${basename}")
    if os.path.exists(lock_file):
        return True
    try:
        with open(filepath, "r+b"):
            pass
        return False
    except PermissionError:
        return True
    except OSError:
        return True


def create_backup(filepath: str) -> str:
    """Create a timestamped backup inside a Backups/ folder next to the Excel file."""
    folder = os.path.dirname(os.path.abspath(filepath))
    backups_dir = os.path.join(folder, "Backups")
    os.makedirs(backups_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(filepath))
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(backups_dir, f"{stem}_backup_{ts}{ext}")
    shutil.copy2(filepath, backup_path)
    return backup_path


def is_cell_empty(cell) -> bool:
    """Return True if an Excel cell's value is None or whitespace-only."""
    return cell.value is None or str(cell.value).strip() == ""


def apply_neighbour_style(ws, target_row: int, col_idx: int, first_data_row: int) -> None:
    """Copy font, border, and alignment from the nearest bordered data row above target_row."""
    ref_row = None
    for rr in range(target_row - 1, first_data_row - 1, -1):
        b = ws.cell(rr, col_idx).border
        if b and ((b.top and b.top.style) or (b.bottom and b.bottom.style) or (b.left and b.left.style)):
            ref_row = rr
            break
    if ref_row is None and target_row - 1 >= first_data_row:
        ref_row = target_row - 1
    if ref_row is None:
        return

    src = ws.cell(ref_row, col_idx)
    dst = ws.cell(target_row, col_idx)
    if src.has_style:
        if src.font:
            dst.font = copy(src.font)
        if src.border:
            dst.border = copy(src.border)
        if src.alignment:
            dst.alignment = copy(src.alignment)


def parse_date_flexible(text) -> datetime.date | None:
    """Parse dates from portal or Excel strings into a datetime.date."""
    if text is None:
        return None
    if isinstance(text, datetime.datetime):
        return text.date()
    if isinstance(text, datetime.date):
        return text

    s = str(text).strip()
    if not s:
        return None

    # Clean common suffixes like "(Auto Fixation)" or weekday names
    cleaned = re.sub(r"\(.*?\)", "", s).strip()
    cleaned = re.sub(
        r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s*",
        "",
        cleaned,
        flags=re.IGNORECASE,
    ).strip()

    m_num = re.search(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{4})\b", cleaned)
    if m_num:
        d, m, y = int(m_num.group(1)), int(m_num.group(2)), int(m_num.group(3))
        try:
            return datetime.date(y, m, d)
        except ValueError:
            pass

    m_mon = re.search(r"\b(\d{1,2})[-/\s]([A-Za-z]{3,9})[-/\s](\d{2,4})\b", cleaned)
    if m_mon:
        day_s, mon_s, yr_s = m_mon.group(1), m_mon.group(2), m_mon.group(3)
        if len(yr_s) == 2:
            yr_s = "20" + yr_s
        for fmt in ("%d %b %Y", "%d %B %Y"):
            try:
                return datetime.datetime.strptime(f"{day_s} {mon_s} {yr_s}", fmt).date()
            except ValueError:
                continue

    for fmt in ("%B %d, %Y", "%b %d, %Y", "%B %d %Y", "%b %d %Y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue

    return None


def split_case_title_parties(title_text: str) -> tuple[str, str]:
    """
    Split a case title into (petitioner, respondent) on 'vs', 'v/s', 'v/z', or 'versus'.
    Handles prefixes like 'PPC 514/2021, ...' and 'Re-...'.
    """
    if not title_text:
        return "", ""

    s = re.sub(r"\s+", " ", str(title_text)).strip()
    if s.startswith("(") and s.endswith(")"):
        s = s[1:-1].strip()

    # Strip leading '<Category> <No>/<Year>, ' prefix if present
    s = re.sub(r"^[^,]{2,60}?\b\d+[A-Za-z]?/\d{4}\s*,\s*", "", s).strip()

    parts = re.split(r"\s+(?:v/?s\.?|v/?z\.?|versus)\s+", s, maxsplit=1, flags=re.IGNORECASE)
    if len(parts) != 2:
        return "", ""

    pet = re.sub(r"^Re[-:\s]+", "", parts[0].strip(), flags=re.IGNORECASE).strip(" ,")
    resp = parts[1].strip(" ,")
    # Remove trailing lower-court cross-reference like ' - F.C Suit No.101/2024' if needed? Keep clean
    return pet, resp


def extract_inline_parties(raw_text: str) -> tuple[str, str]:
    """
    Extract party names from inline parentheticals in Column C such as:
    'Direct Complaint No.12 of 2025 (Re-Tariq Mahmood v/z Kamran Siddiqui & Others)'.
    """
    if not raw_text:
        return "", ""
    for chunk in re.findall(r"\(([^)]+)\)", str(raw_text)):
        if re.search(r"\b(?:v/?s\.?|v/?z\.?|versus)\b", chunk, re.IGNORECASE):
            pet, resp = split_case_title_parties(chunk)
            if pet or resp:
                return pet, resp
    return "", ""


def parse_case_string(raw_text: str, court_name: str = "") -> dict | None:
    """
    Parse Column C case numbers including typos and inline parenthetical party names:
      - 'CP No.D-101 0f 2024'
      - 'Execution Application14 of 2019'
      - 'F.C.S No.205 2024'
      - 'C.P No.D/ 312 of 2025'
      - 'Civil Appeal Np. 48 of 2026'
      - 'Civil Misc. Application No118 of2023'
      - 'Cr Case no 77 of 2022'
      - 'Direct Complaint No.12 of 2025 (Re-Tariq Mahmood v/z Kamran Siddiqui & Others)'
    """
    if not raw_text:
        return None

    inline_pet, inline_resp = extract_inline_parties(str(raw_text))

    lines = [ln.strip() for ln in str(raw_text).splitlines() if ln.strip()]
    if not lines:
        return None
    text = lines[-1]

    # Remove parenthetical notes (such as inline parties) before extracting case number/year
    text_no_parens = re.sub(r"\([^)]*\)", " ", text).strip() or text

    # Fix common typos
    norm = re.sub(r"\b0f\b", "of", text_no_parens, flags=re.IGNORECASE)
    norm = re.sub(r"\bNp\.?", "No.", norm, flags=re.IGNORECASE)
    norm = re.sub(r"\bof(\d{4})\b", r"of \1", norm, flags=re.IGNORECASE)
    norm = re.sub(r"\bNo\.?(\d)", r"No. \1", norm, flags=re.IGNORECASE)
    norm = re.sub(r"([A-Za-z])(\d)", r"\1 \2", norm)
    norm = re.sub(r"\s+", " ", norm).strip()

    low = norm.lower()
    court_low = (court_name or "").lower()

    is_crime = bool(re.search(r"\bcrime\s+no\b|\bfir\b", low))

    bench = ""
    m_bench = re.search(r"\b([DS])\s*[-/]\s*\d+", norm, flags=re.IGNORECASE)
    if m_bench:
        bench = m_bench.group(1).upper()

    m_num_yr = re.search(
        r"(?:[DS]\s*[-/]\s*)?(\d+)\s*(?:of\s+|/\s*|\s+)((?:19|20)\d{2})\b",
        norm,
        flags=re.IGNORECASE,
    )
    if not m_num_yr:
        return None

    case_no = str(int(m_num_yr.group(1)))
    case_year = m_num_yr.group(2)

    is_shc = (
        "high court" in court_low
        or bool(
            re.search(
                r"\b(c\.?\s*p\.?|const\.?\s*p|civil\s+revision|second\s+appeal|iind\s+appeal|ist\s+appeal|h\.?\s*c\.?\s*a)\b",
                low,
            )
        )
    )

    if is_shc:
        portal = "shc_khi" if "karachi" in court_low else "shc_hyd"
        shc_code = "21"
        if "civil revision" in low:
            shc_code = "7"
        elif "second appeal" in low or "iind appeal" in low or "2nd appeal" in low:
            shc_code = "40"
        elif "ist appeal" in low or "1st appeal" in low or "first appeal" in low:
            shc_code = "17"
        elif "misc appeal" in low:
            shc_code = "16"
        elif "h.c.a" in low or "high court appeal" in low:
            shc_code = "14"
        elif "f.r.a" in low:
            shc_code = "10"
        elif "cr.appeal" in low or "criminal appeal" in low:
            shc_code = "11"
        elif "cr.bail" in low or "bail" in low:
            shc_code = "8"
        elif "cr.misc" in low:
            shc_code = "25"
        elif "cr.rev" in low:
            shc_code = "24"
        elif "suit" in low:
            shc_code = "45"

        return {
            "portal": portal,
            "shc_category_code": shc_code,
            "shc_bench": bench,
            "dist_cat_keywords": [],
            "case_no": case_no,
            "case_year": case_year,
            "is_crime": is_crime,
            "normalized": norm,
            "inline_petitioner": inline_pet,
            "inline_respondent": inline_resp,
        }

    dist_keywords = []
    if re.search(r"\bf\.?\s*c\.?\s*s\b|\bf\.?\s*c\.?\s*suit\b|\bf\.?\s*c\s+no\b", low):
        dist_keywords = ["ist class civil", "first class", "civil suit", "f.c"]
    elif "direct complaint" in low:
        dist_keywords = ["direct complaint"]
    elif "private complaint" in low:
        dist_keywords = ["private complaint", "direct complaint"]
    elif "family execution" in low:
        dist_keywords = ["family execution"]
    elif "family" in low:
        dist_keywords = ["family"]
    elif "execution" in low and "rent" in low:
        dist_keywords = ["rent execution", "execution", "rent"]
    elif "execution" in low:
        dist_keywords = ["execution"]
    elif "civil appeal" in low or re.search(r"\bc\.?\s*a\b", low):
        dist_keywords = ["civil appeal"]
    elif "civil misc" in low:
        dist_keywords = ["civil misc"]
    elif "encroachment" in low:
        dist_keywords = ["encroachment"]
    elif "criminal misc" in low or "cr. misc" in low:
        dist_keywords = ["criminal misc", "cr. misc", "crl. misc"]
    elif "rent" in low:
        dist_keywords = ["rent"]
    elif re.search(r"\bcr\.?\s*case\b|\bsessions\s+case\b|\bppc\b", low):
        dist_keywords = ["ppc", "murder", "sessions", "sindh arms", "narcotics", "other cr"]

    return {
        "portal": "district",
        "shc_category_code": "",
        "shc_bench": "",
        "dist_cat_keywords": dist_keywords,
        "case_no": case_no,
        "case_year": case_year,
        "is_crime": is_crime,
        "normalized": norm,
        "inline_petitioner": inline_pet,
        "inline_respondent": inline_resp,
    }


def _norm_token(tok: str) -> str:
    t = tok.lower().strip()
    repl = {"osama": "usama", "shaikh": "sheikh", "mohammad": "muhammad", "mohammed": "muhammad"}
    return repl.get(t, t)


def _extract_party_tokens(text: str) -> list[str]:
    if not text:
        return []
    raw_tokens = re.findall(r"[a-zA-Z]{3,}", text.lower())
    distinct = [_norm_token(t) for t in raw_tokens if t not in GENERIC_NON_IDENTIFYING and t not in COMMON_NAME_TOKENS]
    if distinct:
        return distinct
    # Only fall back to common personal names (Muhammad/Ali/Khan/Ahmed), NEVER generic boilerplate like 'others'/'state'
    return [_norm_token(t) for t in raw_tokens if t in COMMON_NAME_TOKENS]


def score_district_row(
    row_title: str,
    row_court: str,
    cat_keywords: list[str],
    petitioner: str,
    respondent: str,
    excel_court: str,
) -> tuple[float, int]:
    """
    Score a District Courts search result row against party names, category, and court.
    Returns (total_score, party_token_hits).
    """
    row_title_low = (row_title or "").lower()
    row_words = {_norm_token(w) for w in re.findall(r"[a-zA-Z]{3,}", row_title_low)}
    combined_low = f"{row_title} {row_court}".lower()

    party_tokens = _extract_party_tokens(petitioner) + _extract_party_tokens(respondent)
    # Deduplicate preserving order
    seen = set()
    uniq_tokens = []
    for t in party_tokens:
        if t not in seen:
            seen.add(t)
            uniq_tokens.append(t)

    party_hits = 0
    for tok in uniq_tokens:
        if tok in row_words or re.search(rf"\b{re.escape(tok)}\b", row_title_low):
            party_hits += 1
        elif any(
            rw[0] == tok[0] and SequenceMatcher(None, tok, rw).ratio() >= 0.86
            for rw in row_words
        ):
            party_hits += 1

    score = 0.0
    if uniq_tokens:
        score += (party_hits / len(uniq_tokens)) * 3.0
    for party in (petitioner, respondent):
        if party:
            score += SequenceMatcher(None, str(party).lower()[:40], row_title_low[:80]).ratio() * 0.5

    if cat_keywords and any(kw in combined_low for kw in cat_keywords):
        score += 2.0

    # Match expanded court terms (e.g. 'II ADJ' -> 'additional district' + 'ii')
    if excel_court:
        ec = excel_court.lower()
        ec_exp = (
            ec.replace("adj", "additional district")
            .replace("jam", "judicial magistrate")
            .replace("scj", "senior civil judge")
        )
        court_words = [w for w in re.findall(r"[a-z0-9]+", ec_exp) if len(w) >= 2]
        rc_low = (row_court or "").lower()
        c_hits = sum(1 for w in court_words if re.search(rf"\b{re.escape(w)}\b", rc_low))
        if court_words:
            score += (c_hits / len(court_words)) * 1.0

    return score, party_hits


def _empty_portal_result(status: str = "") -> dict:
    return {
        "found": False,
        "next_date": None,
        "last_date": None,
        "status": status,
        "petitioner": "",
        "respondent": "",
        "court_name": "",
        "nature_of_case": "",
        "present_status": "",
    }


def query_shc_portal(page, url: str, parsed: dict, petitioner: str = "", respondent: str = "") -> dict:
    """Query SHC Hyderabad or SHC Karachi portal and return structured case metadata & dates."""
    page.goto(url, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_selector("#casessearch-caseno", timeout=15000)

    cat_code = parsed["shc_category_code"]
    if cat_code:
        page.evaluate(
            """(val) => {
                const sel = document.querySelector('#casessearch-casenamecode');
                if (sel) {
                    sel.value = val;
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                    if (window.jQuery) { window.jQuery(sel).trigger('change'); }
                }
            }""",
            cat_code,
        )

    page.fill("#casessearch-caseno", parsed["case_no"])
    page.fill("#casessearch-caseyear", parsed["case_year"])

    if parsed["shc_bench"]:
        try:
            page.select_option("#casessearch-bench", parsed["shc_bench"])
        except Exception:
            pass

    with page.expect_navigation(wait_until="domcontentloaded", timeout=30000):
        page.click("#submit_search")

    current_url = page.url

    if "cases%2Fview" in current_url or "cases/view" in current_url:
        return extract_shc_detail_view(page, parsed)

    rows = page.locator("#crud-datatable table tbody tr")
    row_count = rows.count()
    if row_count > 0:
        eff_pet = petitioner or parsed.get("inline_petitioner", "")
        eff_resp = respondent or parsed.get("inline_respondent", "")

        best_idx = 0
        best_score = -1.0
        for idx in range(row_count):
            cells = rows.nth(idx).locator("td")
            if cells.count() < 12:
                continue
            title_txt = cells.nth(6).inner_text().strip()
            bench_txt = cells.nth(4).inner_text().strip()
            sc, _ = score_district_row(title_txt, "", [], eff_pet, eff_resp, "")
            if parsed["shc_bench"] and bench_txt.upper() == parsed["shc_bench"]:
                sc += 1.0
            if sc > best_score:
                best_score = sc
                best_idx = idx

        best_row = rows.nth(best_idx)
        cells = best_row.locator("td")
        grid_last = parse_date_flexible(cells.nth(8).inner_text().strip()) if cells.count() >= 12 else None
        grid_next = parse_date_flexible(cells.nth(9).inner_text().strip()) if cells.count() >= 12 else None
        grid_disp = parse_date_flexible(cells.nth(10).inner_text().strip()) if cells.count() >= 12 else None
        grid_status = cells.nth(11).inner_text().strip() if cells.count() >= 12 else ""
        grid_title = cells.nth(6).inner_text().strip() if cells.count() >= 12 else ""
        grid_matter = cells.nth(7).inner_text().strip() if cells.count() >= 12 else ""
        grid_circuit = cells.nth(5).inner_text().strip() if cells.count() >= 12 else ""

        view_link = best_row.locator("a.btn-primary, a[href*='cases%2Fview'], a[href*='cases/view']").first
        if view_link.count() > 0:
            try:
                with page.expect_navigation(wait_until="domcontentloaded", timeout=30000):
                    view_link.click()
                res = extract_shc_detail_view(page, parsed)
                if not res["last_date"]:
                    res["last_date"] = grid_last or grid_disp
                if not res["next_date"]:
                    res["next_date"] = grid_next
                if not res["nature_of_case"] and grid_matter and "not entered" not in grid_matter.lower():
                    res["nature_of_case"] = grid_matter
                return res
            except Exception:
                pass

        pet, resp = split_case_title_parties(grid_title)
        court_str = f"High Court of Sindh, {grid_circuit}" if grid_circuit else "High Court of Sindh"
        return {
            "found": True,
            "next_date": grid_next,
            "last_date": grid_last or grid_disp,
            "status": grid_status or "Found",
            "petitioner": pet or eff_pet,
            "respondent": resp or eff_resp,
            "court_name": court_str,
            "nature_of_case": "" if "not entered" in grid_matter.lower() else grid_matter,
            "present_status": grid_status,
        }

    return _empty_portal_result("Not found on SHC portal")


def extract_shc_detail_view(page, parsed: dict | None = None) -> dict:
    """Extract Next Hearing, Last Hearing, and full case metadata from SHC cases/view."""
    today = datetime.date.today()
    body_text = page.locator("body").inner_text()

    status_txt = "Pending"
    m_stat = re.search(
        r"Case Profile\s*\(\d+\)\s*(Pending|Disposed[^\n]*|Dismissed[^\n]*)",
        body_text,
        re.IGNORECASE,
    )
    if m_stat:
        status_txt = m_stat.group(1).strip()

    # Extract Case No. line & Circuit / Matter parenthetical
    case_no_val = ""
    case_title_val = ""
    undersection_val = ""
    last_detail_txt = ""
    next_detail_txt = ""
    last_roster_txt = ""
    future_fix_txt = ""
    last_adlreg_txt = ""
    disposal_txt = ""

    for tr in page.locator("table#w1 tr").all():
        ths = [th.inner_text().strip() for th in tr.locator("th").all()]
        tds = [td.inner_text().strip() for td in tr.locator("td").all()]
        if not ths or not tds or len(ths) != len(tds):
            continue
        for header, val in zip(ths, tds):
            hl = header.lower()
            if hl.startswith("case no"):
                case_no_val = val
            elif hl.startswith("case title"):
                case_title_val = val
            elif "undersection" in hl:
                undersection_val = val
            elif hl.startswith("last hearing detail"):
                last_detail_txt = val
            elif hl.startswith("next hearing detail"):
                next_detail_txt = val
            elif hl.startswith("last roster discharged"):
                last_roster_txt = val
            elif hl.startswith("future fixation"):
                future_fix_txt = val
            elif hl.startswith("last hg"):
                last_adlreg_txt = val
            elif hl.startswith("disposal date"):
                disposal_txt = val

    # Parties
    pet, resp = split_case_title_parties(case_title_val)
    if not pet and not resp:
        m_title = re.search(r"Case Title\s+([^\n]+)", body_text, re.IGNORECASE)
        if m_title:
            pet, resp = split_case_title_parties(m_title.group(1))
    if parsed:
        pet = pet or parsed.get("inline_petitioner", "")
        resp = resp or parsed.get("inline_respondent", "")

    # Court Name
    cn_low = case_no_val.lower()
    if "hyderabad" in cn_low:
        court_name = "High Court of Sindh Circuit Bench Hyderabad"
    elif "karachi" in cn_low:
        court_name = "High Court of Sindh Karachi"
    elif "sukkur" in cn_low:
        court_name = "High Court of Sindh Circuit Court Sukkur"
    elif "larkana" in cn_low:
        court_name = "High Court of Sindh Circuit Court Larkana"
    else:
        court_name = "High Court of Sindh"

    # Nature of Case (parenthetical in Case No. or Undersection)
    nature_of_case = ""
    m_mat = re.search(r"\(([^()]{3,})\)\s*$", case_no_val.replace("\n", " ").strip())
    if m_mat and m_mat.group(1).strip().upper() not in ("S", "D"):
        nature_of_case = m_mat.group(1).strip().title()
    elif undersection_val and "(not set)" not in undersection_val.lower():
        nature_of_case = undersection_val.strip().rstrip(",")

    # Next Hearing Date
    next_date = None
    m_nd = re.search(r"Date:\s*([^\n|]+)", next_detail_txt or "", re.IGNORECASE)
    if m_nd:
        next_date = parse_date_flexible(m_nd.group(1))
    if next_date is None and future_fix_txt:
        next_date = parse_date_flexible(future_fix_txt)
    if next_date is None:
        m_next = re.search(r"Next Hearing Detail\s*Date:\s*([^\n]+)", body_text, re.IGNORECASE)
        if m_next:
            next_date = parse_date_flexible(m_next.group(1))
    if next_date is None:
        m_fut = re.search(r"Future Fixation\s*\(Tentative\)\s*([^\n]+)", body_text, re.IGNORECASE)
        if m_fut:
            next_date = parse_date_flexible(m_fut.group(1))

    # Last Hearing Date: pick newest past date (<= today and != next_date) from detail/roster/adlreg/disposal
    past_candidates = []
    for txt in (last_detail_txt, last_roster_txt, last_adlreg_txt, disposal_txt):
        if not txt or "(not set)" in txt.lower():
            continue
        m_d = re.search(r"Date:\s*([^\n|]+)", txt, re.IGNORECASE)
        dt = parse_date_flexible(m_d.group(1) if m_d else txt)
        if dt and dt <= today and (next_date is None or dt <= next_date):
            past_candidates.append(dt)

    last_date = max(past_candidates) if past_candidates else None

    # Fallback: check Case History tab if last_date is still None
    if last_date is None:
        try:
            hist_tab = page.locator("a[role='tab']:has-text('Case History')").first
            if hist_tab.count() > 0:
                hist_tab.click()
                page.wait_for_selector("#w3-tab2 table.kv-grid-table tbody tr", timeout=6000)
                hist_rows = page.locator("#w3-tab2 table.kv-grid-table tbody tr")
                for i in range(hist_rows.count()):
                    first_cell = hist_rows.nth(i).locator("td").first
                    dt = parse_date_flexible(first_cell.inner_text().strip())
                    if dt and dt <= today:
                        last_date = dt
                        break
        except Exception:
            pass

    # Stage / Present Status
    stage = ""
    m_stage = re.search(r"Stage:\s*([^\n|]+)", last_detail_txt or "", re.IGNORECASE)
    if m_stage:
        stage = m_stage.group(1).strip()
    if stage and stage.lower() not in ("(not set)", "not set", ""):
        present_status = f"{status_txt} ({stage})"
    else:
        present_status = status_txt

    return {
        "found": True,
        "next_date": next_date,
        "last_date": last_date,
        "status": status_txt,
        "petitioner": pet,
        "respondent": resp,
        "court_name": court_name,
        "nature_of_case": nature_of_case,
        "present_status": present_status,
    }


def _extract_district_diary_and_meta(page, next_date: datetime.date | None, today: datetime.date) -> tuple[datetime.date | None, str, str]:
    """
    Extract (last_date, under_section, diary_stage) from District Courts #cprofile modal.
    """
    last_date = None
    under_section = ""
    diary_stage = ""

    tables = page.locator("#cprofile table")
    for t_idx in range(tables.count()):
        tbl = tables.nth(t_idx)
        headers = [th.inner_text().strip().lower() for th in tbl.locator("th").all()]

        # Case Detail table
        if any("under section" in h for h in headers):
            for tr in tbl.locator("tr").all():
                cells = tr.locator("th, td").all()
                for ci in range(len(cells) - 1):
                    if "under section" in cells[ci].inner_text().strip().lower():
                        under_section = cells[ci + 1].inner_text().strip()
                        break

        # Proceedings History / Diary table (columns: S.No, Diary, Date)
        if "diary" in headers and "date" in headers:
            rows = tbl.locator("tbody tr")
            if rows.count() == 0:
                rows = tbl.locator("tr")
            for r_idx in range(rows.count()):
                tds = rows.nth(r_idx).locator("td")
                if tds.count() < 3:
                    continue
                diary_txt = tds.nth(1).inner_text().strip()
                date_txt = tds.nth(2).inner_text().strip()
                dt = parse_date_flexible(date_txt)
                if dt is None:
                    continue
                # Skip the empty future hearing placeholder row (top row == next_date with blank diary)
                if not diary_txt and (dt > today or (next_date is not None and dt == next_date)):
                    continue
                last_date = dt
                if diary_txt:
                    # Extract stage snippet if present (e.g. 'Put off to 26-11-2026. For ...')
                    m_for = re.search(r"(?:put\s+(?:up|off)\s+to[^.]{0,30}\.?\s*)(for\s+[^.\n]{3,80})", diary_txt, re.IGNORECASE)
                    if m_for:
                        stage_txt = m_for.group(1).strip()
                        if stage_txt:
                            diary_stage = stage_txt[0].upper() + stage_txt[1:]
                break

    return last_date, under_section, diary_stage


def query_district_portal(
    page, parsed: dict, petitioner: str, respondent: str, court_name: str
) -> dict:
    """Query Sindh District Courts portal for Hyderabad and return structured metadata & dates."""
    today = datetime.date.today()
    eff_pet = petitioner or parsed.get("inline_petitioner", "")
    eff_resp = respondent or parsed.get("inline_respondent", "")

    rows = None
    count = 0
    timed_out = False
    for attempt in range(3):
        timed_out = False
        try:
            page.goto(DIST_URL, wait_until="domcontentloaded", timeout=45000)
            page.wait_for_selector("select#district", timeout=20000)
            page.wait_for_selector("button[type='submit'].btn-success", timeout=10000)

            # District = Hyderabad ("7")
            page.select_option("select#district", "7")
            page.wait_for_timeout(300)

            # Removal of Anti Encroachment Tribunal option value is "22"
            if "encroachment" in parsed["dist_cat_keywords"]:
                try:
                    page.select_option("select#courttype", "22")
                    page.wait_for_timeout(500)
                except Exception:
                    pass

            page.fill("input#caseno", parsed["case_no"])
            page.fill("input#caseyear", parsed["case_year"])

            all_radio = page.locator("input#inlineCheckbox3[name='status[]']")
            if all_radio.count() > 0:
                all_radio.check()

            with page.expect_navigation(wait_until="domcontentloaded", timeout=45000):
                page.click("button[type='submit'].btn-success")

            rows = page.locator("table.table.table-striped tbody tr")
            count = rows.count()
            if count > 0:
                break
        except PlaywrightTimeoutError:
            timed_out = True
            count = 0

        if attempt < 2:
            page.wait_for_timeout(1200)

    if timed_out and count == 0:
        return _empty_portal_result("Timeout")
    if count == 0:
        return _empty_portal_result("Not found on District portal")

    best_idx = 0
    best_score = -1.0
    best_party_hits = 0
    cat_keywords = parsed["dist_cat_keywords"]

    for idx in range(count):
        cells = rows.nth(idx).locator("td")
        if cells.count() < 5:
            continue
        case_title_txt = cells.nth(1).inner_text().strip()
        court_txt = cells.nth(2).inner_text().strip()
        sc, p_hits = score_district_row(
            case_title_txt, court_txt, cat_keywords, eff_pet, eff_resp, court_name
        )
        if sc > best_score:
            best_score = sc
            best_idx = idx
            best_party_hits = p_hits

    # Safety check: if party names are known, prevent false matches with completely unrelated cases
    has_party_info = bool(_extract_party_tokens(eff_pet) or _extract_party_tokens(eff_resp))
    if has_party_info and best_party_hits == 0:
        return _empty_portal_result("No matching party on District portal")

    best_row = rows.nth(best_idx)
    best_cells = best_row.locator("td")
    if best_cells.count() < 5:
        return _empty_portal_result("Invalid result row")

    col1_title = best_cells.nth(1).inner_text().strip()
    portal_court = best_cells.nth(2).inner_text().strip()
    status_txt = best_cells.nth(3).inner_text().strip()
    hearing_txt = best_cells.nth(4).inner_text().strip()
    next_date = parse_date_flexible(hearing_txt)

    p_pet, p_resp = split_case_title_parties(col1_title)
    p_pet = p_pet or eff_pet
    p_resp = p_resp or eff_resp

    # Category prefix fallback (before '<caseno>/<year>,')
    cat_prefix = ""
    m_cat = re.match(r"^(.+?)\s+\d+[A-Za-z]?/\d{4}\s*,", col1_title)
    if m_cat:
        cat_prefix = m_cat.group(1).strip()

    # Open #cprofile modal to retrieve Last Hearing Date (Diary) & Under Section
    last_date = None
    under_section = ""
    diary_stage = ""
    pview_btn = best_row.locator("button.pview")
    if pview_btn.count() > 0:
        try:
            pview_btn.first.click()
            page.wait_for_selector("#cprofile.show table", timeout=10000)
            page.wait_for_timeout(350)
            last_date, under_section, diary_stage = _extract_district_diary_and_meta(page, next_date, today)
        except Exception:
            pass

    # If modal had no diary date (or failed), fallback to status date only if disposed
    if last_date is None and "disposed" in status_txt.lower():
        m_sd = re.search(r"(\d{1,2}/[A-Za-z]{3}/\d{4})", status_txt)
        if m_sd:
            last_date = parse_date_flexible(m_sd.group(1))

    nature_of_case = under_section if under_section and under_section != "-" else cat_prefix
    present_status = status_txt
    if diary_stage and "pending" in status_txt.lower():
        present_status = f"Pending — {diary_stage}"

    return {
        "found": True,
        "next_date": next_date,
        "last_date": last_date,
        "status": status_txt,
        "petitioner": p_pet,
        "respondent": p_resp,
        "court_name": portal_court,
        "nature_of_case": nature_of_case,
        "present_status": present_status,
    }


def detect_sheet_layout(ws) -> tuple[int, int]:
    """
    Detect header row and first data row dynamically.
    Returns (header_row, first_data_row).
    """
    header_row = 4
    for r in range(1, min(15, ws.max_row + 1)):
        c3 = str(ws.cell(r, 3).value or "").lower()
        c9 = str(ws.cell(r, 9).value or "").lower()
        c10 = str(ws.cell(r, 10).value or "").lower()
        if "petition" in c3 or ("last hearing" in c9 and "next hearing" in c10):
            header_row = r
            break

    first_data_row = header_row + 1
    if ws.cell(first_data_row, 1).value == 1 and ws.cell(first_data_row, 2).value == 2:
        first_data_row += 1

    return header_row, first_data_row


def build_upcoming_sheet(wb, upcoming_rows: list[dict], today: datetime.date) -> None:
    """Create or replace the 'Upcoming Hearings' worksheet for cases within the next 3 days."""
    sheet_title = "Upcoming Hearings"
    if sheet_title in wb.sheetnames:
        del wb[sheet_title]

    ws = wb.create_sheet(title=sheet_title)
    thin = Side(border_style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:J1")
    ws["A1"] = f"Upcoming Hearings (Next 3 Days: {today.strftime('%d.%m.%Y')} - {(today + datetime.timedelta(days=3)).strftime('%d.%m.%Y')})"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    headers = [
        "Sr. #",
        "Name of Council",
        "Case / Petition No.",
        "Filed by",
        "Respondents",
        "Name of Court",
        "Nature of Case",
        "Present Status",
        "Last Hearing",
        "Next Hearing",
    ]
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(2, col_idx, h)
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.fill = header_fill
        cell.border = border_all
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 24

    upcoming_rows.sort(key=lambda x: x["next_date"])

    if not upcoming_rows:
        ws.merge_cells("A3:J3")
        ws["A3"] = "No hearings scheduled within the next 3 days."
        ws["A3"].font = Font(name="Calibri", size=11, italic=True)
        ws["A3"].alignment = Alignment(horizontal="center", vertical="center")
    else:
        for idx, item in enumerate(upcoming_rows, start=1):
            r = idx + 2
            vals = [
                idx,
                item["council"],
                item["case_no_str"],
                item["filed_by"],
                item["respondents"],
                item["court"],
                item["nature"],
                item["status"],
                item["last_hearing"],
                item["next_date"].strftime("%d.%m.%Y"),
            ]
            for col_idx, val in enumerate(vals, start=1):
                c = ws.cell(r, col_idx, val)
                c.font = Font(name="Calibri", size=10)
                c.border = border_all
                c.alignment = Alignment(vertical="center", wrap_text=True)
                if col_idx == 10:
                    c.fill = FILL_URGENT
                    c.font = Font(name="Calibri", size=10, bold=True)

    widths = [7, 24, 26, 24, 28, 28, 18, 26, 14, 14]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


APP_DEFAULT_TITLE = "Sindh Courts Case Diary Updater"


class CourtUpdaterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.settings = load_settings()
        self.advocate_name = str(self.settings.get("advocate_name") or "").strip()
        self.advocate_subtitle = str(self.settings.get("advocate_subtitle") or "").strip()

        window_title = f"{APP_DEFAULT_TITLE} — {self.advocate_name}" if self.advocate_name else APP_DEFAULT_TITLE
        self.root.title(window_title)
        self.root.geometry("880x620")
        self.root.minsize(760, 520)
        self._apply_window_icon()

        self.stop_requested = False
        self.worker_thread = None

        self._build_ui()
        self._load_initial_file()

    def _apply_window_icon(self):
        for base_dir in (get_app_dir(), getattr(sys, "_MEIPASS", "")):
            if not base_dir:
                continue
            icon_path = os.path.join(base_dir, "app_icon.ico")
            if os.path.exists(icon_path):
                try:
                    self.root.iconbitmap(icon_path)
                    return
                except Exception:
                    pass

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        hdr = tk.Frame(self.root, bg="#1F4E79", pady=10)
        hdr.pack(fill="x")
        tk.Label(
            hdr,
            text=APP_DEFAULT_TITLE,
            font=("Segoe UI", 14, "bold"),
            fg="white",
            bg="#1F4E79",
        ).pack()

        if self.advocate_name and self.advocate_subtitle:
            subtitle_text = f"{self.advocate_name} • {self.advocate_subtitle}"
        elif self.advocate_name:
            subtitle_text = self.advocate_name
        else:
            subtitle_text = "High Court of Sindh & District Courts • Automated Case Status Sync"

        tk.Label(
            hdr,
            text=subtitle_text,
            font=("Segoe UI", 9),
            fg="#D9E1F2",
            bg="#1F4E79",
        ).pack()

        frm_file = ttk.LabelFrame(self.root, text="Excel Workbook & Worksheet", padding=10)
        frm_file.pack(fill="x", **pad)

        row1 = tk.Frame(frm_file)
        row1.pack(fill="x", pady=3)
        tk.Label(row1, text="Excel File:", width=12, anchor="w").pack(side="left")
        self.file_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.file_var, state="readonly").pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row1, text="Browse Excel File...", command=self.browse_excel).pack(side="left")

        row2 = tk.Frame(frm_file)
        row2.pack(fill="x", pady=3)
        tk.Label(row2, text="Worksheet Tab:", width=12, anchor="w").pack(side="left")
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(row2, textvariable=self.sheet_var, state="readonly", width=28)
        self.sheet_combo.pack(side="left", padx=5)

        self.visible_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            row2,
            text="Show Browser Window (Visible Mode — for CAPTCHA / visual check)",
            variable=self.visible_var,
        ).pack(side="left", padx=15)

        frm_ctrl = tk.Frame(self.root)
        frm_ctrl.pack(fill="x", **pad)
        self.btn_start = ttk.Button(frm_ctrl, text="▶ Start Updating Cases", command=self.start_update)
        self.btn_start.pack(side="left", padx=(0, 8))
        self.btn_stop = ttk.Button(frm_ctrl, text="■ Stop", command=self.request_stop, state="disabled")
        self.btn_stop.pack(side="left")

        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(frm_ctrl, textvariable=self.status_var, anchor="w", fg="#333333").pack(side="left", padx=14)

        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(fill="x", padx=12, pady=(0, 6))

        frm_log = ttk.LabelFrame(self.root, text="Live Update Log", padding=6)
        frm_log.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.log_box = scrolledtext.ScrolledText(frm_log, font=("Consolas", 9), state="disabled", wrap="word")
        self.log_box.pack(fill="both", expand=True)
        self.log_box.tag_config("green", foreground="#1E7E34")
        self.log_box.tag_config("red", foreground="#C82333")
        self.log_box.tag_config("yellow", foreground="#996500")
        self.log_box.tag_config("info", foreground="#1F4E79")

    def _load_initial_file(self):
        last_path = self.settings.get("last_excel_path", "")
        if last_path and os.path.exists(last_path):
            self.load_workbook_tabs(last_path, self.settings.get("last_sheet_name"))

    def browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Court Cases Excel File",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if path:
            self.load_workbook_tabs(path)

    def load_workbook_tabs(self, path: str, preferred_sheet: str | None = None):
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            sheets = [s for s in wb.sheetnames if s != "Upcoming Hearings"]
            wb.close()
        except Exception as e:
            messagebox.showerror("Error", f"Could not open Excel file:\n{e}")
            return

        self.file_var.set(path)
        self.sheet_combo["values"] = sheets
        if preferred_sheet and preferred_sheet in sheets:
            self.sheet_var.set(preferred_sheet)
        elif "Sheet2" in sheets:
            self.sheet_var.set("Sheet2")
        elif sheets:
            self.sheet_var.set(sheets[0])

        self.settings["last_excel_path"] = path
        self.settings["last_sheet_name"] = self.sheet_var.get()
        save_settings(self.settings)
        self.log(f"Loaded workbook: {path} (Sheet: {self.sheet_var.get()})", "info")

    def log(self, message: str, tag: str | None = None):
        def _append():
            self.log_box.configure(state="normal")
            if tag:
                self.log_box.insert("end", message + "\n", tag)
            else:
                self.log_box.insert("end", message + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")

        self.root.after(0, _append)

    def set_status(self, text: str):
        self.root.after(0, lambda: self.status_var.set(text))

    def set_progress(self, value: int, maximum: int):
        def _upd():
            self.progress["maximum"] = max(maximum, 1)
            self.progress["value"] = value

        self.root.after(0, _upd)

    def request_stop(self):
        self.stop_requested = True
        self.set_status("Stopping after current case...")

    def start_update(self):
        filepath = self.file_var.get().strip()
        sheet_name = self.sheet_var.get().strip()

        if not filepath or not os.path.exists(filepath):
            messagebox.showwarning("No File", "Please select a valid Excel file first.")
            return
        if not sheet_name:
            messagebox.showwarning("No Sheet", "Please select a worksheet tab.")
            return

        if is_excel_locked(filepath):
            messagebox.showerror(
                "Excel File is Open",
                "The selected Excel file appears to be open in Microsoft Excel.\n\n"
                "Please close it completely and try again.",
            )
            return

        self.settings["last_excel_path"] = filepath
        self.settings["last_sheet_name"] = sheet_name
        save_settings(self.settings)

        self.stop_requested = False
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")

        headless = not self.visible_var.get()
        self.worker_thread = threading.Thread(
            target=self._run_update_worker,
            args=(filepath, sheet_name, headless),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_update_worker(self, filepath: str, sheet_name: str, headless: bool):
        try:
            backup_path = create_backup(filepath)
            self.log(f"Backup created: {backup_path}", "info")

            wb = openpyxl.load_workbook(filepath)
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"Worksheet '{sheet_name}' not found.")
            ws = wb[sheet_name]

            summary = process_worksheet(
                ws,
                wb,
                headless=headless,
                log_fn=self.log,
                status_fn=self.set_status,
                progress_fn=self.set_progress,
                stop_check=lambda: self.stop_requested,
            )
            wb.save(filepath)

            self.log("\n" + summary, "info")
            self.set_status(summary)
            self.root.after(0, lambda: messagebox.showinfo("Update Complete", summary))

        except Exception as e:
            self.log(f"FATAL ERROR: {e}", "red")
            self.set_status("Error occurred.")
            self.root.after(0, lambda err=str(e): messagebox.showerror("Error", err))
        finally:
            self.root.after(0, lambda: self.btn_start.configure(state="normal"))
            self.root.after(0, lambda: self.btn_stop.configure(state="disabled"))


def process_worksheet(
    ws,
    wb,
    headless: bool = True,
    log_fn=None,
    status_fn=None,
    progress_fn=None,
    stop_check=None,
) -> str:
    """Process all case rows in worksheet `ws`, update wb in-place, and build Upcoming Hearings."""
    if log_fn is None:
        log_fn = lambda msg, tag=None: None
    if status_fn is None:
        status_fn = lambda msg: None
    if progress_fn is None:
        progress_fn = lambda v, m: None
    if stop_check is None:
        stop_check = lambda: False

    _, first_data_row = detect_sheet_layout(ws)
    today = datetime.date.today()

    target_rows = []
    for r in range(first_data_row, ws.max_row + 1):
        sr = ws.cell(r, 1).value
        col_c = ws.cell(r, 3).value
        if sr is None and col_c is None:
            continue
        if isinstance(sr, (int, float)) or col_c:
            target_rows.append(r)

    total = len(target_rows)
    progress_fn(0, total)

    updated_count = 0
    urgent_count = 0
    warn_count = 0
    upcoming_rows = []

    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(channel="chrome", headless=headless)
            log_fn("Launched browser: Google Chrome", "info")
        except Exception:
            browser = p.chromium.launch(channel="msedge", headless=headless)
            log_fn("Launched browser: Microsoft Edge (fallback)", "info")

        context = browser.new_context(viewport={"width": 1280, "height": 850})
        page = context.new_page()

        for idx, r in enumerate(target_rows, start=1):
            if stop_check():
                log_fn("Update stopped by user.", "yellow")
                break

            sr_val = ws.cell(r, 1).value
            council = ws.cell(r, 2).value or ""
            raw_case = ws.cell(r, 3).value
            filed_by = ws.cell(r, 4).value or ""
            respondents = ws.cell(r, 5).value or ""
            court_name = ws.cell(r, 6).value or ""
            old_last_raw = ws.cell(r, 9).value
            old_next_raw = ws.cell(r, 10).value
            old_next_date = parse_date_flexible(old_next_raw)

            status_fn(f"Checking {idx}/{total}: {raw_case or 'Empty Case No.'}")
            progress_fn(idx, total)

            parsed = parse_case_string(raw_case, str(court_name))
            if not parsed or parsed["is_crime"]:
                reason = (
                    "Missing Case No."
                    if not raw_case
                    else ("Crime/FIR No. (manual check)" if (parsed and parsed["is_crime"]) else f"Unparsed: {raw_case}")
                )
                apply_neighbour_style(ws, r, 10, first_data_row)
                ws.cell(r, 10).fill = FILL_WARN
                warn_count += 1
                log_fn(f"[Row {r} | Sr {sr_val}] {reason} -> Highlighted yellow", "yellow")
                continue

            portal = parsed["portal"]
            res = _empty_portal_result()
            try:
                if portal == "shc_hyd":
                    res = query_shc_portal(page, SHC_HYD_URL, parsed, str(filed_by), str(respondents))
                    if not res["found"]:
                        res = query_shc_portal(page, SHC_KHI_URL, parsed, str(filed_by), str(respondents))
                elif portal == "shc_khi":
                    res = query_shc_portal(page, SHC_KHI_URL, parsed, str(filed_by), str(respondents))
                else:
                    res = query_district_portal(
                        page, parsed, str(filed_by), str(respondents), str(court_name)
                    )
            except PlaywrightTimeoutError:
                res = _empty_portal_result("Timeout")
            except Exception as exc:
                res = _empty_portal_result(f"Error: {exc}")

            # STRICT NON-DESTRUCTIVE RULE: Fill Columns D..H ONLY when cell is completely empty
            fill_map = {
                4: res.get("petitioner") or parsed.get("inline_petitioner", ""),
                5: res.get("respondent") or parsed.get("inline_respondent", ""),
                6: res.get("court_name", ""),
                7: res.get("nature_of_case", ""),
                8: res.get("present_status", ""),
            }
            filled_cols = []
            for col_idx, candidate in fill_map.items():
                cell = ws.cell(r, col_idx)
                if is_cell_empty(cell) and candidate and str(candidate).strip():
                    cell.value = str(candidate).strip()
                    apply_neighbour_style(ws, r, col_idx, first_data_row)
                    filled_cols.append(get_column_letter(col_idx))
            if filled_cols:
                log_fn(f"[Row {r} | Sr {sr_val}] Filled blank cells {','.join(filled_cols)}", "info")

            # ALWAYS POPULATE LAST HEARING (Column I) & UPDATE NEXT HEARING (Column J)
            new_date = res.get("next_date")
            portal_last_date = res.get("last_date")
            portal_status = res.get("status", "")
            final_date = old_next_date
            was_updated = False

            if new_date is not None and (old_next_date is None or new_date > old_next_date):
                # Choose newer between old Col J date and portal last_date (as long as < new_date)
                chosen_last_dt = old_next_date
                if portal_last_date is not None and portal_last_date < new_date:
                    if chosen_last_dt is None or portal_last_date > chosen_last_dt:
                        chosen_last_dt = portal_last_date

                if chosen_last_dt is not None:
                    apply_neighbour_style(ws, r, 9, first_data_row)
                    ws.cell(r, 9).value = chosen_last_dt.strftime("%d.%m.%Y")
                    ws.cell(r, 9).fill = FILL_UPDATED
                elif old_next_raw is not None and str(old_next_raw).strip().lower() not in ("", "not available"):
                    apply_neighbour_style(ws, r, 9, first_data_row)
                    ws.cell(r, 9).value = str(old_next_raw).strip()
                    ws.cell(r, 9).fill = FILL_UPDATED

                apply_neighbour_style(ws, r, 10, first_data_row)
                ws.cell(r, 10).value = new_date.strftime("%d.%m.%Y")
                ws.cell(r, 10).fill = FILL_UPDATED
                final_date = new_date
                was_updated = True
                updated_count += 1
                log_fn(
                    f"[Row {r} | Sr {sr_val}] UPDATED {raw_case}: {old_next_raw} -> {new_date.strftime('%d.%m.%Y')} ({portal_status})",
                    "green",
                )
            elif is_cell_empty(ws.cell(r, 9)) and portal_last_date is not None:
                # Col I was blank -> always populate it with portal's last_date
                apply_neighbour_style(ws, r, 9, first_data_row)
                ws.cell(r, 9).value = portal_last_date.strftime("%d.%m.%Y")
                ws.cell(r, 9).fill = FILL_UPDATED
                log_fn(
                    f"[Row {r} | Sr {sr_val}] Populated blank Last Hearing -> {portal_last_date.strftime('%d.%m.%Y')}",
                    "green",
                )

            # Ensure Col J styling matches neighbours on newly added rows without borders
            b10 = ws.cell(r, 10).border
            if not (b10 and ((b10.top and b10.top.style) or (b10.bottom and b10.bottom.style))):
                apply_neighbour_style(ws, r, 10, first_data_row)

            # Re-read post-fill row values for Upcoming Hearings sheet
            cur_filed_by = ws.cell(r, 4).value or ""
            cur_respondents = ws.cell(r, 5).value or ""
            cur_court = ws.cell(r, 6).value or ""
            cur_nature = ws.cell(r, 7).value or ""
            cur_status = ws.cell(r, 8).value or ""
            cur_last_hearing = str(ws.cell(r, 9).value or "")

            if final_date is not None:
                days_diff = (final_date - today).days
                if 0 <= days_diff <= 3:
                    ws.cell(r, 10).fill = FILL_URGENT
                    urgent_count += 1
                    upcoming_rows.append(
                        {
                            "council": council,
                            "case_no_str": str(raw_case or ""),
                            "filed_by": cur_filed_by,
                            "respondents": cur_respondents,
                            "court": cur_court,
                            "nature": cur_nature,
                            "status": cur_status,
                            "last_hearing": cur_last_hearing,
                            "next_date": final_date,
                        }
                    )
                    log_fn(
                        f"   -> URGENT: Hearing in {days_diff} day(s) on {final_date.strftime('%d.%m.%Y')}",
                        "red",
                    )
                elif days_diff < 0:
                    ws.cell(r, 10).fill = FILL_WARN
                    warn_count += 1
                    if not was_updated:
                        log_fn(
                            f"[Row {r} | Sr {sr_val}] OVERDUE/PAST ({final_date.strftime('%d.%m.%Y')}) | Portal: {portal_status}",
                            "yellow",
                        )
                elif not was_updated:
                    log_fn(
                        f"[Row {r} | Sr {sr_val}] Unchanged ({final_date.strftime('%d.%m.%Y')}) | Portal: {portal_status}"
                    )
            else:
                ws.cell(r, 10).fill = FILL_WARN
                warn_count += 1
                log_fn(
                    f"[Row {r} | Sr {sr_val}] No date found for {raw_case} ({portal_status}) -> Highlighted yellow",
                    "yellow",
                )

        browser.close()

    build_upcoming_sheet(wb, upcoming_rows, today)
    return (
        f"Done! Updated: {updated_count} | Within 3 Days: {urgent_count} | "
        f"Overdue/Unfound: {warn_count} | 'Upcoming Hearings' tab updated."
    )


def main():
    root = tk.Tk()
    CourtUpdaterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
