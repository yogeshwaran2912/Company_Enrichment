"""
================================================================================
LIVE DESKTOP CHROME AUTOMATION - COMPANY CORE BUSINESS ENRICHMENT
================================================================================

DESCRIPTION:
    Automates Google Chrome directly on your desktop screen so you can watch it
    work live! For each company in your Excel file:
      1. Reads the company name from Excel.
      2. Adds the new column 'Core Business' if not already present.
      3. Types the search into a real, visible Google Chrome browser window on your desktop.
      4. Searches live for what the company does (products, services, core business).
      5. Extracts a clean, concise, 1-2 sentence summary of their core business.
      6. Writes the core business back into the Excel spreadsheet row by row.
      7. Saves progress incrementally every N rows (default: 5).
      8. Safe Resume: Skips rows where 'Core Business' is already populated.

HOW TO RUN:
    Simply open PowerShell or Command Prompt and run:
        python enrich_core_business.py

    This automatically launches a visible Google Chrome window maximized on your desktop
    and enriches:
        C:/Users/Elakkiya Tamilselvan/Desktop/chennai companys.xlsx

OPTIONS:
    Test on just 5 companies first:
        python enrich_core_business.py --limit 5

    Specify a different Excel file:
        python enrich_core_business.py --file "C:/path/to/my_file.xlsx"

    Run silently in background (headless):
        python enrich_core_business.py --headless

DEPENDENCIES:
    pip install openpyxl selenium webdriver-manager python-dotenv
================================================================================
"""

import os
import re
import sys
import time
import logging
import argparse
from typing import Dict, Optional, Tuple, List, Any

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Alignment
from dotenv import load_dotenv

from core_business_extractor import extract_short_core_business

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

load_dotenv()

# Logging configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("core_business_enricher")

# Silence noisy third-party loggers
for mod in ("selenium", "urllib3", "webdriver_manager"):
    logging.getLogger(mod).setLevel(logging.WARNING)

COMPANY_HEADER_CANDIDATES = [
    "company name", "company_name", "company", "name", "organization",
    "firm", "account name", "companies"
]

CORE_BUSINESS_HEADER_CANDIDATES = [
    "core business", "core_business", "nature of business", "business overview",
    "business description", "industry / business", "what they do"
]


def clean_company_name(name: Any) -> str:
    """Cleans whitespace, quotes, and punctuation from company names."""
    if not name:
        return ""
    cleaned = str(name).strip().strip('\'"“”')
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned


def clean_snippet_body(snippet: str, title: str = "", company_name: str = "") -> Optional[str]:
    """
    Extracts a short, concise, professional core business category (e.g. 'IT Infrastructure',
    'Industrial Automation', 'Advertising Technology (AdTech)').
    """
    if not snippet and not title and not company_name:
        return None
    return extract_short_core_business(snippet, title=title, company_name=company_name)


class LiveBrowserAutomator:
    """Automates visible Google Chrome directly on user's desktop."""

    def __init__(self, headless: bool = False):
        self.headless = headless
        self.driver: Optional[webdriver.Chrome] = None

    def start(self) -> None:
        """Launches the Chrome browser window maximized on desktop."""
        logger.info("Launching Google Chrome browser on desktop...")
        options = Options()
        if self.headless:
            options.add_argument('--headless=new')
        else:
            options.add_argument('--start-maximized')

        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        options.add_experimental_option('useAutomationExtension', False)

        self.driver = webdriver.Chrome(options=options)
        self.driver.get('https://www.google.com')
        time.sleep(1.5)

    def close(self) -> None:
        """Quits the browser."""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    def type_and_search(self, query: str) -> None:
        """Visibly types query into the search box with keyboard simulation."""
        box = None
        for selector in [(By.NAME, 'q'), (By.CSS_SELECTOR, 'textarea[name="q"]'), (By.CSS_SELECTOR, 'input[name="q"]')]:
            try:
                box = self.driver.find_element(*selector)
                break
            except Exception:
                continue

        if not box:
            self.driver.get('https://www.google.com')
            time.sleep(1.0)
            box = self.driver.find_element(By.NAME, 'q')


        # Clear existing text
        box.send_keys(Keys.CONTROL + "a")
        box.send_keys(Keys.BACKSPACE)
        time.sleep(0.1)

        # Type character by character visibly
        for char in query:
            box.send_keys(char)
            time.sleep(0.012)

        box.send_keys(Keys.RETURN)
        time.sleep(2.5)

    def search_core_business(self, company_name: str) -> Optional[str]:
        """
        Visibly types query into Chrome, searches live, and extracts a short, concise
        core business category (e.g. 'IT Infrastructure', 'Industrial Automation').
        """
        if not self.driver:
            self.start()

        clean_comp = clean_company_name(company_name)
        # Suffix-free brand name for cleaner queries
        suffix_free = re.sub(
            r'\b(?:pvt\.?\s*ltd\.?|private\s+limited|p\s+ltd|ltd\.?|limited|inc\.?|llc|corp\.?|co\.?)\b',
            '', clean_comp, flags=re.IGNORECASE
        ).strip()
        search_brand = suffix_free if len(suffix_free) >= 3 else clean_comp

        queries = [
            f'"{search_brand}" company core business products services',
            f'"{clean_comp}" what we do overview',
            f'{search_brand} Chennai products services about us',
            f'{search_brand} Chennai'
        ]

        collected_snippets = []

        for query in queries:
            logger.info(f"   [Chrome Screen] Searching: {query}")
            try:
                self.type_and_search(query)

                # Collect result articles
                articles = self.driver.find_elements(
                    By.CSS_SELECTOR,
                    'article[data-testid="result"], article, div[data-layout="organic"]'
                )

                for item in articles[:6]:
                    lines = [l.strip() for l in item.text.split('\n') if l.strip()]
                    title = lines[0] if lines else ""
                    if len(lines) >= 4:
                        body = max(lines[3:], key=len)
                    elif len(lines) >= 3:
                        body = lines[2]
                    elif lines:
                        body = lines[-1]
                    else:
                        body = ""

                    collected_snippets.append((title, body))

                    short_biz = clean_snippet_body(body, title=title, company_name=clean_comp)
                    if short_biz:
                        return short_biz

            except Exception as e:
                logger.warning(f"Browser search error for {query}: {e}")

        # If not matched directly from individual articles, test all gathered snippets
        all_text = " ".join([f"{t} {b}" for t, b in collected_snippets])
        fallback = extract_short_core_business(all_text, company_name=clean_comp)
        if fallback:
            return fallback

        # Fallback to company name taxonomy triggers
        return extract_short_core_business("", company_name=clean_comp)


class CoreBusinessEnricher:
    """Manages Excel file enrichment with live Chrome visualization for Core Business."""

    def __init__(
        self,
        file_path: str,
        sheet_name: Optional[str] = None,
        headless: bool = False,
        save_interval: int = 5,
        delay_seconds: float = 1.5,
        limit: Optional[int] = None
    ):
        self.file_path = os.path.abspath(file_path)
        self.sheet_name = sheet_name
        self.headless = headless
        self.save_interval = max(1, save_interval)
        self.delay_seconds = delay_seconds
        self.limit = limit

        self.wb: Optional[openpyxl.Workbook] = None
        self.sheet: Optional[Worksheet] = None
        self.company_col: int = 1
        self.core_biz_col: int = 2
        self.is_dirty: bool = False
        self.automator = LiveBrowserAutomator(headless=self.headless)

        self.stats = {
            "total_rows": 0,
            "empty_rows": 0,
            "skipped_already_done": 0,
            "processed": 0,
            "found": 0,
            "not_found": 0
        }

    def prepare_excel(self) -> None:
        """Loads spreadsheet and configures 'Core Business' column."""
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"Excel file not found at: {self.file_path}")

        logger.info(f"Opening Excel file: {self.file_path}")
        self.wb = openpyxl.load_workbook(self.file_path)

        if self.sheet_name and self.sheet_name in self.wb.sheetnames:
            self.sheet = self.wb[self.sheet_name]
        else:
            self.sheet = self.wb.active

        logger.info(f"Using Sheet: '{self.sheet.title}' (Total Rows: {self.sheet.max_row})")

        headers = {}
        for c in range(1, self.sheet.max_column + 1):
            val = self.sheet.cell(1, c).value
            if val is not None:
                headers[c] = str(val).strip().lower()

        # Find Company Column
        comp_col = None
        for c, text in headers.items():
            for cand in COMPANY_HEADER_CANDIDATES:
                if text == cand or text.startswith(cand):
                    comp_col = c
                    break
            if comp_col:
                break
        self.company_col = comp_col or 1

        # Find or create 'Core Business' column
        biz_col = None
        for c, text in headers.items():
            for cand in CORE_BUSINESS_HEADER_CANDIDATES:
                if text == cand or text.startswith(cand):
                    biz_col = c
                    break
            if biz_col:
                break

        if biz_col:
            self.core_biz_col = biz_col
            col_letter = openpyxl.utils.get_column_letter(biz_col)
            logger.info(f"Found existing 'Core Business' column at Column {col_letter}")
        else:
            # Append new 'Core Business' column at the end
            curr_max = self.sheet.max_column + 1
            new_cell = self.sheet.cell(1, curr_max, value="Core Business")
            new_cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            new_cell.fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            new_cell.alignment = Alignment(horizontal="center", vertical="center")
            col_letter = openpyxl.utils.get_column_letter(curr_max)
            self.sheet.column_dimensions[col_letter].width = 38
            self.core_biz_col = curr_max
            self.is_dirty = True
            logger.info(f"Added new column 'Core Business' at Column {col_letter}")

        logger.info(
            "Column Mapping -> Company: Col %s ('%s'), Core Business: Col %s ('Core Business')",
            openpyxl.utils.get_column_letter(self.company_col),
            self.sheet.cell(1, self.company_col).value,
            openpyxl.utils.get_column_letter(self.core_biz_col)
        )

        if self.is_dirty:
            self.save()

    def process_row(self, row_idx: int) -> bool:
        """Processes a single company row via visible Chrome."""
        raw_company = self.sheet.cell(row=row_idx, column=self.company_col).value
        cleaned_company = clean_company_name(raw_company)

        if not cleaned_company:
            self.stats["empty_rows"] += 1
            return False

        current_val = self.sheet.cell(row=row_idx, column=self.core_biz_col).value
        if current_val and str(current_val).strip():
            val_str = str(current_val).strip()
            # If existing core business is a long sentence, shorten it!
            if len(val_str.split()) > 5 or len(val_str) > 40:
                short_biz = extract_short_core_business(val_str, company_name=cleaned_company)
                if short_biz and short_biz != val_str:
                    logger.info(f"[Row {row_idx}] Shortened existing long Core Business:")
                    logger.info(f"   Old: {val_str[:60]}...")
                    logger.info(f"   New: {short_biz}")
                    self.sheet.cell(row=row_idx, column=self.core_biz_col, value=short_biz)
                    self.is_dirty = True
            self.stats["skipped_already_done"] += 1
            return False

        total_rows = self.sheet.max_row
        print("-" * 65)
        logger.info(f"[Row {row_idx}/{total_rows}] Processing Company: '{cleaned_company}'")
        self.stats["processed"] += 1

        core_biz = self.automator.search_core_business(cleaned_company)

        if core_biz:
            self.sheet.cell(row=row_idx, column=self.core_biz_col, value=core_biz)
            logger.info(f"   [FOUND CORE BUSINESS] -> {core_biz}")
            self.stats["found"] += 1
            self.is_dirty = True
        else:
            logger.info("   [CORE BUSINESS]       -> [Not Found]")
            self.stats["not_found"] += 1

        return True

    def save(self) -> None:
        """Saves Excel file incrementally, with fallback if open in Excel."""
        if self.wb:
            try:
                self.wb.save(self.file_path)
                self.is_dirty = False
                logger.info(f"[EXCEL SAVED] Progress saved to '{self.file_path}'")
            except PermissionError:
                base, ext = os.path.splitext(self.file_path)
                fallback_path = f"{base}_updated{ext}"
                logger.warning(
                    f"\n[FILE LOCKED] '{os.path.basename(self.file_path)}' is currently open in Microsoft Excel!\n"
                    f"   -> Saving progress to '{os.path.basename(fallback_path)}' so no data is lost.\n"
                    f"   -> (TIP: Close the file in Excel to allow the script to save directly to the original file.)"
                )
                try:
                    self.wb.save(fallback_path)
                    logger.info(f"[EXCEL SAVED] Progress secured in '{fallback_path}'")
                except Exception as e2:
                    logger.error(f"Fallback save error: {e2}")
            except Exception as e:
                logger.error(f"Error saving Excel file: {e}")

    def run(self) -> None:
        """Runs the loop through all Excel rows."""
        self.prepare_excel()
        self.automator.start()

        max_row = self.sheet.max_row
        processed_count = 0
        since_last_save = 0

        logger.info(f"Starting live Chrome automation for {max_row - 1} rows...")

        try:
            for row_idx in range(2, max_row + 1):
                self.stats["total_rows"] += 1

                if self.limit and processed_count >= self.limit:
                    logger.info(f"Reached limit of {self.limit} companies. Stopping.")
                    break

                try:
                    searched = self.process_row(row_idx)

                    if searched:
                        processed_count += 1
                        since_last_save += 1

                        if since_last_save >= self.save_interval and self.is_dirty:
                            self.save()
                            since_last_save = 0

                        if row_idx < max_row and (not self.limit or processed_count < self.limit):
                            time.sleep(self.delay_seconds)

                except KeyboardInterrupt:
                    logger.warning("\nExecution stopped by user (Ctrl+C). Saving...")
                    break
                except Exception as e:
                    logger.error(f"[Row {row_idx}] Error: {e}")

        finally:
            if self.is_dirty:
                self.save()
            self.automator.close()
            self.print_summary()

    def print_summary(self) -> None:
        """Prints summary statistics."""
        print("\n" + "=" * 65)
        print("           LIVE CORE BUSINESS ENRICHMENT SUMMARY")
        print("=" * 65)
        print(f"  Total Rows Inspected:       {self.stats['total_rows']}")
        print(f"  Empty Rows Skipped:         {self.stats['empty_rows']}")
        print(f"  Already Enriched (Skipped): {self.stats['skipped_already_done']}")
        print(f"  Companies Processed:        {self.stats['processed']}")
        print(f"  Core Business Found:        {self.stats['found']}")
        print(f"  Not Found:                  {self.stats['not_found']}")
        print("=" * 65)


def find_default_excel() -> str:
    desktop_file = r"C:\Users\ADV PRO\Desktop\company_enrichment\chennai companys.xlsx"
    if os.path.exists(desktop_file):
        return desktop_file
    if os.path.exists("companies.xlsx"):
        return "companies.xlsx"
    if os.path.exists("companies_sample.xlsx"):
        return "companies_sample.xlsx"
    return desktop_file


def main():
    parser = argparse.ArgumentParser(
        description="Live desktop Google Chrome automation for adding & enriching 'Core Business' in Excel."
    )
    parser.add_argument(
        "-f", "--file",
        dest="file_path",
        default=os.getenv("EXCEL_FILE_PATH", find_default_excel()),
        help="Path to the Excel file (.xlsx)."
    )
    parser.add_argument(
        "-s", "--sheet",
        dest="sheet_name",
        default=os.getenv("SHEET_NAME", None),
        help="Sheet name to process."
    )
    parser.add_argument(
        "-n", "--save-interval",
        dest="save_interval",
        type=int,
        default=int(os.getenv("SAVE_INTERVAL", "5")),
        help="Save to Excel every N processed rows (default: 5)."
    )
    parser.add_argument(
        "-d", "--delay",
        dest="delay",
        type=float,
        default=float(os.getenv("SEARCH_DELAY", "1.5")),
        help="Delay in seconds between company searches (default: 1.5s)."
    )
    parser.add_argument(
        "-l", "--limit",
        dest="limit",
        type=int,
        default=None,
        help="Limit number of companies to process (useful for testing, e.g. --limit 5)."
    )
    parser.add_argument(
        "--headless",
        dest="headless",
        action="store_true",
        default=False,
        help="Run Chrome silently in background (headless mode)."
    )

    args = parser.parse_args()

    enricher = CoreBusinessEnricher(
        file_path=args.file_path,
        sheet_name=args.sheet_name,
        headless=args.headless,
        save_interval=args.save_interval,
        delay_seconds=args.delay,
        limit=args.limit
    )

    enricher.run()


if __name__ == "__main__":
    main()
