"""
================================================================================
DUCKDUCKGO SEARCH AUTOMATION - COMPANY WEBSITE & NUMBER OF EMPLOYEES
================================================================================

DESCRIPTION:
    Automates DuckDuckGo search to extract official company website URLs and
    number of employees for each company in your Excel file (e.g. companies_sample.xlsx):

    1. Company Website:
       - Visibly navigates DuckDuckGo using the exact query format:
         Query: [company name] website url  (e.g. 'AMPLE TECHNOLOGIES website url')
       - Identifies the official company website URL (filtering out directories,
         social media, and aggregators like LinkedIn, Justdial, ZaubaCorp, IndiaMart).
       - Copies the clean official URL and pastes it into the 'Company Website' column.

    2. Number of Employees:
       - Visibly navigates DuckDuckGo using the exact query format:
         Query: [company name] no of employees  (e.g. 'AMPLE TECHNOLOGIES no of employees')
       - Extracts employee count / company size (e.g. '1,436 employees', '3,939 employees', '284 employees')
         from verified search result snippets.
       - Copies the headcount and pastes it into the 'No of Employees' column.

    3. Excel Updates & Safety:
       - Automatically maps or creates 'Company Website' and 'No of Employees' columns.
       - Saves progress incrementally after each company.
       - Safe Resume: Skips rows where both fields are already populated.
       - Excel Lock Protection: If the file is open in Microsoft Excel, safely saves
         to <file>_updated.xlsx so no data is lost.
       - Strictly DuckDuckGo: DOES NOT use Google to search.

HOW TO RUN:
    python enrich_website_and_employees.py

OPTIONS:
    Test on 5 companies first:
        python enrich_website_and_employees.py --limit 5

    Specify a different Excel file:
        python enrich_website_and_employees.py --file "companies_sample.xlsx"

    Run in background without opening visible browser window:
        python enrich_website_and_employees.py --headless

    Re-search and overwrite existing fields in Excel:
        python enrich_website_and_employees.py --overwrite

    Custom auto-save interval and delay:
        python enrich_website_and_employees.py --save-interval 1 --delay 1.0

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
from typing import Dict, Optional, List, Any
from urllib.parse import urlparse, quote_plus

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Alignment
from dotenv import load_dotenv

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
logger = logging.getLogger("duckduckgo_enricher")

# Silence noisy third-party loggers
for mod in ("selenium", "urllib3", "webdriver_manager"):
    logging.getLogger(mod).setLevel(logging.WARNING)

# Header candidates for dynamic column detection
COMPANY_HEADER_CANDIDATES = [
    "name", "company", "company name", "company_name", "organization",
    "firm", "account name", "companies"
]

WEBSITE_HEADER_CANDIDATES = [
    "company website", "companys website", "company's website",
    "website", "company url", "web address", "official website", "url"
]

EMPLOYEES_HEADER_CANDIDATES = [
    "no of employees", "number of employees", "employees", "employee count",
    "company size", "headcount", "no. of employees", "total employees"
]

# Domains to ignore when identifying official company website
IGNORE_DOMAINS = {
    'linkedin.com', 'facebook.com', 'twitter.com', 'x.com', 'instagram.com', 'youtube.com',
    'wikipedia.org', 'zaubacorp.com', 'tofler.in', 'instafinancials.com', 'justdial.com',
    'indiamart.com', 'glassdoor.com', 'ambitionbox.com', 'indeed.com', 'crunchbase.com',
    'zoominfo.com', 'yellowpages.com', 'quora.com', 'naukri.com', 'mapsofindia.com',
    'google.com', 'duckduckgo.com', 'bing.com', 'bloomberg.com', 'tracxn.com',
    'tutorialspoint.com', 'geeksforgeeks.org', 'w3schools.com', 'javatpoint.com',
    'reddit.com', 'pinterest.com', 'github.com', 'medium.com', 'tradeindia.com',
    'quickr.com', 'sulekha.com', 'mouthshut.com', 'economictimes.indiatimes.com',
    'moneycontrol.com', 'livemint.com', 'business-standard.com', 'thehindubusinessline.com',
    'pitchbook.com', 'owler.com', 'apollo.io', 'signalhire.com', 'rocketreach.co',
    'yahoo.com', 'msn.com', 'apple.com', 'microsoft.com', 'easyleadz.com', 'craft.co',
    'goldenchennai.com', 'cybo.com', 'indiabizclub.com', 'dial4trade.com', 'indiacom.com',
    'companieshouse.gov.uk', 'dnb.com', 'kompass.com', 'yelp.com', '99centswebhost.com',
    'reveliolabs.com', 'emis.com', 'datanyze.com', 'vymaps.com', 'findglocal.com'
}

PORTAL_SUBDOMAINS = (
    "smartpay.", "smarthro.", "hr.", "login.", "portal.", "mail.",
    "webmail.", "beta.", "app.", "staging.", "auth.", "sso."
)


def clean_company_name(name: Any) -> str:
    """Cleans whitespace, quotes, and trailing punctuation from company names."""
    if not name:
        return ""
    cleaned = str(name).strip().strip('\'"“”')
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned


def extract_brand_name(company_name: str) -> str:
    """Extracts clean brand name without corporate legal suffixes (Pvt Ltd, LLC, etc.)."""
    clean = clean_company_name(company_name)
    suffix_free = re.sub(
        r'\b(?:pvt\.?\s*ltd\.?|private\s+limited|p\s+ltd|ltd\.?|limited|inc\.?|llc|corp\.?|co\.?|\(i\))\b',
        '', clean, flags=re.IGNORECASE
    ).strip()
    suffix_free = re.sub(r'[\s.,&;/-]+$', '', suffix_free).strip()
    return suffix_free if len(suffix_free) >= 3 else clean


def clean_url(raw_url: str) -> Optional[str]:
    """Cleans and normalizes candidate URL to root domain."""
    if not raw_url:
        return None
    url = raw_url.strip()
    if not url.startswith("http"):
        return None

    try:
        parsed = urlparse(url)
        netloc = parsed.netloc.lower()
        if netloc.startswith("www."):
            netloc = netloc[4:]

        # Ignore common aggregators, directories, search engines
        if any(ign in netloc for ign in IGNORE_DOMAINS):
            return None

        # Clean portal/hr/login subdomains
        for p_sub in PORTAL_SUBDOMAINS:
            if netloc.startswith(p_sub):
                netloc = netloc[len(p_sub):]
                break

        # Ignore login/cart paths
        path = parsed.path.lower()
        if any(p in path for p in ("/login", "/signin", "/signup", "/register", "/cart", "/search")):
            return None

        clean_root = f"{parsed.scheme}://www.{netloc}" if not netloc.startswith("www.") else f"{parsed.scheme}://{netloc}"
        return clean_root
    except Exception:
        return None


def extract_website_from_urls(brand: str, urls: List[str]) -> Optional[str]:
    """
    Selects the best official company website from candidate URLs.
    Scores URLs higher if domain matches company brand tokens.
    """
    if not urls:
        return None

    brand_tokens = [t.lower() for t in re.split(r'[^a-zA-Z0-9]', brand) if len(t) >= 3]
    best_url = None
    best_score = -1

    for raw_url in urls:
        root_url = clean_url(raw_url)
        if not root_url:
            continue

        try:
            parsed = urlparse(root_url)
            netloc = parsed.netloc.lower()

            score = 1
            for token in brand_tokens:
                if token in netloc:
                    score += 30
                elif token in parsed.path:
                    score += 5

            if score > best_score:
                best_score = score
                best_url = root_url

        except Exception:
            continue

    return best_url


def extract_employees_from_text(text: str) -> Optional[str]:
    """
    Extracts number of employees / company size from snippet text.
    Handles ranges ('501-1,000 employees'), exact counts ('447 employees'),
    labeled patterns ('Company size: 1,001-5,000 employees'),
    and sentence patterns ('employs 1,436 people', 'has around 812 employees').
    """
    if not text:
        return None

    # 1. Standard employee range: e.g. '51-200 employees', '501-1,000 employees', '1,200+ employees', '10,001+ employees'
    m_range = re.search(
        r'\b([0-9]{1,3}(?:,[0-9]{3})*\s*(?:-|to)\s*[0-9]{1,3}(?:,[0-9]{3})*|\b[0-9]{1,3}(?:,[0-9]{3})*\+)\s*(?:employees|employee|staff|people|workforce)\b',
        text, re.IGNORECASE
    )
    if m_range:
        val = m_range.group(1).strip()
        return f"{val} employees"

    # 2. Labeled company size / headcount: e.g. 'Company size: 501-1,000 employees' or 'Headcount: 250'
    m_label = re.search(
        r'(?:company size|headcount|workforce|employee count|team size|size|no\.?\s*of employees)[:\s]+([0-9,+-]+(?:\s*(?:to|-)\s*[0-9,+-]+)?(?:\s*employees)?)',
        text, re.IGNORECASE
    )
    if m_label:
        val = m_label.group(1).strip()
        if 'employee' not in val.lower():
            val += " employees"
        return val

    # 3. Employs pattern: 'employs 1,436 people', 'currently employs 501 - 1,000 people'
    m_employs = re.search(
        r'\b(?:employs|employed|staff of|team of)\s+([0-9]{1,3}(?:,[0-9]{3})*|[0-9]+\s*-\s*[0-9]+)\s*(?:people|employees|staff|members)?\b',
        text, re.IGNORECASE
    )
    if m_employs:
        val = m_employs.group(1).strip()
        return f"{val} employees"

    # 4. Sentence pattern: 'has around 812 employees' or 'employee count of 812' or 'has 1097 employees'
    m_sentence = re.search(
        r'(?:employee count|headcount|workforce|team size)(?:[^\.\n]{0,80}?(?:is|was|of|:|=))\s*([0-9]{1,3}(?:,[0-9]{3})*|[0-9]+(?:\s*-\s*[0-9]+|\+)?)\b',
        text, re.IGNORECASE
    )
    if m_sentence:
        val = m_sentence.group(1).strip()
        return f"{val} employees" if "employee" not in val.lower() else val

    # 5. Direct count: e.g. '4,820 employees', '447 employees', '283 employees'
    m_count = re.search(r'\b([0-9]{1,3}(?:,[0-9]{3})*)\s+(?:employees|employee)\b', text, re.IGNORECASE)
    if m_count:
        val = m_count.group(1).strip()
        try:
            digits = int(val.replace(',', ''))
            if digits > 0:
                return f"{val} employees"
        except ValueError:
            pass

    return None


class DuckDuckGoAutomator:
    """Automates Chrome using DuckDuckGo (Strictly NO Google)."""

    def __init__(self, headless: bool = False):
        self.headless = headless
        self.driver: Optional[webdriver.Chrome] = None
        self.profile_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "chrome_ddg_profile"))

    def start(self) -> None:
        """Launches Chrome for DuckDuckGo automation."""
        mode_str = "silent background mode" if self.headless else "visible desktop window"
        logger.info(f"Launching Chrome for DuckDuckGo in {mode_str}...")
        os.makedirs(self.profile_dir, exist_ok=True)

        options = Options()
        if self.headless:
            options.add_argument('--headless=new')
        else:
            options.add_argument('--start-maximized')

        options.add_argument(f'--user-data-dir={self.profile_dir}')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
        )

        self.driver = webdriver.Chrome(options=options)
        self.driver.get('https://duckduckgo.com')
        time.sleep(1.0)

    def close(self) -> None:
        """Closes the browser."""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    def navigate_search(self, query: str) -> None:
        """Visibly types into DuckDuckGo search or navigates to search URL."""
        if not self.driver:
            self.start()

        # Visibly type into search bar if present
        box = None
        for selector in ['input[name="q"]', 'input#searchbox_input', 'input[type="text"]']:
            try:
                box = self.driver.find_element(By.CSS_SELECTOR, selector)
                break
            except Exception:
                continue

        if box:
            try:
                box.send_keys(Keys.CONTROL + "a")
                box.send_keys(Keys.BACKSPACE)
                time.sleep(0.05)
                for char in query:
                    box.send_keys(char)
                    time.sleep(0.01)
                box.send_keys(Keys.RETURN)
                time.sleep(2.0)
                return
            except Exception:
                pass

        # Direct navigation to DuckDuckGo search URL
        url = f"https://duckduckgo.com/?q={quote_plus(query)}"
        self.driver.get(url)
        time.sleep(2.0)

    def search_website(self, company_name: str) -> Optional[str]:
        """
        Searches DuckDuckGo for the company's official website.
        Query format: [company name] website url
        """
        clean_comp = clean_company_name(company_name)
        brand = extract_brand_name(clean_comp)

        # Exact requested format: [company name] website url
        queries = [f"{clean_comp} website url"]
        if brand.lower() != clean_comp.lower():
            queries.append(f"{brand} website url")

        for query in queries:
            logger.info(f"   [DuckDuckGo] Typing Query for Website: '{query}'")
            try:
                self.navigate_search(query)

                candidate_urls = []
                # Result title links from DuckDuckGo
                elements = self.driver.find_elements(
                    By.CSS_SELECTOR,
                    'a[data-testid="result-title-a"], a[data-testid="result-extras-url-link"], article a[href]'
                )
                for el in elements:
                    raw_href = el.get_attribute("href")
                    if raw_href and "duckduckgo.com" not in raw_href and raw_href.startswith("http"):
                        candidate_urls.append(raw_href)

                website = extract_website_from_urls(brand, candidate_urls)
                if website:
                    return website

            except Exception as e:
                logger.warning(f"Error during DuckDuckGo website search for '{query}': {e}")

        return None

    def search_employees(self, company_name: str) -> Optional[str]:
        """
        Searches DuckDuckGo for the company's number of employees.
        Query format: [company name] no of employees
        """
        clean_comp = clean_company_name(company_name)
        brand = extract_brand_name(clean_comp)

        # Exact requested format: [company name] no of employees
        queries = [f"{clean_comp} no of employees"]
        if brand.lower() != clean_comp.lower():
            queries.append(f"{brand} no of employees")

        for query in queries:
            logger.info(f"   [DuckDuckGo] Typing Query for Employees: '{query}'")
            try:
                self.navigate_search(query)

                articles = self.driver.find_elements(
                    By.CSS_SELECTOR,
                    'article, [data-testid="result"]'
                )
                for art in articles:
                    txt = art.text.strip()
                    employees = extract_employees_from_text(txt)
                    if employees:
                        return employees

            except Exception as e:
                logger.warning(f"Error during DuckDuckGo employee search for '{query}': {e}")

        return None


class CompanyEnricherApp:
    """Manages Excel file enrichment with DuckDuckGo automation."""

    def __init__(
        self,
        file_path: str,
        sheet_name: Optional[str] = None,
        headless: bool = False,
        save_interval: int = 1,
        delay_seconds: float = 1.0,
        limit: Optional[int] = None,
        overwrite: bool = False
    ):
        self.file_path = os.path.abspath(file_path)
        self.sheet_name = sheet_name
        self.headless = headless
        self.save_interval = max(1, save_interval)
        self.delay_seconds = delay_seconds
        self.limit = limit
        self.overwrite = overwrite

        self.wb: Optional[openpyxl.Workbook] = None
        self.sheet: Optional[Worksheet] = None
        self.col_map: Dict[str, int] = {}
        self.is_dirty: bool = False
        self.automator = DuckDuckGoAutomator(headless=self.headless)

        self.stats = {
            "total_rows": 0,
            "empty_rows": 0,
            "skipped_already_done": 0,
            "processed": 0,
            "website_found": 0,
            "employees_found": 0,
        }

    def prepare_excel(self) -> None:
        """Loads spreadsheet, maps columns, and adds new columns if needed."""
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"Excel file not found at: {self.file_path}")

        logger.info(f"Opening Excel file: {self.file_path}")
        self.wb = openpyxl.load_workbook(self.file_path)

        if self.sheet_name and self.sheet_name in self.wb.sheetnames:
            self.sheet = self.wb[self.sheet_name]
        else:
            self.sheet = self.wb.active

        logger.info(f"Using Sheet: '{self.sheet.title}' (Total Rows: {self.sheet.max_row})")

        headers: Dict[int, str] = {}
        for c in range(1, self.sheet.max_column + 1):
            val = self.sheet.cell(1, c).value
            if val is not None:
                headers[c] = str(val).strip().lower()

        def match_col(candidates: List[str]) -> Optional[int]:
            for c, text in headers.items():
                for cand in candidates:
                    if cand == text or text.startswith(cand):
                        return c
            return None

        # 1. Company Name Column
        company_col = match_col(COMPANY_HEADER_CANDIDATES) or 2
        self.col_map["company"] = company_col

        # 2. Target Columns: Website and No of Employees
        targets = [
            ("website", WEBSITE_HEADER_CANDIDATES, "Company Website", 35),
            ("employees", EMPLOYEES_HEADER_CANDIDATES, "No of Employees", 25),
        ]

        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_align = Alignment(horizontal="center", vertical="center")

        curr_max = self.sheet.max_column
        for key, candidates, title, col_width in targets:
            c = match_col(candidates)
            if c:
                self.col_map[key] = c
                logger.info(f"Found existing column '{title}' at Column {openpyxl.utils.get_column_letter(c)}")
            else:
                curr_max += 1
                new_cell = self.sheet.cell(1, curr_max, value=title)
                new_cell.font = header_font
                new_cell.fill = header_fill
                new_cell.alignment = header_align
                col_letter = openpyxl.utils.get_column_letter(curr_max)
                self.sheet.column_dimensions[col_letter].width = col_width
                self.col_map[key] = curr_max
                headers[curr_max] = title.lower()
                self.is_dirty = True
                logger.info(f"Appended new column '{title}' at Column {col_letter}")

        logger.info(
            "Columns mapped: Company=%s, Website=%s, Employees=%s",
            openpyxl.utils.get_column_letter(self.col_map["company"]),
            openpyxl.utils.get_column_letter(self.col_map["website"]),
            openpyxl.utils.get_column_letter(self.col_map["employees"]),
        )

        if self.is_dirty:
            self.save()

    def process_row(self, row_idx: int) -> bool:
        """Processes a single company row via DuckDuckGo."""
        raw_company = self.sheet.cell(row=row_idx, column=self.col_map["company"]).value
        cleaned_company = clean_company_name(raw_company)

        if not cleaned_company:
            self.stats["empty_rows"] += 1
            return False

        c_website = self.sheet.cell(row=row_idx, column=self.col_map["website"]).value
        c_employees = self.sheet.cell(row=row_idx, column=self.col_map["employees"]).value

        has_website = bool(c_website and str(c_website).strip())
        has_employees = bool(c_employees and str(c_employees).strip())

        # Safe Resume: Skip rows already populated with both fields unless overwrite is enabled
        if not self.overwrite and has_website and has_employees:
            self.stats["skipped_already_done"] += 1
            return False

        total_rows = self.sheet.max_row
        print("-" * 75)
        logger.info(f"[Row {row_idx}/{total_rows}] Processing Company: '{cleaned_company}'")
        self.stats["processed"] += 1

        fields_changed = False

        # 1. Search Company Website in format: [company name] website url
        if not has_website or self.overwrite:
            website = self.automator.search_website(cleaned_company)
            if website:
                self.sheet.cell(row=row_idx, column=self.col_map["website"], value=website)
                logger.info(f"   [FOUND WEBSITE]   -> {website}")
                fields_changed = True
                self.stats["website_found"] += 1
            else:
                logger.info("   [WEBSITE]         -> [Not Found]")
        else:
            logger.info(f"   -> Website already present: {c_website}")

        time.sleep(0.4)

        # 2. Search Number of Employees in format: [company name] no of employees
        if not has_employees or self.overwrite:
            employees = self.automator.search_employees(cleaned_company)
            if employees:
                self.sheet.cell(row=row_idx, column=self.col_map["employees"], value=employees)
                logger.info(f"   [FOUND EMPLOYEES] -> {employees}")
                fields_changed = True
                self.stats["employees_found"] += 1
            else:
                logger.info("   [EMPLOYEES]       -> [Not Found]")
        else:
            logger.info(f"   -> Employees already present: {c_employees}")

        if fields_changed:
            self.is_dirty = True

        return True

    def save(self) -> None:
        """Saves Excel file incrementally, with fallback if open in Excel."""
        if self.wb:
            try:
                self.wb.save(self.file_path)
                self.is_dirty = False
                logger.info(f"[EXCEL SAVED] Progress saved to '{os.path.basename(self.file_path)}'")
            except PermissionError:
                base, ext = os.path.splitext(self.file_path)
                fallback_path = f"{base}_updated{ext}"
                logger.warning(
                    f"\n[FILE LOCKED] '{os.path.basename(self.file_path)}' is open in Microsoft Excel!\n"
                    f"   -> Saving progress to '{os.path.basename(fallback_path)}' so no data is lost.\n"
                    f"   -> (Close the file in Excel to allow saving directly to the original file.)"
                )
                try:
                    self.wb.save(fallback_path)
                    logger.info(f"[EXCEL SAVED] Progress secured in '{os.path.basename(fallback_path)}'")
                except Exception as e2:
                    logger.error(f"Fallback save error: {e2}")
            except Exception as e:
                logger.error(f"Error saving Excel file: {e}")

    def run(self) -> None:
        """Runs the enrichment loop through all Excel rows."""
        self.prepare_excel()
        self.automator.start()

        max_row = self.sheet.max_row
        processed_count = 0
        since_last_save = 0

        logger.info(f"Starting DuckDuckGo search automation for {max_row - 1} rows...")

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
                    logger.warning("\nExecution paused by user (Ctrl+C). Saving...")
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
        print("   DUCKDUCKGO WEBSITE & EMPLOYEES ENRICHMENT SUMMARY")
        print("=" * 65)
        print(f"  Total Rows Inspected:       {self.stats['total_rows']}")
        print(f"  Empty Rows Skipped:         {self.stats['empty_rows']}")
        print(f"  Already Enriched (Skipped): {self.stats['skipped_already_done']}")
        print(f"  Companies Processed:        {self.stats['processed']}")
        print(f"  Websites Found:             {self.stats['website_found']}")
        print(f"  Employee Counts Found:      {self.stats['employees_found']}")
        print("=" * 65 + "\n")


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="DuckDuckGo Search Automation - Company Website & Employees Count Enrichment (Strictly No Google)"
    )
    parser.add_argument(
        "--file", "-f",
        default="companies_sample.xlsx",
        help="Path to Excel spreadsheet (default: companies_sample.xlsx)"
    )
    parser.add_argument(
        "--sheet", "-s",
        default=None,
        help="Sheet name to process (default: active sheet)"
    )
    parser.add_argument(
        "--limit", "-l",
        type=int,
        default=None,
        help="Process only first N companies (for testing)"
    )
    parser.add_argument(
        "--delay", "-d",
        type=float,
        default=1.0,
        help="Delay in seconds between searches (default: 1.0)"
    )
    parser.add_argument(
        "--save-interval",
        type=int,
        default=1,
        help="Save Excel workbook every N rows (default: 1)"
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run browser silently in background without desktop window"
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Re-search and overwrite existing fields in Excel"
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_arguments()
    excel_path = os.path.abspath(args.file)

    if not os.path.exists(excel_path):
        desktop_candidate = os.path.join(os.path.expanduser("~"), "Desktop", os.path.basename(excel_path))
        if os.path.exists(desktop_candidate):
            excel_path = desktop_candidate
        else:
            print(f"Error: File not found at '{excel_path}'")
            sys.exit(1)

    enricher = CompanyEnricherApp(
        file_path=excel_path,
        sheet_name=args.sheet,
        headless=args.headless,
        save_interval=args.save_interval,
        delay_seconds=args.delay,
        limit=args.limit,
        overwrite=args.overwrite
    )

    enricher.run()
