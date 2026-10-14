"""
================================================================================
LIVE DESKTOP CHROME AUTOMATION - COMPANY DATA ENRICHMENT
(CEO, HR HEAD, HEADQUARTERS & ADMIN / SUPPORT PERSON)
================================================================================

DESCRIPTION:
    Automates Google Chrome directly on your desktop screen so you can watch it
    work live! For each company in your Excel file:
      1. Reads the company name from Excel.
      2. Types the search into a real, visible Google Chrome browser window on your desktop.
      3. Searches live for the real CEO name & LinkedIn profile URL.
      4. Searches live for the real HR / Head of People name & LinkedIn profile URL.
      5. Searches live for the company's real Headquarters location (City, State, Country).
      6. Searches live for the company's Support Person or Admin name & LinkedIn profile URL.
      7. Writes all 7 values back into the Excel spreadsheet in the respective columns.
      8. Saves progress incrementally every N rows (default: 5).
      9. Safe Resume: Skips rows/fields that are already populated and resumes seamlessly.

HOW TO RUN:
    Simply open PowerShell or Command Prompt and run:
        python enrich_companies.py

    This automatically launches a visible Google Chrome window maximized on your desktop
    and enriches:
        C:/Users/Elakkiya Tamilselvan/Desktop/chennai companys.xlsx

OPTIONS:
    Test on just 5 companies:
        python enrich_companies.py --limit 5

    Specify a different Excel file:
        python enrich_companies.py --file "C:/path/to/my_file.xlsx"

    Run silently in background (headless):
        python enrich_companies.py --headless

DEPENDENCIES:
    pip install openpyxl selenium webdriver-manager python-dotenv
================================================================================
"""

import os
import re
import sys
import time
import random
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
logger = logging.getLogger("desktop_enricher")

# Silence noisy third-party loggers
for mod in ("selenium", "urllib3", "webdriver_manager"):
    logging.getLogger(mod).setLevel(logging.WARNING)

LINKEDIN_PROFILE_REGEX = re.compile(
    r'https?://(?:[a-z]{2,3}\.)?linkedin\.com/in/([a-zA-Z0-9_\-%]+)/?',
    re.IGNORECASE
)

COMPANY_HEADER_CANDIDATES = [
    "name", "company", "company name", "company_name", "organization",
    "firm", "account name", "companies"
]
CORE_BUSINESS_HEADER_CANDIDATES = [
    "core business", "core_business", "nature of business", "business overview",
    "business description", "industry / business", "what they do"
]
CEO_NAME_HEADER_CANDIDATES = [
    "ceo name", "ceo", "chief executive officer", "ceo_name", "current ceo", "md & ceo"
]
CEO_LINKEDIN_HEADER_CANDIDATES = [
    "ceo linkedin", "ceo linkedin url", "ceo linkedin profile",
    "ceo_linkedin", "ceo_linkedin_url", "ceo profile"
]
HR_NAME_HEADER_CANDIDATES = [
    "hr name", "hr head name", "hr head", "head of hr", "chief people officer",
    "chro", "head of people", "hr director", "hr_name", "hr_head_name", "vp hr"
]
HR_LINKEDIN_HEADER_CANDIDATES = [
    "hr linkedin", "hr linkedin url", "hr head linkedin", "head of hr linkedin",
    "hr_linkedin", "hr_linkedin_url", "hr profile", "head of people linkedin"
]
HEADQUARTERS_HEADER_CANDIDATES = [
    "headquarters", "hq", "headquarter", "location", "headquarters location",
    "head office", "company location", "head office location"
]
ADMIN_NAME_HEADER_CANDIDATES = [
    "admin / support person name", "support person or admin name",
    "support person or admin", "support / admin person name", "support / admin name",
    "admin name", "support person name", "support person", "admin person",
    "office admin", "support / admin", "admin / support", "administrator"
]
ADMIN_LINKEDIN_HEADER_CANDIDATES = [
    "admin / support person linkedin", "admin / support person linkedin url",
    "support person or admin linkedin", "support person or admin linkedin url",
    "support / admin person linkedin", "support / admin person linkedin url",
    "support / admin linkedin", "admin / support linkedin", "support person linkedin",
    "admin linkedin", "admin linkedin url", "support linkedin", "support linkedin url"
]


def clean_company_name(name: str) -> str:
    """Cleans whitespace, quotes, and punctuation from company names."""
    if not name:
        return ""
    cleaned = str(name).strip().strip('\'"“”')
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned


def clean_linkedin_url(url: str) -> Optional[str]:
    """Returns a canonical https://www.linkedin.com/in/<username>/ URL."""
    if not url:
        return None
    match = LINKEDIN_PROFILE_REGEX.search(url)
    if match:
        username = match.group(1).rstrip('/')
        if username.lower() in ("search", "company", "dir", "feed", "posts", "pulse"):
            return None
        return f"https://www.linkedin.com/in/{username}/"
    return None


def extract_person_name(title: str, company_name: str) -> Optional[str]:
    """Extracts and sanitizes a real person's name from a search title."""
    if not title:
        return None

    title = title.split('\n')[0]
    title = re.split(r'[\s|–—\-]+LinkedIn.*$', title, flags=re.IGNORECASE)[0].strip()

    parts = re.split(r'\s+[-–—:|@]\s+|\s+at\s+', title)
    if not parts:
        return None

    cand = parts[0].strip()
    if '...' in cand:
        cand = cand.split('...')[0].strip()

    job_prefix_regex = (
        r'^(?:CEO|Founder|Managing Director|MD|Director|President|Head of HR|'
        r'HR Head|CHRO|Chief Executive Officer|VP|Vice President|Admin|Administrator|'
        r'Support|Office Manager|Executive Assistant)\s+'
    )
    cand = re.sub(job_prefix_regex, '', cand, flags=re.IGNORECASE).strip()
    cand = re.sub(r'^(?:Dr\.|Mr\.|Mrs\.|Ms\.|Prof\.)\s+', '', cand, flags=re.IGNORECASE).strip()
    cand = re.sub(r'\s*\((?:Admin|Administrator|Support|IT|HR|Operations)\)', '', cand, flags=re.IGNORECASE).strip()
    cand = re.sub(r'^[^\w]+|[^\w.]+$', '', cand).strip()

    words = cand.split()
    if not (1 <= len(words) <= 4 and 2 <= len(cand) <= 35):
        return None

    blacklist = {
        'overview', 'profile', 'profiles', 'jobs', 'directory', 'salaries',
        'top', 'careers', 'working', 'team', 'pvt', 'ltd', 'limited', 'private',
        'technologies', 'solutions', 'systems', 'services', 'network', 'india',
        'chennai', 'bangalore', 'mumbai', 'delhi', 'software', 'infotech'
    }

    if any(w.lower() in blacklist for w in words):
        return None

    comp_clean = clean_company_name(company_name).lower()
    if comp_clean in cand.lower() or cand.lower() in comp_clean:
        return None

    if not re.search(r'[A-Za-z]', cand):
        return None

    return cand


def extract_headquarters_from_text(text: str) -> Optional[str]:
    """
    Parses headquarters location (City, State, Country) from text blocks.
    """
    if not text:
        return None

    # Pattern 1: 'headquarters: <City, State, Country>' or 'headquartered in ...'
    mA = re.search(
        r'(?:headquarters|headquartered\s+in|head\s+office|registered\s+office)\s*[:\-–]?\s*([A-Za-z\s]+,\s*(?:[A-Za-z\s]+,\s*)?(?:India|USA|UK|Singapore|[A-Z]{2}))',
        text, re.I
    )
    if mA:
        cand = mA.group(1).strip()
        cand = re.sub(r'^(?:located\s+in|headquartered\s+in|based\s+in|are\s+in|is\s+in|situated\s+in)\s+', '', cand, flags=re.I).strip()
        if 3 <= len(cand) <= 50:
            return cand

    # Pattern 2: City, State, Country
    mB = re.search(
        r'\b([A-Za-z\s]{3,20},\s*(?:Tamil Nadu|Maharashtra|Karnataka|Telangana|Gujarat|Delhi|Uttar Pradesh|Haryana|West Bengal|Kerala|Andhra Pradesh))(?:,\s*\d{5,6})?(?:,\s*(India|USA|UK))?\b',
        text, re.I
    )
    if mB:
        city_state = mB.group(1).strip()
        city_state = re.sub(r'^(?:located\s+in|headquartered\s+in|based\s+in)\s+', '', city_state, flags=re.I).strip()
        country = mB.group(2).strip() if mB.group(2) else ('India' if 'india' in text.lower() else '')
        if country and country.lower() not in city_state.lower():
            return f'{city_state}, {country}'
        return city_state

    # Pattern 3: Major global / Indian business cities
    major_cities = [
        'Chennai', 'Bangalore', 'Bengaluru', 'Mumbai', 'Hyderabad', 'Pune',
        'Noida', 'Gurgaon', 'Gurugram', 'Delhi', 'New Delhi', 'Coimbatore',
        'Kolkata', 'Ahmedabad', 'Kochi', 'Trivandrum', 'Madurai', 'Trichy',
        'San Francisco', 'New York', 'London', 'Singapore'
    ]
    city_pattern = r'\b(' + '|'.join(major_cities) + r')\b(?:,\s*(India|USA|UK|Singapore))?'
    mC = re.search(city_pattern, text, re.I)
    if mC:
        city = mC.group(1).strip()
        country = mC.group(2).strip() if mC.group(2) else ('India' if 'india' in text.lower() else '')
        return f'{city}, {country}' if country else city

    return None


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
        """Launches the Chrome browser window."""
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
            time.sleep(0.015)

        box.send_keys(Keys.RETURN)
        time.sleep(2.5)

    def search_executive(self, company_name: str, role_type: str = "ceo") -> Tuple[Optional[str], Optional[str]]:
        """
        Visibly types query into Chrome, presses Enter, and parses the top LinkedIn result.
        role_type can be 'ceo', 'hr', or 'admin'.
        Returns: (Name, LinkedIn_URL)
        """
        if not self.driver:
            self.start()

        clean_comp = clean_company_name(company_name)
        suffix_free = re.sub(
            r'\b(?:pvt\.?\s*ltd\.?|private\s+limited|p\s+ltd|ltd\.?|limited|inc\.?|llc|corp\.?|co\.?)\b',
            '', clean_comp, flags=re.IGNORECASE
        ).strip()
        search_brand = suffix_free if len(suffix_free) >= 3 else clean_comp

        if role_type == "ceo":
            queries = [f'"{search_brand}" CEO LinkedIn']
        elif role_type == "hr":
            queries = [
                f'"{search_brand}" HR Head LinkedIn',
                f'{search_brand} HR LinkedIn'
            ]
        elif role_type == "admin":
            queries = [
                f'{search_brand} Admin LinkedIn',
                f'{search_brand} Support LinkedIn',
                f'{search_brand} Office Admin LinkedIn'
            ]
        else:
            queries = [f'{search_brand} {role_type} LinkedIn']

        for query in queries:
            logger.info(f"   [Chrome Screen] Typing query: {query}")
            try:
                self.type_and_search(query)

                articles = self.driver.find_elements(
                    By.CSS_SELECTOR,
                    'article[data-testid="result"], article, div[data-layout="organic"]'
                )

                best_candidate = None

                for item in articles:
                    item_text = item.text.strip()
                    if not item_text:
                        continue
                    if 'linkedin.com/in/' not in item_text and 'LinkedIn' not in item_text:
                        continue

                    links = item.find_elements(By.TAG_NAME, 'a')
                    for a in links:
                        href = a.get_attribute('href') or ''
                        clean_url = clean_linkedin_url(href)
                        if clean_url:
                            lines = [line.strip() for line in item_text.split('\n') if line.strip()]
                            name = None

                            for line in lines:
                                if any(sep in line for sep in (' - ', ' – ', ' | ', ' at ')) and 'linkedin.com' not in line.lower():
                                    cand_name = extract_person_name(line, clean_comp)
                                    if cand_name:
                                        name = cand_name
                                        break

                            if not name and len(lines) > 2:
                                for line in lines[2:]:
                                    cand_name = extract_person_name(line, clean_comp)
                                    if cand_name:
                                        name = cand_name
                                        break

                            if not name:
                                continue

                            # Brand relevance match check
                            is_brand_match = (search_brand.lower() in item_text.lower())
                            if is_brand_match:
                                return name, clean_url

                            if not best_candidate:
                                best_candidate = (name, clean_url)

                if best_candidate:
                    return best_candidate

            except Exception as e:
                logger.warning(f"Browser search error for {query}: {e}")

        return None, None

    def search_headquarters(self, company_name: str) -> Optional[str]:
        """
        Visibly types headquarters search query into Chrome, presses Enter,
        and extracts the city/location from top search cards.
        """
        if not self.driver:
            self.start()

        clean_comp = clean_company_name(company_name)
        suffix_free = re.sub(
            r'\b(?:pvt\.?\s*ltd\.?|private\s+limited|p\s+ltd|ltd\.?|limited|inc\.?|llc|corp\.?|co\.?)\b',
            '', clean_comp, flags=re.IGNORECASE
        ).strip()
        search_brand = suffix_free if len(suffix_free) >= 3 else clean_comp

        queries = [
            f'"{search_brand}" headquarters location',
            f'"{clean_comp}" office address',
            f'"{search_brand}" headquarters'
        ]

        nav_words = {'all', 'images', 'maps', 'videos', 'news'}

        for query in queries:
            logger.info(f"   [Chrome Screen] Typing query: {query}")
            try:
                self.type_and_search(query)

                articles = self.driver.find_elements(
                    By.CSS_SELECTOR,
                    'article[data-testid="result"], article, div[data-layout="organic"]'
                )

                for item in articles[:6]:
                    item_text = item.text.strip()
                    if not item_text or item_text.lower() in nav_words:
                        continue
                    hq = extract_headquarters_from_text(item_text)
                    if hq:
                        return hq

            except Exception as e:
                logger.warning(f"Headquarters search error for {query}: {e}")

        return None

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


class ExcelLiveRunner:
    """Manages Excel file enrichment with live Chrome visualization."""

    def __init__(
        self,
        file_path: str,
        sheet_name: Optional[str] = None,
        headless: bool = False,
        save_interval: int = 5,
        delay_seconds: float = 2.0,
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
        self.col_map: Dict[str, int] = {}
        self.is_dirty: bool = False
        self.automator = LiveBrowserAutomator(headless=self.headless)

        self.stats = {
            "total_rows": 0,
            "empty_rows": 0,
            "skipped_already_done": 0,
            "processed": 0,
            "ceo_found": 0,
            "hr_found": 0,
            "hq_found": 0,
            "admin_found": 0,
            "core_biz_found": 0,
        }

    def prepare_excel(self) -> None:
        """Loads spreadsheet and configures headers."""
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

        def match_col(candidates):
            for c, text in headers.items():
                for cand in candidates:
                    if cand == text or text.startswith(cand):
                        return c
            return None

        # Company Column
        company_col = match_col(COMPANY_HEADER_CANDIDATES) or 1
        self.col_map["company"] = company_col

        # Enrichment columns (CEO, HR, Headquarters, Admin & Core Business)
        targets = [
            ("ceo_name", CEO_NAME_HEADER_CANDIDATES, "CEO Name", 24),
            ("ceo_url", CEO_LINKEDIN_HEADER_CANDIDATES, "CEO LinkedIn URL", 45),
            ("hr_name", HR_NAME_HEADER_CANDIDATES, "HR Head Name", 24),
            ("hr_url", HR_LINKEDIN_HEADER_CANDIDATES, "HR Head LinkedIn URL", 45),
            ("headquarters", HEADQUARTERS_HEADER_CANDIDATES, "Headquarters", 30),
            ("admin_name", ADMIN_NAME_HEADER_CANDIDATES, "Admin / Support Person Name", 28),
            ("admin_url", ADMIN_LINKEDIN_HEADER_CANDIDATES, "Admin / Support Person LinkedIn", 45),
            ("core_business", CORE_BUSINESS_HEADER_CANDIDATES, "Core Business", 38),
        ]

        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_align = Alignment(horizontal="center", vertical="center")

        curr_max = self.sheet.max_column
        for key, candidates, title, col_width in targets:
            c = match_col(candidates)
            if c:
                self.col_map[key] = c
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
                logger.info(f"Appended column '{title}' at Column {col_letter}")

        logger.info(
            "Columns mapped: Company=%s, CEO Name=%s, CEO LinkedIn=%s, HR Name=%s, HR LinkedIn=%s, HQ=%s, Admin Name=%s, Admin LinkedIn=%s, Core Business=%s",
            openpyxl.utils.get_column_letter(self.col_map["company"]),
            openpyxl.utils.get_column_letter(self.col_map["ceo_name"]),
            openpyxl.utils.get_column_letter(self.col_map["ceo_url"]),
            openpyxl.utils.get_column_letter(self.col_map["hr_name"]),
            openpyxl.utils.get_column_letter(self.col_map["hr_url"]),
            openpyxl.utils.get_column_letter(self.col_map["headquarters"]),
            openpyxl.utils.get_column_letter(self.col_map["admin_name"]),
            openpyxl.utils.get_column_letter(self.col_map["admin_url"]),
            openpyxl.utils.get_column_letter(self.col_map["core_business"]),
        )

        if self.is_dirty:
            self.save()

    def process_row(self, row_idx: int) -> bool:
        """Processes a single company row via visible Chrome."""
        raw_company = self.sheet.cell(row=row_idx, column=self.col_map["company"]).value
        cleaned_company = clean_company_name(raw_company)

        if not cleaned_company:
            self.stats["empty_rows"] += 1
            return False

        c_ceo_name = self.sheet.cell(row=row_idx, column=self.col_map["ceo_name"]).value
        c_ceo_url = self.sheet.cell(row=row_idx, column=self.col_map["ceo_url"]).value
        c_hr_name = self.sheet.cell(row=row_idx, column=self.col_map["hr_name"]).value
        c_hr_url = self.sheet.cell(row=row_idx, column=self.col_map["hr_url"]).value
        c_hq = self.sheet.cell(row=row_idx, column=self.col_map["headquarters"]).value
        c_admin_name = self.sheet.cell(row=row_idx, column=self.col_map["admin_name"]).value
        c_admin_url = self.sheet.cell(row=row_idx, column=self.col_map["admin_url"]).value
        c_core_biz = self.sheet.cell(row=row_idx, column=self.col_map["core_business"]).value

        has_ceo_name = bool(c_ceo_name and str(c_ceo_name).strip())
        has_ceo_url = bool(c_ceo_url and str(c_ceo_url).strip())
        has_hr_name = bool(c_hr_name and str(c_hr_name).strip())
        has_hr_url = bool(c_hr_url and str(c_hr_url).strip())
        has_hq = bool(c_hq and str(c_hq).strip())
        has_admin_name = bool(c_admin_name and str(c_admin_name).strip())
        has_admin_url = bool(c_admin_url and str(c_admin_url).strip())
        has_core_biz = bool(c_core_biz and str(c_core_biz).strip())

        # Skip rows already completely enriched with all 8 fields
        if has_ceo_name and has_ceo_url and has_hr_name and has_hr_url and has_hq and has_admin_name and has_admin_url and has_core_biz:
            self.stats["skipped_already_done"] += 1
            return False

        total_rows = self.sheet.max_row
        print("-" * 65)
        logger.info(f"[Row {row_idx}/{total_rows}] Processing Company: '{cleaned_company}'")
        self.stats["processed"] += 1

        fields_changed = False

        # 1. Search CEO
        if not (has_ceo_name and has_ceo_url):
            ceo_name, ceo_url = self.automator.search_executive(cleaned_company, role_type="ceo")
            if not has_ceo_name and ceo_name:
                self.sheet.cell(row=row_idx, column=self.col_map["ceo_name"], value=ceo_name)
                logger.info(f"   [FOUND CEO NAME]        -> {ceo_name}")
                fields_changed = True
                self.stats["ceo_found"] += 1
            elif not has_ceo_name:
                logger.info("   [CEO NAME]              -> [Not Found]")

            if not has_ceo_url and ceo_url:
                self.sheet.cell(row=row_idx, column=self.col_map["ceo_url"], value=ceo_url)
                logger.info(f"   [FOUND CEO LINKEDIN]    -> {ceo_url}")
                fields_changed = True
            elif not has_ceo_url:
                logger.info("   [CEO LINKEDIN]          -> [Not Found]")
        else:
            logger.info("   -> CEO already present in Excel.")

        time.sleep(0.5)

        # 2. Search HR
        if not (has_hr_name and has_hr_url):
            hr_name, hr_url = self.automator.search_executive(cleaned_company, role_type="hr")
            if not has_hr_name and hr_name:
                self.sheet.cell(row=row_idx, column=self.col_map["hr_name"], value=hr_name)
                logger.info(f"   [FOUND HR NAME]         -> {hr_name}")
                fields_changed = True
                self.stats["hr_found"] += 1
            elif not has_hr_name:
                logger.info("   [HR NAME]               -> [Not Found]")

            if not has_hr_url and hr_url:
                self.sheet.cell(row=row_idx, column=self.col_map["hr_url"], value=hr_url)
                logger.info(f"   [FOUND HR LINKEDIN]     -> {hr_url}")
                fields_changed = True
            elif not has_hr_url:
                logger.info("   [HR LINKEDIN]           -> [Not Found]")
        else:
            logger.info("   -> HR already present in Excel.")

        time.sleep(0.5)

        # 3. Search Headquarters
        if not has_hq:
            hq_loc = self.automator.search_headquarters(cleaned_company)
            if hq_loc:
                self.sheet.cell(row=row_idx, column=self.col_map["headquarters"], value=hq_loc)
                logger.info(f"   [FOUND HEADQUARTERS]    -> {hq_loc}")
                fields_changed = True
                self.stats["hq_found"] += 1
            else:
                logger.info("   [HEADQUARTERS]          -> [Not Found]")
        else:
            logger.info("   -> Headquarters already present in Excel.")

        time.sleep(0.5)

        # 4. Search Support Person / Admin
        if not (has_admin_name and has_admin_url):
            admin_name, admin_url = self.automator.search_executive(cleaned_company, role_type="admin")
            if not has_admin_name and admin_name:
                self.sheet.cell(row=row_idx, column=self.col_map["admin_name"], value=admin_name)
                logger.info(f"   [FOUND ADMIN/SUPPORT]   -> {admin_name}")
                fields_changed = True
                self.stats["admin_found"] += 1
            elif not has_admin_name:
                logger.info("   [ADMIN/SUPPORT NAME]    -> [Not Found]")

            if not has_admin_url and admin_url:
                self.sheet.cell(row=row_idx, column=self.col_map["admin_url"], value=admin_url)
                logger.info(f"   [FOUND ADMIN LINKEDIN]  -> {admin_url}")
                fields_changed = True
            elif not has_admin_url:
                logger.info("   [ADMIN LINKEDIN]        -> [Not Found]")
        else:
            logger.info("   -> Admin / Support Person already present in Excel.")

        time.sleep(0.5)

        # 5. Search Core Business
        if not has_core_biz:
            core_biz = self.automator.search_core_business(cleaned_company)
            if core_biz:
                self.sheet.cell(row=row_idx, column=self.col_map["core_business"], value=core_biz)
                logger.info(f"   [FOUND CORE BUSINESS]   -> {core_biz}")
                fields_changed = True
                self.stats["core_biz_found"] += 1
            else:
                logger.info("   [CORE BUSINESS]         -> [Not Found]")
        else:
            val_str = str(c_core_biz).strip()
            if len(val_str.split()) > 5 or len(val_str) > 40:
                short_biz = extract_short_core_business(val_str, company_name=cleaned_company)
                if short_biz and short_biz != val_str:
                    self.sheet.cell(row=row_idx, column=self.col_map["core_business"], value=short_biz)
                    logger.info(f"   [SHORTENED CORE BUSINESS] -> {short_biz}")
                    fields_changed = True
            else:
                logger.info("   -> Core Business already present in Excel.")

        if fields_changed:
            self.is_dirty = True

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

        logger.info(f"Starting live screen automation for {max_row - 1} rows...")

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
        print("                 LIVE ENRICHMENT SUMMARY")
        print("=" * 65)
        print(f"  Total Rows Inspected:       {self.stats['total_rows']}")
        print(f"  Empty Rows Skipped:         {self.stats['empty_rows']}")
        print(f"  Already Enriched (Skipped): {self.stats['skipped_already_done']}")
        print(f"  Companies Processed Live:   {self.stats['processed']}")
        print(f"  -> CEO Data Found:          {self.stats['ceo_found']}")
        print(f"  -> HR Data Found:           {self.stats['hr_found']}")
        print(f"  -> Headquarters Found:      {self.stats['hq_found']}")
        print(f"  -> Admin/Support Found:     {self.stats['admin_found']}")
        print(f"  -> Core Business Found:     {self.stats['core_biz_found']}")
        print("=" * 65)
        print(f"  Output Excel: {self.file_path}")
        print("=" * 65 + "\n")


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
        description="Live desktop Google Chrome automation for Excel company enrichment (CEO, HR, Headquarters & Support/Admin)."
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
        help="Save to Excel every N rows."
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=2.0,
        help="Delay in seconds between companies."
    )
    parser.add_argument(
        "-l", "--limit",
        type=int,
        default=None,
        help="Max companies to process."
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run silently in the background without opening a browser window."
    )

    args = parser.parse_args()

    print("\n" + "=" * 65)
    print("  LIVE DESKTOP CHROME AUTOMATION - COMPANY ENRICHMENT")
    print("  (CEO, HR, Headquarters & Support Person / Admin)")
    print("=" * 65)

    runner = ExcelLiveRunner(
        file_path=args.file_path,
        sheet_name=args.sheet_name,
        headless=args.headless,
        save_interval=args.save_interval,
        delay_seconds=args.delay,
        limit=args.limit
    )

    runner.run()


if __name__ == "__main__":
    main()
