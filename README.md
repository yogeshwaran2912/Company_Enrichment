# Live Desktop Google Chrome Automation - Company Data Enrichment

This automation drives **Google Chrome directly on your desktop screen** so you can watch it live as it takes each company name from Excel, opens Chrome, searches for its details (CEO, HR Head, Headquarters, Support/Admin, and **Core Business**), and writes the verified data into your Excel spreadsheet.

---

## Columns Enriched in Your Excel Spreadsheet

1. **Company Name** (existing)
2. **CEO Name**
3. **CEO LinkedIn URL**
4. **HR Head Name**
5. **HR Head LinkedIn URL**
6. **Headquarters** *(City, State, Country)*
7. **Admin / Support Person Name**
8. **Admin / Support Person LinkedIn**
9. **Core Business** *(Short domain tags e.g. 'IT Infrastructure')*
10. **Company Website** *(Official URL e.g. 'https://www.usam.in')* ⭐ **NEW**
11. **Company Phone Number** *(Contact number e.g. '044-4343 7171')* ⭐ **NEW**
12. **No of Employees** *(Company size e.g. '311 employees')* ⭐ **NEW**

---

## What You See Happening Live on Your Screen

1. **Google Chrome Opens:** A real, visible Chrome browser launches and maximizes on your desktop.
2. **Company Name Copied:** Reads each company name from Excel row by row.
3. **Live Searches in Chrome:** Visibly searches in Chrome for:
   - Official Company Website
   - Company Contact Phone Number
   - Number of Employees (headcount / company size)
4. **Excel Row Update:** Writes the verified details directly into the 3 columns:
   - `Company Website`
   - `Company Phone Number`
   - `No of Employees`
5. **Incremental Saving:** Automatically saves progress every 5 rows.
6. **Safe Resume:** If you stop and restart, it skips already completed rows and resumes seamlessly.

---

## Quick Tip Before Running

> **IMPORTANT:** Make sure you **close `companies_sample.xlsx` in Microsoft Excel** before running the script so Excel doesn't lock the file. (If it is open, the script will safely save progress to `companies_sample_updated.xlsx` as a backup so no data is lost).

---

## How to Run

Open PowerShell or Command Prompt:

```powershell
cd "C:\Users\ADV PRO\Desktop\company_enrichment"
```

### 1. Enrich Website & Number of Employees (DuckDuckGo Automation) ⭐

Uses **DuckDuckGo** (strictly NO Google) in the exact requested formats:
- **Company Website:** Searches DuckDuckGo for `[company name] website url`, identifies the verified official website URL, and pastes it into the `Company Website` column.
- **No of Employees:** Searches DuckDuckGo for `[company name] no of employees`, extracts the verified employee count / headcount, and pastes it into the `No of Employees` column.

```powershell
python enrich_website_and_employees.py
```

#### Test on 5 Companies First:
```powershell
python enrich_website_and_employees.py --limit 5
```

#### Run in Background (Headless Mode):
```powershell
python enrich_website_and_employees.py --headless
```

---

### 2. Enrich Website, Phone Number & Number of Employees (All 3 Fields)

To enrich all 3 fields (`Company Website`, `Company Phone Number`, and `No of Employees`):

```powershell
python enrich_company_details.py
```

#### Test on 5 Companies First:
```powershell
python enrich_company_details.py --limit 5
```

---

### 2. Enrich Core Business Only (New Column)

To add the **Core Business** column and enrich each company one by one:

```powershell
python enrich_core_business.py
```

---

### 3. Full Executive Enrichment (CEO, HR, HQ, Admin)

To enrich all leadership columns:

```powershell
python enrich_companies.py
```


#### Test on 5 Companies First:
```powershell
python enrich_companies.py --limit 5
```

---

### 3. Instantly Shorten Existing Long Descriptions in Excel (No Browser Needed)

If your spreadsheet already has long paragraphs in the "Core Business" column and you want to instantly convert them into concise tags:

```powershell
python shorten_existing_core_business.py
```
This scans your Excel file in seconds and replaces long sentence paragraphs with short tags like `IT Infrastructure` and `Industrial Automation`.
