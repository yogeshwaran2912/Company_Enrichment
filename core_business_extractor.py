"""
================================================================================
CORE BUSINESS EXTRACTOR & CATEGORIZATION MODULE
================================================================================
Transforms raw web search snippets and descriptions into short, clean,
professional core business labels (e.g., 'IT Infrastructure',
'Industrial Automation', 'Advertising Technology (AdTech)').
================================================================================
"""

import re
from typing import Optional, List, Tuple

# Predefined Domain Taxonomy: (Regex Pattern, Short Standard Industry Label)
TAXONOMY_MAP: List[Tuple[str, str]] = [
    # Explicit company overrides / high-intent matches
    (r'\bample\s+technologies\b', "IT Infrastructure"),
    (r'\bmicrosoft\b', "Enterprise Software & Cloud Platforms"),
    (r'\bapple\b', "Consumer Electronics & Technology"),
    (r'\bnvidia\b', "AI & Semiconductor Computing"),
    (r'\b(?:stripe|payments?|payment processor|fraud detection)\b', "FinTech & Payment Solutions"),
    (r'\b(?:airbnb|vacation rental|travel market|lodging|booking|hospitality)\b', "Online Travel & Hospitality"),
    (r'\b(?:ad\s*tech|advertising technology|programmatic ads?|dsp|ssp|ad server|ad exchange)\b', "Advertising Technology (AdTech)"),
    
    # IT Infrastructure, Cloud & Cybersecurity
    (r'\b(?:it infrastructure|network infrastructure|cloud infrastructure|server infrastructure|data center|apple premium reseller)\b', "IT Infrastructure"),
    (r'\b(?:cloud computing|cloud services|cloud migration|aws|azure|gcp|devops|cloud platform)\b', "Cloud Computing & DevOps"),
    (r'\b(?:cybersecurity|cyber security|network security|information security|firewall|endpoint security|soc services)\b', "Cybersecurity & Network Security"),
    
    # Telecommunications & Networking
    (r'\b(?:data com(?:munications?)?|telecom(?:munications?)?|networking solutions|voip|broadband|optical fiber|network cabling)\b', "Telecommunications & Networking"),
    (r'\b(?:computer hardware|it hardware|server hardware|peripherals|pc components|computer systems)\b', "Computer Hardware & Networking"),

    # CAD & Engineering Design Software
    (r'\b(?:cad|cam|cae|solidworks|autocad|3d design software|engineering design software)\b', "CAD & Engineering Software"),

    # Software & Applications
    (r'\b(?:application development|custom software|software solutions|software development|app development|web development|mobile app)\b', "Application Development & Software Solutions"),
    (r'\b(?:enterprise software|erp software|crm software|business management software|tally prime|saas platform|erp solutions)\b', "Enterprise Software Solutions"),
    (r'\b(?:artificial intelligence|machine learning|data science|data analytics|generative ai|business intelligence|big data)\b', "AI & Data Analytics"),
    (r'\b(?:it consulting|digital transformation|it services and consulting|technology consulting|software consulting|low-code)\b', "IT Consulting & Services"),
    
    # Automation, Robotics & Machinery
    (r'\b(?:cnc machine|cnc machine tools|machine tools|lathe|milling machine|machining center)\b', "CNC Machine Tools & Automation"),
    (r'\b(?:control systems?|control panels?|electrical control|switchgear|power electronics|electrical panels?|instrumentation)\b', "Control Systems & Automation"),
    (r'\b(?:industrial automation|automation systems|plc|scada|process automation|factory automation|robotics)\b', "Industrial Automation"),
    (r'\b(?:precision engineering|machining|tooling|metal fabrication|sheet metal)\b', "Precision Engineering & Manufacturing"),
    (r'\b(?:electronic components?|semiconductors?|pcb assembly|circuit boards)\b', "Electronic Components Manufacturing"),
    (r'\b(?:automobile|automotive|auto parts|auto components|vehicle manufacturing)\b', "Automotive & Auto Components"),

    # Digital, Marketing & E-Commerce
    (r'\b(?:digital marketing|seo|social media marketing|search engine optimization|content marketing)\b', "Digital Marketing"),
    (r'\b(?:e-commerce|ecommerce|online retail|marketplace)\b', "E-Commerce Solutions"),

    # Business Services & Finance
    (r'\b(?:business process outsourcing|bpo|kpo|customer support services|back office|payroll services|bpm|call center)\b', "Business Process Outsourcing (BPO)"),
    (r'\b(?:banking solutions|financial services|fintech|investment banking|wealth management|insurance services)\b', "Financial Services & FinTech"),
    (r'\b(?:staffing services|recruitment|human resources|talent acquisition|payroll management|manpower)\b', "Human Resources & Staffing"),
    (r'\b(?:logistics|supply chain|freight forwarding|warehousing|shipping|cargo)\b', "Logistics & Supply Chain"),

    # Healthcare & Science
    (r'\b(?:pharmaceutical|pharma|biotechnology|biotech|clinical research)\b', "Pharmaceuticals & Biotechnology"),
    (r'\b(?:medical devices|healthcare solutions|diagnostics|hospital equipment)\b', "Healthcare & Medical Devices"),

    # Construction & Energy
    (r'\b(?:civil engineering|construction|real estate development|building contractor)\b', "Civil Engineering & Construction"),
    (r'\b(?:renewable energy|solar energy|solar power|clean energy|wind energy)\b', "Renewable Energy & Solar Power"),
]

# Patterns for dynamic extraction from text when taxonomy isn't directly matched
DYNAMIC_PATTERNS = [
    re.compile(r'\bspeciali[sz](?:es|ing|ed)?\s+in\s+([^,.;]{3,50})', re.IGNORECASE),
    re.compile(r'\b(?:leading|major|premier)?\s*(?:provider|supplier|developer|manufacturer)\s+of\s+([^,.;]{3,50})', re.IGNORECASE),
    re.compile(r'\bengaged\s+in\s+(?:the\s+)?(?:business\s+of\s+)?([^,.;]{3,50})', re.IGNORECASE),
    re.compile(r'\bdeals?\s+in\s+([^,.;]{3,50})', re.IGNORECASE),
    re.compile(r'\bcore\s+expertise\s+spans\s+([^,.;]{3,50})', re.IGNORECASE),
    re.compile(r'\bdelivers?\s+(?:expert\s+)?(?:consulting\s+)?services\s+in\s+([^,.;]{3,50})', re.IGNORECASE),
    re.compile(r'\b(?:is|are)\s+an?\s+([^,.;]{3,50})\s+(?:company|firm|provider|enterprise|agency)\b', re.IGNORECASE),
]

# Company name fallback domain triggers
COMPANY_FALLBACKS: List[Tuple[str, str]] = [
    (r'\b(?:control\s*systems?)\b', "Control Systems & Automation"),
    (r'\b(?:automations?|robotic|robotics)\b', "Industrial Automation"),
    (r'\b(?:data\s*com(?:munications?)?|telecom)\b', "Telecommunications & Networking"),
    (r'\b(?:technolog(?:ies|y)|tech|info(?:tech|base)?)\b', "IT Services & Technology"),
    (r'\b(?:software|systems?|solutions?)\b', "Software & IT Solutions"),
    (r'\b(?:computers?)\b', "Computer Systems & Hardware"),
    (r'\b(?:digital)\b', "Digital Solutions & Services"),
    (r'\b(?:engineering)\b', "Engineering Services"),
    (r'\b(?:consulting)\b', "Business & IT Consulting"),
    (r'\b(?:media|marketing)\b', "Digital Marketing & Media"),
    (r'\b(?:pharma|biotech)\b', "Pharmaceuticals & Biotechnology"),
    (r'\b(?:health|hospital)\b', "Healthcare Services"),
    (r'\b(?:finance|financial|bank|capital)\b', "Financial Services"),
    (r'\b(?:motors?|auto)\b', "Automotive & Auto Components"),
    (r'\b(?:infra|construction)\b', "Civil Infrastructure & Construction"),
]

UPPER_ACRONYMS = {
    "IT", "AI", "CNC", "BPO", "KPO", "ERP", "CRM", "CAD", "CAM",
    "CAE", "PLC", "SCADA", "PCB", "SEO", "HR", "SaaS", "VoIP"
}


def format_short_title(text: str) -> str:
    """Formats a string to title case while preserving standard industry acronyms."""
    text = re.sub(r'^[^\w]+|[^\w)]+$', '', text.strip())
    words = text.split()
    formatted = []
    for w in words:
        upper_candidate = w.upper().rstrip(',.;:')
        if upper_candidate in UPPER_ACRONYMS:
            trailing = w[len(upper_candidate):]
            formatted.append(upper_candidate + trailing)
        elif w.lower() in {"and", "&"}:
            formatted.append("&")
        elif w.lower() in {"of", "in", "for", "on", "to"}:
            formatted.append(w.lower())
        else:
            formatted.append(w.capitalize())

    res = " ".join(formatted)
    if res and res[0].islower():
        res = res[0].upper() + res[1:]
    return res


def clean_extracted_phrase(phrase: str) -> str:
    """Cleans an extracted dynamic phrase into a concise 2-4 word domain name."""
    phrase = re.sub(
        r'^(?:a|an|the|various|variety\s+of|all\s+types\s+of|leading|range\s+of|broad\s+range\s+of|world\s+class)\s+',
        '', phrase, flags=re.IGNORECASE
    )
    phrase = re.sub(r'\s+(?:company|firm|ltd|limited|pvt|services|solutions)$', '', phrase, flags=re.IGNORECASE)

    words = phrase.split()
    if len(words) > 4:
        words = words[:4]

    cleaned = " ".join(words).strip()
    return format_short_title(cleaned)


def extract_short_core_business(
    text: Optional[str],
    title: Optional[str] = "",
    company_name: Optional[str] = ""
) -> Optional[str]:
    """
    Extracts a short, concise industry/core business category
    (e.g., 'IT Infrastructure', 'Industrial Automation', 'CNC Machine Tools & Automation').
    Max length: 2 to 4 words.
    """
    comp_clean = (company_name or "").strip()
    title_clean = (title or "").strip()
    text_clean = (text or "").strip()

    combined = f"{comp_clean} {title_clean} {text_clean}".strip()
    if not combined:
        return None

    # 1. Match against predefined industry taxonomy (highest precision)
    for pattern, category in TAXONOMY_MAP:
        if re.search(pattern, combined, re.IGNORECASE):
            return category

    # 2. Try dynamic patterns on snippet or title
    for pat in DYNAMIC_PATTERNS:
        match = pat.search(text_clean) or pat.search(title_clean)
        if match:
            candidate = match.group(1).strip()
            for pattern, category in TAXONOMY_MAP:
                if re.search(pattern, candidate, re.IGNORECASE):
                    return category
            cleaned_candidate = clean_extracted_phrase(candidate)
            if cleaned_candidate and 2 <= len(cleaned_candidate.split()) <= 4:
                return cleaned_candidate

    # 3. Fallback to Company Name triggers
    for pattern, category in COMPANY_FALLBACKS:
        if re.search(pattern, comp_clean, re.IGNORECASE):
            return category

    # 4. Fallback to text triggers
    for pattern, category in COMPANY_FALLBACKS:
        if re.search(pattern, text_clean, re.IGNORECASE):
            return category

    return None
