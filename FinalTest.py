import os
import re
import pandas as pd
from bs4 import BeautifulSoup

# ==========================================
# 1. CONFIGURATION
# ==========================================
output_folder     = "Final Test"
testers_folder    = os.path.join(output_folder, "Testers")
handlers_folder   = os.path.join(output_folder, "Handlers")
automation_folder = os.path.join(output_folder, "Automation")
general_folder    = os.path.join(output_folder, "General")

for folder in [output_folder, testers_folder, handlers_folder, automation_folder, general_folder]:
    os.makedirs(folder, exist_ok=True)

# UPDATE THIS PATH TO YOUR EXCEL FILE
excel_path = r"C:\\Users\\leesil\\Python Projects\\Git Branch\\Master\\OJTI\\export.xlsx"

automation_names   = ["OHT", "AMR", "E-rack", "Strapping", "ASRS", "AMHS", "MMR", "Point to Point"]
tester_names       = ["EXAScale", "V93K", "LTX", "MMCI", "Advantest", "Ultraflex", "J750", "Flex"]
handler_names      = ["North Star", "Delta Matrix", "Delta Castle", "OSAI", "Multitest", "JHT", "Rasco"]

keywords = tester_names + handler_names + automation_names

SKILL_ROLE = {
    1: "Operator/Technician",
    2: "Operator",
    3: "Operator/Technician",
    4: "Line Technician",
    5: "PM Technician",
    6: "Technician",
}

def level_role_label(i: int) -> str:
    return f"Level {i} ({SKILL_ROLE.get(i, '')})"

SKILL_RE = re.compile(r"skill\s*level\s*[:\-]?\s*(\d)", re.IGNORECASE)

def extract_skill_level(text: str):
    """Return int skill level 1..6 if found, else None."""
    m = SKILL_RE.search(str(text))
    if not m:
        return None
    lvl = int(m.group(1))
    return lvl if 1 <= lvl <= 6 else None

# ==========================================
# 2. LOAD DATA
# ==========================================
data = []
if not os.path.exists(excel_path):
    print(f"❌ Warning: Excel file not found at {excel_path}.")
else:
    try:
        df = pd.read_excel(excel_path, header=None)
        data = df.values.tolist()
        print(f"✅ Loaded {len(data)} rows from Excel.")
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        data = []

# ==========================================
# 3. PROCESS DATA
# ==========================================
keyword_lists = {
    name: {f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)}
    for name in keywords
}
general_lists = {f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)}
matched_links = set()

# --------------------------------------------------------------------------
# HARDCODED DOCS (ALWAYS INCLUDE)
# doc_id -> {"title": "...", "targets": [kw1, kw2, ...], "skill": 3}
#
# targets:
#   - use "ALL" to add to ALL tester/handler/automation keyword tabs
#   - use "GENERAL" to add to General
#   - or list specific keyword names from tester_names/handler_names/automation_names
# --------------------------------------------------------------------------
HARDCODED_DOCS = {
    "Z8F46860675": {
        "title": "OJTI for IBIS Equipment Safe Release (ESR). (Skill Level 3)(TWI)",
        "targets": ["ALL"],
        "skill": 3
    },
    "Z8F46861705": {
        "title": "OJTI on Defect Catalogue for Mechanical Rejects of BGA Packages.(Skill level 3)",
        "targets": ["ALL"],
        "skill": 3
    },
    "Z8F46861706": {
        "title": "OJTI for Defect Catalogue on Mechanical Reject of Leaded and Leadless Packages. (Skill Level 3)",
        "targets": ["ALL"],
        "skill": 3
    },
    "Z8F46861707": {
        "title": "OJTI for Conduct of Offline Defect Sensitivity Assessment(DSA). (Skill Level 3)",
        "targets": ["ALL"],
        "skill": 3
    },
    "Z8F46861711": {
        "title": "OJTI on Process Specification, Control Systems & Alarm Handling for Test Process.(Skill Level 3)(TWI)",
        "targets": ["ALL"],
        "skill": 3
    },
    "Z8F46860019": {
        "title": "OJTI on Final Test Tube/Tray Device Handling During Technical Intervention. (Skill Level 3) (TWI)",
        "targets": ["ALL"],
        "skill": 3
    },
}

def make_entry(title, link):
    title = str(title).strip()
    link  = str(link).strip()
    return (
        f"{title} - "
        f"<a href='https://plmpublishing.icp.infineon.com/api/download-pdf/{link}' "
        f"target='_blank'>{link}</a>"
    )

def add_to_keyword(kw, lvl, entry):
    keyword_lists[kw][f"Skill Level {lvl}"]["list"].append(entry)
    keyword_lists[kw][f"Skill Level {lvl}"]["count"] += 1

def add_to_general(lvl, entry):
    general_lists[f"Skill Level {lvl}"]["list"].append(entry)
    general_lists[f"Skill Level {lvl}"]["count"] += 1

# ==========================================
# KEYWORD MATCHING LOGIC
# ==========================================

# 1. Define Abbreviation Mappings
# Add more here if needed (e.g., "DM": "Delta Matrix")
ABBREV_MAP = {
    "DC": "Delta Castle",
    # "DM": "Delta Matrix",
}

# 2. Build Regex Patterns for Full Keywords
def build_keyword_pattern(kw: str) -> re.Pattern:
    """
    Build a word-boundary-aware regex for a keyword.
    - Supports multi-word keywords (e.g., 'Delta Castle')
    - Uses word boundaries (\b) to prevent partial matches (e.g., 'DC' inside 'ADC')
    """
    escaped = re.escape(kw)
    # Allow flexible separators (space, hyphen, underscore) between words in the keyword
    escaped = re.sub(r'\\ ', r'[\\s\\-_]+', escaped)
    # Use \b for word boundaries. Note: \b doesn't always work perfectly with regex special chars,
    # so we use a lookbehind/lookahead for non-alphanumeric characters instead for safety.
    pattern_str = r'(?<![A-Za-z0-9])' + escaped + r'(?![A-Za-z0-9])'
    return re.compile(pattern_str, re.IGNORECASE)

KEYWORD_PATTERNS = {kw: build_keyword_pattern(kw) for kw in keywords}

# 3. Build Regex Pattern for Abbreviations
# Matches standalone uppercase/acronym strings (2-4 chars) that are not part of a larger word
ABBREV_RE = re.compile(r'(?<![A-Za-z0-9])([A-Z]{2,4})(?![A-Za-z0-9])')

# ==========================================
# EXECUTION
# ==========================================

print("🔄 Processing data...")
total_processed     = 0
debug_general_added = 0

# Set to True if you want to see detailed logs of matches in the console
DEBUG_MODE = False 

# --- A) Normal keyword-based processing from Excel ---
for row in data:
    if len(row) < 2:
        continue

    value = str(row[0]) if pd.notna(row[0]) else ""
    link  = str(row[1]) if pd.notna(row[1]) else ""
    col6  = str(row[5]) if (len(row) > 5 and pd.notna(row[5])) else ""

    if not value.strip() or not link.strip():
        continue

    # Skip if it is in hardcoded docs (we inject separately)
    if link in HARDCODED_DOCS:
        continue

    lvl = extract_skill_level(value)
    if lvl is None:
        continue

    entry = make_entry(value, link)
    matched_specific = False

    # 1) Check for Abbreviations First (e.g., "DC" -> "Delta Castle")
    # Find all standalone acronyms in the title
    found_abbrevs = set(ABBREV_RE.findall(value))
    for abbr in found_abbrevs:
        if abbr in ABBREV_MAP:
            target_kw = ABBREV_MAP[abbr]
            if DEBUG_MODE:
                print(f"🔹 ABBREVIATION MATCH: '{abbr}' -> '{target_kw}' | Title: {value[:50]}...")
            add_to_keyword(target_kw, lvl, entry)
            total_processed += 1
            matched_specific = True
            matched_links.add(link)
            break # Matched one abbreviation, move to next row

    if matched_specific:
        continue

    # 2) Check for Full Keyword Matches (e.g., "Delta Castle")
    for kw in keywords:
        if KEYWORD_PATTERNS[kw].search(value):
            if DEBUG_MODE:
                print(f"🔹 KEYWORD MATCH: '{kw}' | Title: {value[:50]}...")
            add_to_keyword(kw, lvl, entry)
            total_processed += 1
            matched_specific = True
            matched_links.add(link)
            break  # Remove this 'break' if you want a doc to appear in multiple tabs

    if matched_specific:
        continue

    # 3) General: col6 must contain "testing" and not matched
    if "testing" in col6.lower() and link not in matched_links:
        if DEBUG_MODE:
            print(f"🔹 GENERAL MATCH (via col6) | Title: {value[:50]}...")
        add_to_general(lvl, entry)
        total_processed += 1
        debug_general_added += 1

# --- B) Hardcoded injection (ALWAYS INCLUDE) ---
hardcoded_added = 0
for doc_id, cfg in HARDCODED_DOCS.items():
    lvl = int(cfg.get("skill", 3))
    title = cfg.get("title", doc_id)
    entry = make_entry(title, doc_id)
    targets = cfg.get("targets", ["ALL"])

    if "ALL" in targets:
        for kw in keywords:
            add_to_keyword(kw, lvl, entry)
        hardcoded_added += 1
        continue

    if "GENERAL" in targets:
        add_to_general(lvl, entry)
        hardcoded_added += 1
        continue

    # Specific targets
    for kw in targets:
        if kw in keyword_lists:
            add_to_keyword(kw, lvl, entry)
    hardcoded_added += 1

print(f"✅ Total entries processed (Excel rules) : {total_processed}")
print(f"➕ Added to General (Excel rules)        : {debug_general_added}")
print(f"📌 Hardcoded docs injected              : {hardcoded_added}")

total_general = sum(general_lists[f"Skill Level {i}"]["count"] for i in range(1, 7))
print(f"📌 General total                         : {total_general}")

print("\n📊 Keyword counts:")
for kw in keywords:
    c = sum(keyword_lists[kw][f"Skill Level {i}"]["count"] for i in range(1, 7))
    print(f"   {kw:<20} : {c}")

# ==========================================
# 4. CSS DEFINITIONS
# ==========================================
dashboard_css = """
body { font-family: Arial, sans-serif; margin: 20px; background-color: #f0f0f0; }
.container { width: 80%; margin: 0 auto; background-color: #fff; padding: 20px; border: 1px solid #ddd; border-radius: 10px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); }
.header { background-color: #f0f0f0; padding: 15px; border-bottom: 1px solid #ddd; display: flex; justify-content: space-between; align-items: center; }
.header h1 { margin: 0; text-align: center; color: #333; }
.content { padding: 20px; display: flex; }
.sidebar { width: 20%; background-color: #08665c; color: #fff; padding: 20px; font-size: 24px; text-align: center; border-radius: 10px 0 0 10px; display: flex; justify-content: center; align-items: center; }
.main-content { width: 80%; padding: 20px; }
.post-it-container { display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; }
.post-it { width: 160px; height: 160px; background-color: #ADD8E6; padding: 10px; border: 1px solid #ccc; border-radius: 10px; display: flex; justify-content: center; align-items: center; cursor: pointer; transition: all 0.2s; }
.post-it:hover { background-color: #87CEEB; transform: translateY(-5px); }
.post-it a { text-decoration: none; color: #000; width: 100%; height: 100%; display: flex; justify-content: center; align-items: center; font-size: 16px; }
.post-it:hover a { color: #fff; }
.back-btn { display: inline-block; padding: 8px 16px; background-color: #08665c; color: white; text-decoration: none; border-radius: 5px; font-size: 14px; transition: background-color 0.3s; }
.back-btn:hover { background-color: #054a40; }
.header-btn { padding: 6px 12px; font-size: 13px; }
"""

detail_css = """
body { font-family: Arial, sans-serif; margin: 20px; background-color: #f0f0f0; }
.header { background-color: #f0f0f0; padding: 15px; border-bottom: 1px solid #ddd; display: flex; justify-content: space-between; align-items: center; border-radius: 10px 10px 0 0; margin-bottom: 20px; }
.header h1 { margin: 0; color: #333; font-size: 24px; }
.back-btn { display: inline-block; padding: 8px 16px; background-color: #08665c; color: white; text-decoration: none; border-radius: 5px; font-size: 14px; transition: background-color 0.3s; }
.back-btn:hover { background-color: #054a40; }
.data-table { border-collapse: collapse; width: 100%; font-size: 18px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); background-color: #fff; border-radius: 0 0 10px 10px; }
.header-row { background-color: #f0f0f0; color: #333; font-weight: bold; }
.header-row th { padding: 12px; text-align: left; }
.skill-level { text-align: left; font-size: 16px; padding: 12px; border-bottom: 1px solid #ddd; }
.data-cell { text-align: left; font-size: 16px; padding: 12px; border-bottom: 1px solid #ddd; line-height: 1.5; }
.data-cell.empty { background-color: #cccccc; }
.data-cell:hover { background-color: #f0f0f0; }
"""

LEVEL_HEADER = "<tr class='header-row'><th>Level/Role</th><th>Documents</th></tr>"

# ==========================================
# 5. GENERATE HTML
# ==========================================
def write_detail_page(title, back_href, skill_dict, output_path, empty_msg):
    html  = f"<html><head><style>{detail_css}</style></head><body>"
    html += f"<div class='header'><h1>{title}</h1>"
    html += f"<a href='{back_href}' class='back-btn header-btn'>← Back</a></div>"
    html += f"<table class='data-table'>{LEVEL_HEADER}"

    rows_found = False
    for i in range(1, 7):
        entries = skill_dict[f"Skill Level {i}"]["list"]
        if entries:
            rows_found = True
            html += (
                f"<tr><td class='skill-level'>{level_role_label(i)}</td>"
                f"<td class='data-cell'>{'<br/><br/>'.join(entries)}</td></tr>"
            )

    if not rows_found:
        html += f"<tr><td class='skill-level' colspan='2'>{empty_msg}</td></tr>"

    html += "</table></body></html>"

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(BeautifulSoup(html, "html.parser").prettify())

# --- MASTER DASHBOARD ---
html_master  = f"<html><head><style>{dashboard_css}</style></head><body>"
html_master += "<div class='container'><div class='header'>"
html_master += "<a href='../index.html' class='back-btn header-btn'>Return to Main Page</a>"
html_master += "</div><div class='content'>"
html_master += "<div class='sidebar'>Final Test</div>"
html_master += "<div class='main-content'><div class='post-it-container'>"
html_master += "<div class='post-it'><a href='Testers/view.html'>Tester</a></div>"
html_master += "<div class='post-it'><a href='Handlers/view.html'>Handler</a></div>"
html_master += "<div class='post-it'><a href='Automation/view.html'>Automation</a></div>"
html_master += "<div class='post-it'><a href='General/view.html'>General</a></div>"
html_master += "</div></div></div></div></body></html>"

with open(os.path.join(output_folder, "FinalTest.html"), "w", encoding="utf-8") as f:
    f.write(BeautifulSoup(html_master, "html.parser").prettify())

# --- TESTER LIST VIEW ---
html_tester_view = (
    f"<html><head><style>{dashboard_css}</style></head><body>"
    f"<div class='container'><div class='header'>"
    f"<a href='../FinalTest.html' class='back-btn header-btn'>← Back to Final Test</a>"
    f"</div><div class='content'><div class='sidebar'>Testers</div>"
    f"<div class='main-content'><div class='post-it-container'>"
)
for t in tester_names:
    html_tester_view += f"<div class='post-it'><a href='{t}.html'>{t}</a></div>"
html_tester_view += "</div></div></div></div></body></html>"

with open(os.path.join(testers_folder, "view.html"), "w", encoding="utf-8") as f:
    f.write(BeautifulSoup(html_tester_view, "html.parser").prettify())

# --- HANDLER LIST VIEW ---
html_handler_view = (
    f"<html><head><style>{dashboard_css}</style></head><body>"
    f"<div class='container'><div class='header'>"
    f"<a href='../FinalTest.html' class='back-btn header-btn'>← Back to Final Test</a>"
    f"</div><div class='content'><div class='sidebar'>Handlers</div>"
    f"<div class='main-content'><div class='post-it-container'>"
)
for h in handler_names:
    safe_name = h.replace(" ", "_")
    html_handler_view += f"<div class='post-it'><a href='{safe_name}.html'>{h}</a></div>"
html_handler_view += "</div></div></div></div></body></html>"

with open(os.path.join(handlers_folder, "view.html"), "w", encoding="utf-8") as f:
    f.write(BeautifulSoup(html_handler_view, "html.parser").prettify())

# --- AUTOMATION LIST VIEW ---
html_auto_view = (
    f"<html><head><style>{dashboard_css}</style></head><body>"
    f"<div class='container'><div class='header'>"
    f"<a href='../FinalTest.html' class='back-btn header-btn'>← Back to Final Test</a>"
    f"</div><div class='content'><div class='sidebar'>Automation</div>"
    f"<div class='main-content'><div class='post-it-container'>"
)
for a in automation_names:
    safe_name = a.replace(" ", "_")
    html_auto_view += f"<div class='post-it'><a href='{safe_name}.html'>{a}</a></div>"
html_auto_view += "</div></div></div></div></body></html>"

with open(os.path.join(automation_folder, "view.html"), "w", encoding="utf-8") as f:
    f.write(BeautifulSoup(html_auto_view, "html.parser").prettify())

# --- DETAIL PAGES ---
for t in tester_names:
    write_detail_page(
        t,
        "view.html",
        keyword_lists[t],
        os.path.join(testers_folder, f"{t}.html"),
        "No documents found for this Tester."
    )

for h in handler_names:
    safe_name = h.replace(" ", "_")
    write_detail_page(
        h,
        "view.html",
        keyword_lists[h],
        os.path.join(handlers_folder, f"{safe_name}.html"),
        "No documents found for this Handler."
    )

for a in automation_names:
    safe_name = a.replace(" ", "_")
    write_detail_page(
        a,
        "view.html",
        keyword_lists[a],
        os.path.join(automation_folder, f"{safe_name}.html"),
        "No documents found for this Automation Equipment."
    )

write_detail_page(
    "General",
    "../FinalTest.html",
    general_lists,
    os.path.join(general_folder, "view.html"),
    "No documents found for General."
)

print("\n✅ HTML generated.")