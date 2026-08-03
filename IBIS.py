from bs4 import BeautifulSoup
import pandas as pd
import os
import re

# --- 1. FOLDER SETUP ---
main_folder = "IBIS"
os.makedirs(main_folder, exist_ok=True)
nested_folder = os.path.join(main_folder, "IBIS")
os.makedirs(nested_folder, exist_ok=True)

# --- 2. DATA LOADING ---
excel_path = r"C:\\Users\\leesil\\Python Projects\\Git Branch\\Master\\OJTI\\export.xlsx"
data = []

if not os.path.exists(excel_path):
    print(f"❌ Warning: Excel file not found at {excel_path}. Skipping data processing.")
else:
    try:
        df = pd.read_excel(excel_path, header=None)
        data = df.values.tolist()
        print(f"✅ Loaded {len(data)} rows from Excel.")
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        data = []

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
generation_names = ["MGT", "Mirae", "Gen2", "Gen 3", "Gen 32", "General"]
generation_display_names = {
    "MGT":     "MGT",
    "Mirae":   "MIS",
    "Gen2":    "Gen2",
    "Gen 3":   "Gen 3",
    "Gen 32":  "Gen 32",
    "General": "General",
}

SKILL_ROLE = {
    1: "Operator",
    2: "Operator",
    3: "Technician",
    4: "Technician",
    5: "Technician",
    6: "Technician",
}

LEVEL_HEADER = "<tr class='header-row'><th>Level/Role</th><th>Documents</th></tr>"

generation_skill_lists = {
    generation: {f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)}
    for generation in generation_names
}

# ---------------------------------------------------------------------------
# EXCLUSION LISTS (for General tab only)
# ---------------------------------------------------------------------------
packing_equipment_names = ["DPL", "Despatch"]
ibis_names              = ["MGT", "Mirae", "Gen2", "Gen 3", "Gen 32"]
keywords                = ["DPLP", "Subcon", "Shipping"]
equipment_names         = ["KLA", "SRM", "ISMECA", "TTM", "ETM", "Peel Force Tester"]
automation_names        = ["OHT", "AMR", "E-rack", "Strapping", "ASRS", "AMHS", "MMR", "Point to Point"]
tester_names            = ["V93K", "LTX", "MMCI", "Advantest", "Ultraflex", "Rasco", "EXAScale", "J750"]
handler_names           = ["North Star", "Delta Matrix", "Delta Castle", "OSAI", "Multitest", "JHT"]

GENERAL_EXCLUDE_TERMS = (
    packing_equipment_names
    + ibis_names
    + keywords
    + equipment_names
    + automation_names
    + tester_names
    + handler_names
)

# ---------------------------------------------------------------------------
# SPECIFIC ID ROUTING TABLE
# doc_id -> list of tabs to add to
# Add any future special cases here without touching routing logic
# ---------------------------------------------------------------------------
SPECIFIC_ID_ROUTES = {
    "Z8F46861657": ["Gen2",  "Gen 3"],               # Basic Operation of IBIS Oven
    "Z8F46861656": ["Mirae", "MGT", "Gen2", "Gen 32"], # Inteligent Burn-in System (IBIS)
    "Z8F46860272": ["Gen 3", "Gen 32"],               # Gen 3 IBIS Process
}

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def make_entry(value, link):
    clean_val  = str(value).strip()
    clean_link = str(link).strip()
    url        = f"https://plmpublishing.icp.infineon.com/api/download-pdf/{clean_link}"
    return f"{clean_val} - <a href='{url}' target='_blank'>{clean_link}</a>"

def add_entry(generation, skill_level_key, entry):
    target = generation_skill_lists[generation][skill_level_key]
    if entry not in target["list"]:
        target["list"].append(entry)
        target["count"] += 1

def detect_skill_level(value):
    s = str(value)
    match = re.search(r"\(Skill\s+[Ll]evel\s*(\d)\)", s)
    if match:
        return int(match.group(1))
    return None

def detect_all_generations(value):
    text  = str(value).lower()
    found = []

    if re.search(r"gen[\s\-]*32", text):
        found.append("Gen 32")
    if re.search(r"gen[\s\-]*3", text) and "Gen 32" not in found:
        found.append("Gen 3")
    if re.search(r"gen[\s\-]*2", text):
        found.append("Gen2")
    if re.search(r"mgt", text) or re.search(r"mb8[1i]", text):
        found.append("MGT")
    if re.search(r"mis", text) or re.search(r"mirae", text) or re.search(r"m850[1i]", text):
        found.append("Mirae")
    if re.search(r"\bgeneral\b", text):
        found.append("General")

    seen   = set()
    unique = []
    for g in found:
        if g not in seen:
            seen.add(g)
            unique.append(g)

    return unique if unique else ["General"]

def should_exclude_from_general(title):
    t = str(title).lower()
    return any(term.lower() in t for term in GENERAL_EXCLUDE_TERMS)

def is_ibis_row(col1, col6):
    t1 = str(col1).lower()
    t6 = str(col6).lower()

    if "ibis" in t6:
        return True

    col6_empty = col6.strip() == "" or col6.strip().lower() in ["nan", "none"]
    if col6_empty:
        if re.search(r"\bibis\b", t1):
            return True
        if re.search(r"\bmgt\b", t1) or re.search(r"\bmb8[1i]\b", t1):
            return True
        if re.search(r"\bmirae\b", t1) or re.search(r"\bm850[1i]\b", t1):
            return True
        if re.search(r"\bgen[\s\-]*32\b", t1) or re.search(r"\bgen[\s\-]*3\b", t1) or re.search(r"\bgen[\s\-]*2\b", t1):
            return True

    return False

# ---------------------------------------------------------------------------
# DATA PROCESSING
# ---------------------------------------------------------------------------
ibis_count       = 0
skipped_count    = 0
excluded_general = 0

for row in data:
    padded_row = list(row) + [""] * max(0, 6 - len(row))
    col1 = str(padded_row[0]).strip() if pd.notna(padded_row[0]) else ""
    col2 = str(padded_row[1]).strip() if pd.notna(padded_row[1]) else ""
    col6 = str(padded_row[5]).strip() if pd.notna(padded_row[5]) else ""

    if not col1:
        continue

    if not is_ibis_row(col1, col6):
        skipped_count += 1
        continue

    ibis_count += 1

    skill_level = detect_skill_level(col1)
    skill_level = skill_level if skill_level in range(1, 7) else 1
    skill_key   = f"Skill Level {skill_level}"
    entry       = make_entry(col1, col2)
    title_lower = col1.lower()
    is_oven     = "oven" in title_lower

    # -----------------------------------------------------------------------
    # RULE 0: Specific ID overrides (highest priority)
    # -----------------------------------------------------------------------
    if col2 in SPECIFIC_ID_ROUTES:
        for gen in SPECIFIC_ID_ROUTES[col2]:
            add_entry(gen, skill_key, entry)
        continue

    # -----------------------------------------------------------------------
    # RULE 1: Skill Level 3 -> ALL tabs
    # -----------------------------------------------------------------------
    if skill_level == 3:
        for gen in generation_names:
            add_entry(gen, skill_key, entry)
        continue

    detected_tabs = detect_all_generations(col1)

    # -----------------------------------------------------------------------
    # RULE 2: Skill Level 4 -> detected tabs else Gen2
    # -----------------------------------------------------------------------
    if skill_level == 4:
        if detected_tabs == ["General"]:
            add_entry("Gen2", skill_key, entry)
        else:
            for t in detected_tabs:
                add_entry(t, skill_key, entry)
        continue

    # -----------------------------------------------------------------------
    # RULE 3: Oven -> detected tabs (apply General exclusion)
    # -----------------------------------------------------------------------
    if is_oven:
        for t in detected_tabs:
            if t == "General":
                if should_exclude_from_general(col1):
                    excluded_general += 1
                else:
                    add_entry("General", skill_key, entry)
            else:
                add_entry(t, skill_key, entry)
        continue

    # -----------------------------------------------------------------------
    # RULE 4: Default -> detected tabs (apply General exclusion)
    # -----------------------------------------------------------------------
    for t in detected_tabs:
        if t == "General":
            if should_exclude_from_general(col1):
                excluded_general += 1
            else:
                add_entry("General", skill_key, entry)
        else:
            add_entry(t, skill_key, entry)

print(f"\n📊 Rows processed       : {ibis_count} IBIS rows")
print(f"📊 Rows skipped         : {skipped_count} non-IBIS rows")
print(f"🚫 Blocked from General : {excluded_general} rows")

# ---------------------------------------------------------------------------
# SHARED STYLE
# ---------------------------------------------------------------------------
SHARED_STYLE = """
    body { font-family: Arial, sans-serif; font-size: 16px; color: #333; margin: 20px; background-color: #f0f0f0; }
    a { color: #08665c; text-decoration: none; }
    a:hover { color: #054a40; text-decoration: underline; }
"""

# --- 4. MAIN DASHBOARD ---
html = "<html><head><style>"
html += SHARED_STYLE
html += """
    .container { width: 80%; margin: 0 auto; background-color: #fff; padding: 20px; border: 1px solid #ddd; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
    .header { background-color: #f0f0f0; padding: 10px; border-bottom: 1px solid #ddd; display: flex; justify-content: flex-start; align-items: center; }
    .home-btn { display: inline-block; padding: 8px 16px; background-color: #08665c; color: #fff; text-decoration: none; border-radius: 5px; font-size: 16px; transition: background-color 0.3s; }
    .home-btn:hover { background-color: #054a40; text-decoration: none; }
    .content { padding: 20px; display: flex; }
    .sidebar { width: 20%; background-color: #08665c; color: #fff; padding: 20px; font-size: 24px; text-align: center; border-radius: 10px 0 0 10px; display: flex; justify-content: center; align-items: center; }
    .main-content { width: 80%; padding: 20px; }
    .post-it-container { display: flex; flex-wrap: wrap; justify-content: center; }
    .post-it { width: 150px; height: 150px; background-color: #ADD8E6; padding: 10px; border: 1px solid #ccc; border-radius: 10px; margin: 10px; display: flex; justify-content: center; align-items: center; cursor: pointer; transition: background-color 0.2s ease-in-out; }
    .post-it:hover { background-color: #87CEEB; }
    .post-it a { text-decoration: none; color: #333; font-size: 16px; width: 100%; height: 100%; display: flex; justify-content: center; align-items: center; text-align: center; }
    .post-it:hover a { color: #fff; }
"""
html += "</style></head><body><div class='container'>"
html += "<div class='header'><a href='../index.html' class='home-btn'>Return to Home</a></div>"
html += "<div class='content'><div class='sidebar'>IBIS</div>"
html += "<div class='main-content'><div class='post-it-container'>"

for generation in generation_names:
    display = generation_display_names.get(generation, generation)
    html += f"<div class='post-it'><a href='IBIS/{generation}.html'>{display}</a></div>"

html += "</div></div></div></body></html>"

with open(os.path.join(main_folder, "IBIS.html"), "w", encoding="utf-8") as f:
    print(BeautifulSoup(html, "html.parser").prettify(), file=f)

# --- 5. INDIVIDUAL GENERATION PAGES ---
for generation in generation_names:
    display = generation_display_names.get(generation, generation)

    generation_html = "<html><head><style>"
    generation_html += SHARED_STYLE
    generation_html += """
    .back-btn { display: inline-block; padding: 10px 20px; background-color: #08665c; color: #fff; text-decoration: none; border-radius: 5px; font-size: 16px; margin-bottom: 20px; transition: background-color 0.3s; }
    .back-btn:hover { background-color: #054a40; text-decoration: none; }
    h1 { font-size: 24px; color: #333; margin-bottom: 20px; }
    .data-table { border-collapse: collapse; width: 100%; font-size: 16px; color: #333; box-shadow: 0 0 10px rgba(0,0,0,0.1); background-color: #fff; }
    .header-row { background-color: #f0f0f0; }
    .header-row th { padding: 12px; text-align: left; font-size: 16px; color: #333; font-weight: bold; border-bottom: 2px solid #ddd; }
    .level-cell { text-align: left; font-size: 16px; color: #333; padding: 12px; border-bottom: 1px solid #ddd; white-space: nowrap; vertical-align: top; }
    .data-cell { text-align: left; font-size: 16px; color: #333; padding: 12px; border-bottom: 1px solid #ddd; line-height: 1.5; }
    .data-cell:hover { background-color: #f0f0f0; }
    .data-cell a { color: #08665c; text-decoration: none; font-size: 16px; }
    .data-cell a:hover { color: #054a40; text-decoration: underline; }
    """
    generation_html += "</style></head><body>"
    generation_html += "<a href='../IBIS.html' class='back-btn'>← Back to Dashboard</a>"
    generation_html += f"<h1>{display}</h1>"
    generation_html += "<table class='data-table'>"
    generation_html += LEVEL_HEADER

    has_data = False
    for i in range(1, 7):
        sk   = f"Skill Level {i}"
        docs = generation_skill_lists[generation][sk]["list"]
        if docs:
            has_data     = True
            role         = SKILL_ROLE.get(i, "")
            sorted_docs  = sorted(docs, key=lambda x: x.lower())
            docs_html    = "<br/><br/>".join(sorted_docs)
            generation_html += (
                f"<tr>"
                f"<td class='level-cell'>Level {i} ({role})</td>"
                f"<td class='data-cell'>{docs_html}</td>"
                f"</tr>"
            )

    if not has_data:
        generation_html += "<tr><td class='level-cell' colspan='2'>No documents found for this generation.</td></tr>"

    generation_html += "</table></body></html>"

    with open(os.path.join(nested_folder, f"{generation}.html"), "w", encoding="utf-8") as f:
        print(BeautifulSoup(generation_html, "html.parser").prettify(), file=f)

print("\n✅ Processing Complete.")
print("📁 Dashboard : IBIS/IBIS.html")
print("📁 Pages     : IBIS/IBIS/[Generation].html")
print("📋 Specific ID overrides applied")
print("📋 Strict IBIS filter + General exclusion list applied")
print("📋 Sorting : Level (1-6), then document name A-Z")