from bs4 import BeautifulSoup
import pandas as pd
import os
import re

# --- 1. FOLDER SETUP ---
main_folder = "MSP"
os.makedirs(main_folder, exist_ok=True)
nested_folder = os.path.join(main_folder, "MSP")
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
# CHANGED: "Despatch" -> "Baking"
equipment_names = ["KLA", "SRM", "ISMECA", "TTM", "ETM", "Peel Force Tester", "KLR", "McDry", "Baking"]
ALL_EQUIPMENT   = equipment_names

SKILL_ROLE = {
    1: "Operator",
    2: "Operator",
    3: "Technician",
    4: "Technician",
    5: "Technician",
    6: "Technician",
}

def level_role_label(i: int) -> str:
    return f"Level {i} ({SKILL_ROLE.get(i, '')})"

LEVEL_HEADER = "<tr class='header-row'><th>Level/Role</th><th>Documents</th></tr>"

equipment_skill_lists = {
    equipment: {f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)}
    for equipment in equipment_names
}

# ---------------------------------------------------------------------------
# SPECIFIC ID ROUTING TABLE
# These 6 docs are hardcoded into ALL equipment tabs at Skill Level 3
# ---------------------------------------------------------------------------
SPECIFIC_ID_ROUTES = {
    "Z8F46861705": ALL_EQUIPMENT,  # OJTI on Defect Catalogue for Mechanical Rejects of BGA Packages
    "Z8F46861706": ALL_EQUIPMENT,  # OJTI for Defect Catalogue on Mechanical Reject of Leaded and Leadless Packages
    "Z8F46861707": ALL_EQUIPMENT,  # OJTI for Conduct of Offline Defect Sensitivity Assessment (DSA)
    "Z8F46861709": ALL_EQUIPMENT,  # OJTI on Process Specification, Control System & Alarm Handling for MSP Process
    "Z8F46860659": ALL_EQUIPMENT,  # OJTI for MSP Equipment Safe Release (ESR)
    "Z8F46860837": ALL_EQUIPMENT,  # OJTI for Tube and Tray Transfer Machine Device Handling Procedure
}

FORCE_SKILL_LEVEL_BY_ID = {
    "Z8F46861705": 3,
    "Z8F46861706": 3,
    "Z8F46861707": 3,
    "Z8F46861709": 3,
    "Z8F46860659": 3,
    "Z8F46860837": 3,
}

# ---------------------------------------------------------------------------
# BLOCKED DOC IDs (never appear anywhere)
# ---------------------------------------------------------------------------
BLOCKED_DOC_IDS = set()

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def detect_skill_level(title: str):
    m = re.search(r"\(\s*Skill\s*[Ll]evel\s*([1-6])\s*\)", str(title))
    return int(m.group(1)) if m else None

def make_entry(value, link):
    value = str(value).strip()
    link  = str(link).strip()
    return (
        f"{value} - "
        f"<a href='https://plmpublishing.icp.infineon.com/api/download-pdf/{link}' "
        f"target='_blank'>{link}</a>"
    )

def add_entry(equipment, skill_level_key, entry):
    if equipment not in equipment_skill_lists:
        return
    bucket = equipment_skill_lists[equipment][skill_level_key]
    if entry not in bucket["list"]:
        bucket["list"].append(entry)
        bucket["count"] += 1

def matches_equipment(title: str, equipment: str) -> bool:
    t = str(title)
    
    # Special handling for "Peel Force Tester"
    if equipment.lower() == "peel force tester":
        return bool(
            re.search(r"\bpeel\b", t, flags=re.IGNORECASE) or
            re.search(r"\bpeel\s+force\s+tester\b", t, flags=re.IGNORECASE)
        )
    
    # Special handling for "Baking" (formerly Despatch)
    # We check for "Baking" OR "Despatch" in the title so old titles still route correctly
    if equipment.lower() == "baking":
        return bool(
            re.search(r"\bbaking\b", t, flags=re.IGNORECASE) or
            re.search(r"\bdespatch\b", t, flags=re.IGNORECASE) or
            re.search(r"\bdespatch\s+oven\b", t, flags=re.IGNORECASE)
        )

    # Default: Exact word match for other equipment
    return bool(re.search(rf"\b{re.escape(equipment)}\b", t, flags=re.IGNORECASE))

# ---------------------------------------------------------------------------
# DATA PROCESSING
# Normal rows: keyword matching by equipment name in title
# Hardcoded IDs are always injected at Skill Level 3 into all tabs
# ---------------------------------------------------------------------------
print("🔄 Processing data...")
total_processed = 0
blocked_count   = 0

for row in data:
    if len(row) < 2:
        continue
    
    value = str(row[0]).strip() if pd.notna(row[0]) else ""
    link  = str(row[1]).strip() if pd.notna(row[1]) else ""
    
    if not value:
        continue
    
    if link in BLOCKED_DOC_IDS:
        blocked_count += 1
        continue
        
    # Skip hardcoded IDs from normal processing — handled separately below
    if link in SPECIFIC_ID_ROUTES:
        continue
        
    lvl = detect_skill_level(value)
    if lvl is None:
        continue
        
    skill_key = f"Skill Level {lvl}"
    entry     = make_entry(value, link)
    
    for equipment in equipment_names:
        if matches_equipment(value, equipment):
            add_entry(equipment, skill_key, entry)
            total_processed += 1

print(f"✅ Total entries matched by name : {total_processed}")
print(f"🚫 Blocked entries               : {blocked_count}")

# ---------------------------------------------------------------------------
# INJECT HARDCODED LEVEL 3 IDs INTO ALL TABS
# ---------------------------------------------------------------------------
print("\n🔍 Injecting hardcoded Level 3 IDs into all tabs...")
injected = 0

for doc_id in SPECIFIC_ID_ROUTES:
    # Try to find the title from Excel by doc ID
    found_title = None
    for row in data:
        if len(row) < 2:
            continue
        title = str(row[0]).strip() if pd.notna(row[0]) else ""
        link  = str(row[1]).strip() if pd.notna(row[1]) else ""
        if link == doc_id:
            found_title = title
            break
            
    if found_title:
        entry = make_entry(found_title, doc_id)
        for equipment in ALL_EQUIPMENT:
            add_entry(equipment, "Skill Level 3", entry)
        injected += 1
        print(f"  ✅ Injected : {doc_id} → {found_title}")
    else:
        print(f"  ⚠️  ID not found in Excel : {doc_id}")

print(f"📋 Hardcoded Level 3 injected : {injected}/{len(SPECIFIC_ID_ROUTES)} → all tabs")

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
dashboard_css = """
body { font-family: Arial, sans-serif; margin: 20px; background-color: #f0f0f0; }
.container { width: 80%; margin: 0 auto; background-color: #fff; padding: 20px; border: 1px solid #ddd; border-radius: 10px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); }
.header { background-color: #f0f0f0; padding: 10px; border-bottom: 1px solid #ddd; display: flex; justify-content: flex-start; align-items: center; }
.home-btn { display: inline-block; padding: 8px 16px; background-color: #08665c; color: white; text-decoration: none; border-radius: 5px; font-size: 14px; transition: background-color 0.3s; }
.home-btn:hover { background-color: #054a40; }
.content { padding: 20px; display: flex; }
.sidebar { width: 20%; background-color: #08665c; color: #fff; padding: 20px; font-size: 24px; text-align: center; border-radius: 10px 0 0 10px; display: flex; justify-content: center; align-items: center; }
.main-content { width: 80%; padding: 20px; }
.post-it-container { display: flex; flex-wrap: wrap; justify-content: center; }
.post-it { width: 150px; height: 150px; background-color: #ADD8E6; padding: 10px; border: 1px solid #ccc; border-radius: 10px; margin: 10px; display: flex; justify-content: center; align-items: center; cursor: pointer; transition: background-color 0.2s ease-in-out; }
.post-it:hover { background-color: #87CEEB; }
.post-it a { text-decoration: none; color: #000; width: 100%; height: 100%; display: flex; justify-content: center; align-items: center; transition: color 0.2s ease-in-out; font-weight: normal; }
.post-it:hover a { color: #fff; }
"""

detail_css = """
body { font-family: Arial, sans-serif; margin: 20px; background-color: #f0f0f0; }
.back-btn { display: inline-block; padding: 10px 20px; background-color: #08665c; color: white; text-decoration: none; border-radius: 5px; font-size: 14px; margin-bottom: 20px; transition: background-color 0.3s; }
.back-btn:hover { background-color: #054a40; }
h1 { font-size: 24px; color: #333; margin-bottom: 20px; }
.data-table { border-collapse: collapse; width: 100%; font-size: 18px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); background-color: #fff; }
.header-row { background-color: #f0f0f0; color: #333; font-weight: bold; }
.header-row th { padding: 12px; text-align: left; }
.skill-level { text-align: left; font-size: 16px; padding: 12px; border-bottom: 1px solid #ddd; white-space: nowrap; vertical-align: top; }
.data-cell { text-align: left; font-size: 16px; padding: 12px; border-bottom: 1px solid #ddd; line-height: 1.5; }
.data-cell:hover { background-color: #f0f0f0; }
.data-cell a { color: #08665c; text-decoration: none; font-size: 16px; }
.data-cell a:hover { color: #054a40; text-decoration: underline; }
"""

# --- 3. MAIN DASHBOARD ---
html = f"<html><head><style>{dashboard_css}</style></head><body>"
html += "<div class='container'>"
html += "<div class='header'><a href='../index.html' class='home-btn'>Return to Home</a></div>"
html += "<div class='content'>"
html += "<div class='sidebar'>MSP</div>"
html += "<div class='main-content'><div class='post-it-container'>"

for equipment in equipment_names:
    safe_filename = equipment.replace(" ", "_") + ".html"
    html += f"<div class='post-it'><a href='MSP/{safe_filename}'>{equipment}</a></div>"

html += "</div></div></div></div></body></html>"

with open(os.path.join(main_folder, "MSP.html"), "w", encoding="utf-8") as file:
    print(BeautifulSoup(html, "html.parser").prettify(), file=file)

# --- 4. INDIVIDUAL EQUIPMENT PAGES ---
for equipment in equipment_names:
    safe_filename  = equipment.replace(" ", "_") + ".html"
    equipment_html = f"<html><head><style>{detail_css}</style></head><body>"
    equipment_html += "<a href='../MSP.html' class='back-btn'>← Back to Dashboard</a>"
    equipment_html += f"<h1>{equipment}</h1>"
    equipment_html += "<table class='data-table'>"
    equipment_html += LEVEL_HEADER
    
    has_data = False
    for i in range(1, 7):
        bucket = equipment_skill_lists[equipment][f"Skill Level {i}"]["list"]
        if bucket:
            has_data      = True
            bucket_sorted = sorted(bucket, key=lambda x: x.lower())
            content       = "<br/><br/>".join(bucket_sorted)
            equipment_html += (
                f"<tr>"
                f"<td class='skill-level'>{level_role_label(i)}</td>"
                f"<td class='data-cell'>{content}</td>"
                f"</tr>"
            )
            
    if not has_data:
        equipment_html += (
            "<tr><td class='skill-level' colspan='2'>"
            "No documents found for this equipment."
            "</td></tr>"
        )
        
    equipment_html += "</table></body></html>"
    
    with open(os.path.join(nested_folder, safe_filename), "w", encoding="utf-8") as file:
        print(BeautifulSoup(equipment_html, "html.parser").prettify(), file=file)

print("\n✅ Processing Complete.")
print("📁 Dashboard : MSP/MSP.html")
print("📁 Pages     : MSP/MSP/[Equipment].html")
print("📋 6 hardcoded Level 3 MSP docs routed by ID → all equipment tabs")
print("📋 Normal rows matched by equipment keyword in title")
print("📋 Sorting : Level (1-6), then document name A-Z")