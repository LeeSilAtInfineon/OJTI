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

# Single source of truth for the table header
LEVEL_HEADER = "<tr class='header-row'><th>Level/Role</th><th>Documents</th></tr>"

generation_skill_lists = {
    generation: {f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)}
    for generation in generation_names
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

def detect_primary_generation(value):
    val_lower = str(value).lower()
    if re.search(r"\bgen[\s\-]*32\b", val_lower):
        return "Gen 32"
    if re.search(r"\bgen[\s\-]*3\b", val_lower):
        return "Gen 3"
    if re.search(r"\bgen[\s\-]*2\b", val_lower):
        return "Gen2"
    if re.search(r"\bmb8[1iI]\b", val_lower):
        return "MGT"
    if re.search(r"\bmgt\b", val_lower):
        return "MGT"
    if re.search(r"\bm850[i1]\b", val_lower):
        return "Mirae"
    if re.search(r"\bmirae\b", val_lower):
        return "Mirae"
    return None

# --- 3. DATA PROCESSING ---
ibis_count    = 0
skipped_count = 0

for row in data:
    padded_row = list(row) + [""] * max(0, 6 - len(row))

    col1 = str(padded_row[0]).strip() if pd.notna(padded_row[0]) else ""
    col2 = str(padded_row[1]).strip() if pd.notna(padded_row[1]) else ""
    col6 = str(padded_row[5]).strip() if pd.notna(padded_row[5]) else ""

    if not col1:
        continue

    if "(IBIS)" not in col6:
        skipped_count += 1
        continue

    ibis_count += 1

    skill_level = detect_skill_level(col1)
    skill_level = skill_level if skill_level in range(1, 7) else 1
    skill_key   = f"Skill Level {skill_level}"
    entry       = make_entry(col1, col2)
    is_oven     = "oven" in col1.lower()
    primary_gen = detect_primary_generation(col1)
    title_lower = col1.lower()

    # -----------------------------------------------------------------------
    # RULE 2: ID = Z8F46861657 -> Gen2 AND Gen 3
    # -----------------------------------------------------------------------
    if col2 == "Z8F46861657":
        add_entry("Gen2",  skill_key, entry)
        add_entry("Gen 3", skill_key, entry)
        continue

    # -----------------------------------------------------------------------
    # RULE 3: Skill Level 3
    # NEW EXCEPTION: OJTI + Gen 3 (not Gen 32) -> Gen 3 ONLY
    # Otherwise -> ALL tabs
    # -----------------------------------------------------------------------
    if skill_level == 3:
        is_gen32 = bool(re.search(r"\bgen[\s\-]*32\b", title_lower))
        is_gen3  = bool(re.search(r"\bgen[\s\-]*3\b",  title_lower))

        if "ojti" in title_lower and is_gen3 and not is_gen32:
            # Gen 3 OJTI -> Gen 3 tab only, not all tabs
            add_entry("Gen 3", "Skill Level 3", entry)
        else:
            # Everything else at Skill Level 3 -> ALL tabs
            for gen in generation_names:
                add_entry(gen, "Skill Level 3", entry)
        continue

    # -----------------------------------------------------------------------
    # RULE 4: Skill Level 4 -> specific gen or Gen2
    # -----------------------------------------------------------------------
    if skill_level == 4:
        add_entry(primary_gen if primary_gen else "Gen2", skill_key, entry)
        continue

    # -----------------------------------------------------------------------
    # OVEN RULE: oven + specific gen -> that tab, else General
    # -----------------------------------------------------------------------
    if is_oven:
        add_entry(primary_gen if primary_gen else "General", skill_key, entry)
        continue

    # -----------------------------------------------------------------------
    # STANDARD ROUTING: specific gen -> that tab, else General
    # -----------------------------------------------------------------------
    add_entry(primary_gen if primary_gen else "General", skill_key, entry)

print(f"\n📊 Rows processed : {ibis_count} IBIS rows")
print(f"📊 Rows skipped   : {skipped_count} non-IBIS rows (ignored)")

# ---------------------------------------------------------------------------
# SHARED STYLE
# ---------------------------------------------------------------------------
SHARED_STYLE = """
    body {
        font-family: Arial, sans-serif;
        font-size: 16px;
        color: #333;
        margin: 20px;
        background-color: #f0f0f0;
    }
    a { color: #08665c; text-decoration: none; }
    a:hover { color: #054a40; text-decoration: underline; }
"""

# --- 4. MAIN DASHBOARD (IBIS/IBIS.html) ---
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

    generation_html  = "<html><head><style>"
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
        sk = f"Skill Level {i}"
        if generation_skill_lists[generation][sk]["count"] > 0:
            has_data = True
            role      = SKILL_ROLE.get(i, "")
            docs_html = "<br/><br/>".join(generation_skill_lists[generation][sk]["list"])
            generation_html += (
                f"<tr>"
                f"<td class='level-cell'>Level {i} ({role})</td>"
                f"<td class='data-cell'>{docs_html}</td>"
                f"</tr>"
            )

    if not has_data:
        generation_html += (
            "<tr>"
            "<td class='level-cell' colspan='2'>"
            "No documents found for this generation."
            "</td>"
            "</tr>"
        )

    generation_html += "</table></body></html>"

    with open(os.path.join(nested_folder, f"{generation}.html"), "w", encoding="utf-8") as f:
        print(BeautifulSoup(generation_html, "html.parser").prettify(), file=f)

print("\n✅ Processing Complete.")
print("📁 Dashboard : IBIS/IBIS.html")
print("📁 Pages     : IBIS/IBIS/[Generation].html")
print("📋 Column header : Level/Role")
print("📋 Cell format   : Level 1 (Operator) / Level 3 (Technician)")
print("📋 Rule: OJTI + Gen 3 (SL3) -> Gen 3 tab only")
print("📋 Rule: All other SL3      -> ALL tabs")