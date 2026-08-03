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

# List of equipment names
equipment_names = ["KLA", "SRM", "ISMECA", "TTM", "ETM", "Peel Force Tester"]

# Initialize dictionary
equipment_skill_lists = {
    equipment: {f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)}
    for equipment in equipment_names
}

# NEW: Skill Level -> Role mapping
SKILL_ROLE = {
    1: "Operator",
    2: "Operator",
    3: "Technician",
    4: "Technician",
    5: "Technician",
    6: "Technician",
}

def level_role_label(i: int) -> str:
    """Return 'Level i (Role)' string."""
    return f"Level {i} ({SKILL_ROLE.get(i, '')})"

# Single source of truth for the table header
LEVEL_HEADER = "<tr class='header-row'><th>Level/Role</th><th>Documents</th></tr>"

print("🔄 Processing data with Whole Word Logic...")
total_processed = 0

for row in data:
    if len(row) < 2:
        continue

    value = str(row[0]) if pd.notna(row[0]) else ""
    link = str(row[1]) if pd.notna(row[1]) else ""

    if not value.strip():
        continue

    value_lower = value.lower()

    for equipment in equipment_names:
        pattern = r'\b' + re.escape(equipment.lower()) + r'\b'
        if re.search(pattern, value_lower):
            for i in range(1, 7):
                if f"(Skill Level {i})" in value:
                    entry = (
                        f"{value} - "
                        f"<a href='https://plmpublishing.icp.infineon.com/api/download-pdf/{link}'"
                        f" target='_blank'>{link}</a>"
                    )
                    equipment_skill_lists[equipment][f"Skill Level {i}"]["list"].append(entry)
                    equipment_skill_lists[equipment][f"Skill Level {i}"]["count"] += 1
                    total_processed += 1

print(f"✅ Total entries processed: {total_processed}")

# --- 3. MAIN DASHBOARD (MSP/MSP.html) ---
html = "<html><head>"
html += "<style>"
html += """
body {
  font-family: Arial, sans-serif;
  margin: 20px;
  background-color: #f0f0f0;
}
.container {
  width: 80%;
  margin: 0 auto;
  background-color: #fff;
  padding: 20px;
  border: 1px solid #ddd;
  border-radius: 10px;
  box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
}
.header {
  background-color: #f0f0f0;
  padding: 10px;
  border-bottom: 1px solid #ddd;
  display: flex;
  justify-content: flex-start;
  align-items: center;
}
.home-btn {
  display: inline-block;
  padding: 8px 16px;
  background-color: #08665c;
  color: white;
  text-decoration: none;
  border-radius: 5px;
  font-size: 14px;
  transition: background-color 0.3s;
}
.home-btn:hover {
  background-color: #054a40;
}
.content {
  padding: 20px;
  display: flex;
}
.sidebar {
  width: 20%;
  background-color: #08665c;
  color: #fff;
  padding: 20px;
  font-size: 24px;
  text-align: center;
  border-radius: 10px 0 0 10px;
  display: flex;
  justify-content: center;
  align-items: center;
}
.main-content {
  width: 80%;
  padding: 20px;
}
.post-it-container {
  display: flex;
  flex-wrap: wrap;
  justify-content: center;
}
.post-it {
  width: 150px;
  height: 150px;
  background-color: #ADD8E6;
  padding: 10px;
  border: 1px solid #ccc;
  border-radius: 10px;
  margin: 10px;
  display: flex;
  justify-content: center;
  align-items: center;
  cursor: pointer;
  transition: background-color 0.2s ease-in-out;
}
.post-it:hover {
  background-color: #87CEEB;
}
.post-it a {
  text-decoration: none;
  color: #000;
  width: 100%;
  height: 100%;
  display: flex;
  justify-content: center;
  align-items: center;
  transition: color 0.2s ease-in-out;
  font-weight: normal;
}
.post-it:hover a {
  color: #fff;
}
"""
html += "</style></head><body>"
html += "<div class='container'>"
html += "<div class='header'>"
html += "<a href='../index.html' class='home-btn'>Return to Home</a>"
html += "</div>"
html += "<div class='content'>"
html += "<div class='sidebar'>MSP</div>"
html += "<div class='main-content'>"
html += "<div class='post-it-container'>"

for equipment in equipment_names:
    safe_filename = equipment.replace(" ", "_") + ".html"
    html += f"<div class='post-it'><a href='MSP/{safe_filename}'>{equipment}</a></div>"

html += "</div>"
html += "</div>"
html += "</div>"
html += "</div>"
html += "</body></html>"

soup = BeautifulSoup(html, 'html.parser')
with open(os.path.join(main_folder, "MSP.html"), "w", encoding="utf-8") as file:
    print(soup.prettify(), file=file)

# --- 4. INDIVIDUAL EQUIPMENT PAGES ---
for equipment in equipment_names:
    safe_filename = equipment.replace(" ", "_") + ".html"
    equipment_html = "<html><head>"
    equipment_html += "<style>"
    equipment_html += """
    body {
      font-family: Arial, sans-serif;
      margin: 20px;
      background-color: #f0f0f0;
    }
    .back-btn {
      display: inline-block;
      padding: 10px 20px;
      background-color: #08665c;
      color: white;
      text-decoration: none;
      border-radius: 5px;
      font-size: 14px;
      margin-bottom: 20px;
      transition: background-color 0.3s;
    }
    .back-btn:hover {
      background-color: #054a40;
    }
    .data-table {
      border-collapse: collapse;
      width: 100%;
      font-size: 18px;
      box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
      background-color: #fff;
    }
    .header-row {
      background-color: #f0f0f0;
      color: #333;
      font-weight: bold;
    }
    .header-row th {
      padding: 12px;
      text-align: left;
    }
    .skill-level {
      text-align: left;
      font-size: 16px;
      padding: 12px;
      border-bottom: 1px solid #ddd;
    }
    .data-cell {
      text-align: left;
      font-size: 16px;
      padding: 12px;
      border-bottom: 1px solid #ddd;
      line-height: 1.5;
    }
    .data-cell.empty {
      background-color: #cccccc;
    }
    .data-cell:hover {
      background-color: #f0f0f0;
    }
    """
    equipment_html += "</style></head><body>"
    equipment_html += "<a href='../MSP.html' class='back-btn'>← Back to Dashboard</a>"
    equipment_html += f"<h1>{equipment}</h1>"
    equipment_html += "<table class='data-table'>"

    # Header is now Level/Role
    equipment_html += LEVEL_HEADER

    has_data = False
    for i in range(1, 7):
        if equipment_skill_lists[equipment][f"Skill Level {i}"]["count"] > 0:
            has_data = True
            content = '<br/><br/>'.join(equipment_skill_lists[equipment][f"Skill Level {i}"]["list"])
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
        print(BeautifulSoup(equipment_html, 'html.parser').prettify(), file=file)

print("\n✅ Structure Updated to Nested Folders:")
print("📁 Dashboard : MSP/MSP.html")
print("📁 Pages     : MSP/MSP/[Equipment].html")
print("📋 Column header : Level/Role")
print("📋 Cell format   : Level 1 (Operator) / Level 3 (Technician)")