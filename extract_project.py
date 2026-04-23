"""
extract_project.py
------------------
מייצר קובץ Word מתועד מתוך פרויקט Android.
סורק:
  - כל קבצי *.java
  - קבצי *.xml שנמצאים בתוך res/layout בלבד

שימוש:
    python extract_project.py <PATH_TO_APP_FOLDER>

דוגמא:
    python extract_project.py C:/MyProject/app
"""

import os
from docx import Document
from docx.shared import Pt

# ─────────────────────────────
# הגדרות – רק נתיב אחד!
# ─────────────────────────────
APP_DIR     = r"path/to/your/app"  # <-- עדכן לנתיב שלך 
OUTPUT_FILE = "ProjectBook.docx"

EXCLUDED_DIRS     = {"build", "generated", ".gradle", ".idea", "test", "tests", "androidTest"}
EXCLUDED_KEYWORDS = ["test", "Test", "Mock", "mock"]


def get_file_type(filename):
    if filename.endswith(".java"):
        return "JAVA"
    elif filename.endswith(".xml"):
        return "LAYOUT XML"
    return "CODE"


def is_layout_xml(root):
    folder = os.path.basename(root)
    return folder.startswith("layout")


def should_include(file, root):
    if file.endswith(".java"):
        if any(k in file for k in EXCLUDED_KEYWORDS):
            return False
        # רק תחת src/main
        if "src" + os.sep + "main" not in root:
            return False
        return True
    if file.endswith(".xml"):
        return is_layout_xml(root)
    return False


# ─────────────────────────────
document = Document()
document.add_heading("Project Code Documentation", 0)
document.add_paragraph(APP_DIR)
document.add_page_break()

seen = set()  # מניעת כפילויות

for root, dirs, files in os.walk(APP_DIR):
    # חיתוך תיקיות מוחרגות
    dirs[:] = [d for d in dirs if d not in EXCLUDED_DIRS]

    for file in sorted(files):
        if not should_include(file, root):
            continue
        if file in seen:
            print(f"⚠️  דילוג כפילות: {file}")
            continue
        seen.add(file)

        file_path = os.path.join(root, file)
        relative  = os.path.relpath(file_path, APP_DIR)
        file_type = get_file_type(file)

        print(f"  ➕ {file} [{file_type}]")

        heading = document.add_heading(level=1)
        heading.add_run(f"📄 {file}  [{file_type}]").bold = True
        document.add_paragraph("────────────────────────")

        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    p   = document.add_paragraph()
                    run = p.add_run(line.rstrip())
                    run.font.name = "Courier New"
                    run.font.size = Pt(9)
        except Exception as e:
            document.add_paragraph(f"שגיאה: {e}")

        document.add_paragraph()

document.save(OUTPUT_FILE)
print("\n✅ נוצר קובץ:", OUTPUT_FILE)