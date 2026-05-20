"""
Generate static JSON API files from Database TextBook.xlsx

Run this script whenever the xlsx is updated (e.g. adding Sekolah Rendah data).
Output goes to api/v1/

Usage:
    python generate_api.py
"""

import json
import os
import shutil
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("Installing openpyxl...")
    os.system("pip install openpyxl")
    import openpyxl


XLSX_PATH = Path(__file__).parent / "Database TextBook.xlsx"
API_DIR = Path(__file__).parent / "api" / "v1"

# Curriculum color mapping for book covers
CURRICULUM_COLORS = {
    "KSSM": "#4a90d9",
    "KBDKBT": "#e67e22",
    "KSSMPK": "#2ecc71",
    "MPAK": "#9b59b6",
    "MPEI": "#1abc9c",
    "MPET": "#e74c3c",
    "MPV": "#f39c12",
    "SSeM": "#e91e63",
    "KC": "#00bcd4",
    # Future SR curricula
    "KSSR": "#3498db",
    "KPM": "#8e44ad",
}

STUDENT_LEVEL_ORDER = {
    # Sekolah Rendah (future)
    "TA1": 1, "TA2": 2, "TA3": 3,
    # Peralihan (future)
    "PE": 7,
    # Sekolah Menengah — actual codes from xlsx
    "TB1": 10, "TB2": 11, "TB3": 12, "TA4": 13, "TA5": 14,
}


def read_xlsx():
    """Read the xlsx file and return a list of book dicts."""
    wb = openpyxl.load_workbook(str(XLSX_PATH), read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]

    books = []
    seen_ids = set()

    for row in rows[1:]:
        book = {}
        for i, header in enumerate(headers):
            val = row[i]
            if isinstance(val, str):
                val = val.strip()
            book[header] = val

        # Skip rows with missing essential fields
        if not book.get("id") or not book.get("schoolLevel") or not book.get("subjectName"):
            continue

        # Generate a unique ID if duplicates exist
        book_id = book["id"]
        if book_id in seen_ids:
            # Append a suffix for duplicate IDs
            suffix = 2
            while f"{book_id}_{suffix}" in seen_ids:
                suffix += 1
            book["id"] = f"{book_id}_{suffix}"

        seen_ids.add(book["id"])

        # Add download links placeholder
        book["downloads"] = {
            "telegram": "#",
            "pdf": "#",
            "gdrive": "#"
        }

        # Add cover color based on curriculum
        book["coverColor"] = CURRICULUM_COLORS.get(book.get("silibusCode", ""), "#666666")

        # Ensure dlpStatus is boolean
        book["dlpStatus"] = bool(book.get("dlpStatus", False))

        books.append(book)

    wb.close()
    return books


def ensure_dir(path):
    """Create directory if it doesn't exist."""
    Path(path).mkdir(parents=True, exist_ok=True)


def write_json(filepath, data):
    """Write JSON to file with pretty printing."""
    ensure_dir(os.path.dirname(filepath))
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  [OK] {filepath}")


def generate_api(books):
    """Generate all API endpoint files."""
    # Clean output directory
    if API_DIR.exists():
        shutil.rmtree(API_DIR)

    print(f"\nGenerating API for {len(books)} books...\n")

    # 1. All books
    write_json(API_DIR / "books.json", books)

    # 2. Individual book files
    books_dir = API_DIR / "books"
    for book in books:
        write_json(books_dir / f"{book['id']}.json", book)

    # 3. School levels with student levels
    levels_data = {}
    for book in books:
        sl = book["schoolLevel"]
        sl_name = book["schoolLevelName"]
        stl = book["studentLevel"]
        stl_name = book["studentLevelName"]

        if sl not in levels_data:
            levels_data[sl] = {
                "code": sl,
                "name": sl_name,
                "studentLevels": {}
            }

        if stl not in levels_data[sl]["studentLevels"]:
            levels_data[sl]["studentLevels"][stl] = {
                "code": stl,
                "name": stl_name,
                "bookCount": 0
            }

        levels_data[sl]["studentLevels"][stl]["bookCount"] += 1

    # Convert to sorted list
    levels_list = []
    for sl_code, sl_data in sorted(levels_data.items()):
        student_levels = sorted(
            sl_data["studentLevels"].values(),
            key=lambda x: STUDENT_LEVEL_ORDER.get(x["code"], 99)
        )
        levels_list.append({
            "code": sl_data["code"],
            "name": sl_data["name"],
            "studentLevels": student_levels,
            "totalBooks": sum(s["bookCount"] for s in student_levels)
        })

    write_json(API_DIR / "levels.json", levels_list)

    # 4. Books by school level and student level
    for book in books:
        sl = book["schoolLevel"]
        stl = book["studentLevel"]
        level_dir = API_DIR / "levels" / sl / stl
        ensure_dir(level_dir)

    # Group books by level
    level_books = {}
    for book in books:
        key = (book["schoolLevel"], book["studentLevel"])
        if key not in level_books:
            level_books[key] = []
        level_books[key].append(book)

    for (sl, stl), bks in level_books.items():
        write_json(API_DIR / "levels" / sl / stl / "books.json", bks)

    # 5. Curricula
    curricula_data = {}
    for book in books:
        code = book["silibusCode"]
        name = book["silibusName"]
        if code not in curricula_data:
            curricula_data[code] = {
                "code": code,
                "name": name,
                "color": CURRICULUM_COLORS.get(code, "#666666"),
                "bookCount": 0
            }
        curricula_data[code]["bookCount"] += 1

    curricula_list = sorted(curricula_data.values(), key=lambda x: x["code"])
    write_json(API_DIR / "curricula.json", curricula_list)

    # 6. Books by curriculum
    curriculum_books = {}
    for book in books:
        code = book["silibusCode"]
        if code not in curriculum_books:
            curriculum_books[code] = []
        curriculum_books[code].append(book)

    for code, bks in curriculum_books.items():
        write_json(API_DIR / "curricula" / code / "books.json", bks)

    # 7. Search index (lightweight)
    search_index = []
    for book in books:
        search_index.append({
            "id": book["id"],
            "subjectName": book["subjectName"],
            "displayName": book["displayName"],
            "studentLevelName": book["studentLevelName"],
            "silibusCode": book["silibusCode"],
            "schoolLevel": book["schoolLevel"],
            "dlpStatus": book["dlpStatus"],
        })

    write_json(API_DIR / "search-index.json", search_index)

    print(f"\n[DONE] API generated successfully!")
    print(f"   Total books: {len(books)}")
    print(f"   Individual book files: {len(books)}")
    print(f"   School levels: {len(levels_list)}")
    print(f"   Student levels: {sum(len(l['studentLevels']) for l in levels_list)}")
    print(f"   Curricula: {len(curricula_list)}")


if __name__ == "__main__":
    books = read_xlsx()
    generate_api(books)
