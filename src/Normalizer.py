from pathlib import Path
import pandas as pd
import sqlite3


# # =========================
# # CONFIG
# # =========================

# FILE_DIR = Path(r"C:\Users\Curve System\Desktop\KsyncwAI\Excels")

# RAW_FILE = FILE_DIR / "KowsarExport.xlsx"
# CLEAN_FILE = FILE_DIR / "InStockClean.xlsx"
# DB_FILE = FILE_DIR / "InStock.db"

# # COLUMNS_TO_NORMALIZE = [
# #     "نام كتاب",
# #     "ناشر",
# #     "مولف",
# #     "مترجم",
# #     "گروه",
# #     "GroupFamily",
# # ]


# # =========================
# # NORMALIZATION FUNCTION
# # =========================

# # def normalize_persian(text: str) -> str:
# #     if not isinstance(text, str):
# #         return text

# #     replacements = {
# #         "\u064a": "ی",   # ي -> ی
# #         "\u0649": "ی",   # ى -> ی
# #         "\u0626": "ی",   # ئ -> ی
# #         "\u0643": "ک",   # ك -> ک
# #         "\u0629": "ه",   # ة -> ه
# #         "\u200c": "",    # ZWNJ
# #         "\u0640": "",    # Kashida (ـ)
# #     }

# #     for src, dst in replacements.items():
# #         text = text.replace(src, dst)

# #     return text.strip()


# # =========================
# # MAIN NORMALIZER
# # =========================

# # def build_clean_excel_db():
# #     if not RAW_FILE.exists():
# #         print(f"[ERROR] Raw export not found: {RAW_FILE}")
# #         return

# #     print("[INFO] Loading raw export...")
# #     df = pd.read_excel(RAW_FILE)

# #     print("[INFO] Normalizing text columns...")
# #     for col in COLUMNS_TO_NORMALIZE:
# #         if col in df.columns:
# #             df[col] = df[col].astype(str).map(normalize_persian)
# #             print(f"  ✔ Normalized: {col}")
# #         else:
# #             print(f"  ⚠ Column not found, skipped: {col}")

# #     print("[INFO] Saving clean database...")
# #     df.to_excel(CLEAN_FILE, index=False)

# #     print(f"[SUCCESS] Clean DB updated: {CLEAN_FILE}")

# def excel_to_sqlite(CLEAN_FILE: Path, db_path: Path):
#     """
#     Convert the cleaned Excel file into a SQLite database.
#     Table name: products
#     """
#     if not CLEAN_FILE.exists():
#         print(f"Clean Excel not found, skipping SQLite update: {CLEAN_FILE}")
#         return

#     print(f"Loading clean Excel: {CLEAN_FILE}")
#     df = pd.read_excel(CLEAN_FILE)

#     # You can drop unused columns here if you want:
#     # df = df[["نام كتاب", "كد كتاب", "تعداد", "قيمت", ...]]

#     print(f"Writing {len(df)} rows to SQLite DB: {DB_FILE}")
#     conn = sqlite3.connect(DB_FILE)
#     df.to_sql("products", conn, if_exists="replace", index=False)
#     conn.close()
#     print(f"SQLite DB updated: {DB_FILE}")
# # =========================
# # RUN (ONE-TIME OR IMPORT)
# # =========================

# if __name__ == "__main__":
#     excel_to_sqlite()
# from pathlib import Path
# import pandas as pd
# import sqlite3


# # =========================
# # CONFIG
# # =========================

# FILE_DIR = Path(r"C:\Users\Curve System\Desktop\KsyncwAI\Excels")

# CLEAN_FILE = FILE_DIR / "InStockClean.xlsx"
# DB_FILE = FILE_DIR / "stock.db"

# TABLE_NAME = "products"


# # =========================
# # EXCEL → SQLITE
# # =========================

# def excel_to_sqlite():
#     if not CLEAN_FILE.exists():
#         print(f"[ERROR] Clean Excel not found: {CLEAN_FILE}")
#         return

#     print("[INFO] Loading clean Excel file...")
#     df = pd.read_excel(CLEAN_FILE)

#     print("[INFO] Connecting to SQLite DB...")
#     conn = sqlite3.connect(DB_FILE)

#     print("[INFO] Writing data to SQLite...")
#     df.to_sql(
#         TABLE_NAME,
#         conn,
#         if_exists="replace",   # fully refresh DB each run
#         index=False
#     )

#     conn.close()

#     print(f"[SUCCESS] SQLite DB updated: {DB_FILE}")
#     print(f"[SUCCESS] Table name: {TABLE_NAME}")


# # =========================
# # RUN
# # =========================

# if __name__ == "__main__":
#     excel_to_sqlite()
from pathlib import Path
import pandas as pd
import sqlite3
import re


# =========================
# CONFIG
# =========================

BASE_DIR = Path(r"C:\Users\Curve System\Desktop\KsyncwAI\Excels")

EXCEL_FILE = BASE_DIR / "InStockClean.xlsx"
DB_FILE = BASE_DIR / "stock.db"

TABLE_NAME = "products"

CATEGORY_COLUMNS = [
    "گروه",
    "GroupFamily",
]


# =========================
# NORMALIZATION
# =========================

def normalize_persian(text: str) -> str:
    if not isinstance(text, str):
        return text

    replacements = {
        "\u064a": "ی",   # ي → ی
        "\u0649": "ی",   # ى → ی
        "\u0626": "ی",   # ئ → ی
        "\u0643": "ک",   # ك → ک
        "\u0629": "ه",   # ة → ه
        "\u200c": "",    # ZWNJ
        "\u0640": "",    # Kashida
    }

    for src, dst in replacements.items():
        text = text.replace(src, dst)

    return text.strip()


def clean_category_text(text: str) -> str:
    if not isinstance(text, str):
        return text

    text = normalize_persian(text)

    # remove (123)
    text = re.sub(r"\s*\(\d+\)\s*", "", text)

    # Persian comma → English comma
    text = text.replace("،", ",")

    # normalize spacing
    text = re.sub(r"\s*,\s*", ", ", text)

    # remove duplicate commas
    text = re.sub(r"(,\s*){2,}", ", ", text)

    return text.strip(" ,")


# =========================
# MAIN (ONE TIME)
# =========================

def excel_to_sqlite_one_time():
    if not EXCEL_FILE.exists():
        print(f"[ERROR] Excel file not found: {EXCEL_FILE}")
        return

    print("[INFO] Loading Excel...")
    df = pd.read_excel(EXCEL_FILE)

    print("[INFO] Cleaning category columns...")
    for col in CATEGORY_COLUMNS:
        if col in df.columns:
            df[col] = df[col].astype(str).map(clean_category_text)
            print(f"  ✔ Cleaned: {col}")
        else:
            print(f"  ⚠ Column not found, skipped: {col}")

    print("[INFO] Writing to SQLite...")
    conn = sqlite3.connect(DB_FILE)
    df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)
    conn.close()

    print("[SUCCESS] Done.")
    print(f"SQLite DB: {DB_FILE}")
    print(f"Table: {TABLE_NAME}")


# =========================
# RUN ONCE
# =========================

if __name__ == "__main__":
    excel_to_sqlite_one_time()
