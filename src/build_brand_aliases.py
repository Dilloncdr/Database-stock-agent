import sqlite3
import json
from pathlib import Path

DB_PATH = Path(r"C:\Users\Curve System\Desktop\KsyncwAI\Excels\InStock.db")
OUT_FILE = Path(r"C:\Users\Curve System\Desktop\KsyncwAI\dbmaker\brand_aliases.json")

def normalize(s):
    return s.lower().strip()

# very lightweight phonetic rules (enough for fuzzy search)
LETTER_MAP = {
    "ph": "ف",
    "f": "ف",
    "b": "ب",
    "p": "پ",
    "t": "ت",
    "d": "د",
    "k": "ک",
    "g": "گ",
    "s": "س",
    "sh": "ش",
    "ch": "چ",
    "j": "ج",
    "l": "ل",
    "m": "م",
    "n": "ن",
    "r": "ر",
    "v": "و",
    "w": "و",
    "y": "ی",
    "i": "ی",
    "e": "e",
    "a": "ا",
    "o": "و",
    "u": "و",
}

def phonetic_fa(word):
    w = normalize(word)
    out = w
    for en, fa in LETTER_MAP.items():
        out = out.replace(en, fa)
    return out

def main():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
        SELECT DISTINCT ناشر
        FROM products
        WHERE "مشخصه 4" = 's'
          AND ناشر IS NOT NULL
    """)
    brands = [r[0] for r in cur.fetchall()]
    conn.close()

    aliases = {}

    for brand in brands:
        b = normalize(brand)

        # skip pure Persian brands
        if any("\u0600" <= ch <= "\u06FF" for ch in b):
            continue

        fa_guess = phonetic_fa(b)
        parts = b.split()
        part_aliases = [phonetic_fa(p) for p in parts if len(p) > 2]

        aliases[b] = list(set([fa_guess] + part_aliases))

    OUT_FILE.write_text(json.dumps(aliases, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Generated aliases for {len(aliases)} stationery brands.")
    print(f"Saved to: {OUT_FILE}")

if __name__ == "__main__":
    main()
