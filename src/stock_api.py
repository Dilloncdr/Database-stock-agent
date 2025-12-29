from __future__ import annotations

import difflib
import json
import os
import re
import sqlite3
from pathlib import Path
from typing import Any, Dict, List, Optional, Literal, Union, Tuple

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from rapidfuzz import fuzz


# ==========================================================
# BRAND ALIASES
# ==========================================================

ALIASES_PATH = Path(
    os.getenv(
        "BRAND_ALIASES_PATH",
        r"C:\Users\Curve System\Desktop\KsyncwAI\dbmaker\brand_aliases.json",
    )
)

BRAND_ALIASES: dict[str, list[str]] = {}
if ALIASES_PATH.exists():
    with open(ALIASES_PATH, "r", encoding="utf-8") as f:
        BRAND_ALIASES = json.load(f)


# ==========================================================
# DB CONFIG
# ==========================================================

DB_PATH = Path(
    os.getenv("STOCK_DB_PATH", r"C:\Users\Curve System\Desktop\KsyncwAI\Excels\InStock.db")
)
TABLE_NAME = "products"

COL_NAME = "نام كتاب"
COL_QTY = "تعداد"
COL_PRICE = "پشت‌جلد"  # contains ZWNJ in column name, keep EXACT
COL_CAT = "مشخصه 4"
COL_AUTHOR = "مولف"
COL_TRANSLATOR = "مترجم"
COL_PUBLISHER = "ناشر"
COL_GROUP = "گروه"
COL_SYS = "كدسيستم"
COL_GROUPFAMILY = "GroupFamily"


# ==========================================================
# NORMALIZATION HELPERS
# ==========================================================

_price_digits = re.compile(r"[^\d]+")


def normalize_persian(text: Any) -> str:
    if text is None:
        return ""
    s = str(text)

    # Arabic -> Persian + cleanup
    mapping = {
        "ي": "ی",
        "ك": "ک",
        "\u200c": " ",  # ZWNJ -> space
        "‌": " ",       # sometimes appears as literal
        "\u0640": "",   # tatweel
    }
    for a, b in mapping.items():
        s = s.replace(a, b)

    # collapse spaces
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_query(text: str) -> str:
    if not text:
        return ""
    text = text.strip()
    # remove question marks
    text = re.sub(r"[؟?]", "", text)
    # normalize spaces
    text = re.sub(r"\s+", " ", text)
    return normalize_persian(text)


def parse_price(value: Any) -> Optional[int]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        try:
            return int(value)
        except Exception:
            return None
    s = str(value)
    s = s.replace(",", "").replace("٬", "").replace(" ", "").strip()
    s = _price_digits.sub("", s)
    if not s:
        return None
    try:
        return int(s)
    except Exception:
        return None


def like_pattern(term: str) -> str:
    return f"%{term}%"


def similarity(a: str, b: str) -> float:
    # difflib is fine for fallback; we keep it deterministic and dependency-free
    return difflib.SequenceMatcher(None, normalize_persian(a), normalize_persian(b)).ratio()


# ==========================================================
# Pydantic Models (SearchIntent Contract)
# ==========================================================

class TextFilter(BaseModel):
    include: List[str] = Field(default_factory=list)
    exclude: List[str] = Field(default_factory=list)


class GroupFamilyTags(BaseModel):
    include_all: List[str] = Field(default_factory=list)  # AND
    include_any: List[str] = Field(default_factory=list)  # OR
    exclude: List[str] = Field(default_factory=list)


class PriceFilter(BaseModel):
    min: Optional[int] = None
    max: Optional[int] = None


class StockFilter(BaseModel):
    in_stock_only: Optional[bool] = None


class Filters(BaseModel):
    name: TextFilter = Field(default_factory=TextFilter)
    author_or_type_or_age: TextFilter = Field(default_factory=TextFilter)
    translator_or_playtime: TextFilter = Field(default_factory=TextFilter)
    publisher_or_brand: TextFilter = Field(default_factory=TextFilter)
    group_main: TextFilter = Field(default_factory=TextFilter)
    groupfamily_tags: GroupFamilyTags = Field(default_factory=GroupFamilyTags)
    price: PriceFilter = Field(default_factory=PriceFilter)
    stock: StockFilter = Field(default_factory=StockFilter)


class SortSpec(BaseModel):
    by: Literal["relevance", "price", "qty"] = "relevance"
    direction: Literal["asc", "desc"] = "desc"


class UISpec(BaseModel):
    need_clarification: bool = False
    clarification_question: Optional[str] = None


class SearchIntent(BaseModel):
    version: str = "1.0"

    # allow single or multi-category
    category_code: Optional[Union[str, List[str]]] = None

    uploaded_only: bool = True
    query_text: str = ""

    filters: Filters = Field(default_factory=Filters)
    sort: SortSpec = Field(default_factory=SortSpec)
    limit: int = 10
    ui: UISpec = Field(default_factory=UISpec)

    debug: bool = False


# ==========================================================
# DB ACCESS
# ==========================================================

def get_connection() -> sqlite3.Connection:
    if not DB_PATH.exists():
        raise RuntimeError(f"DB file not found: {DB_PATH}")
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


# ==========================================================
# SQL BUILDER (Deterministic)
# ==========================================================

def build_sql(intent: SearchIntent) -> Tuple[str, List[Any]]:
    where: List[str] = []
    params: List[Any] = []

    # uploaded_only => GroupFamily exists & not empty
    if intent.uploaded_only:
        where.append(f'"{COL_GROUPFAMILY}" IS NOT NULL')
        where.append(f'TRIM("{COL_GROUPFAMILY}") != ""')

    # category_code supports string or list
    if intent.category_code:
        if isinstance(intent.category_code, list):
            cats = [normalize_persian(c).strip().lower() for c in intent.category_code if str(c).strip()]
        else:
            cats = [normalize_persian(intent.category_code).strip().lower()]

        # optional auto-expand: if user sent only "s" but query implies luxury pen too
        pen_keywords = ["خودکار", "خودنویس", "روان نویس", "هدیه", "قلم", "نفیس", "لاکچری"]
        q = normalize_persian(intent.query_text or "")
        if cats == ["s"] and any(k in q for k in pen_keywords):
            cats = ["s", "l"]

        if len(cats) == 1:
            where.append(f'"{COL_CAT}" = ?')
            params.append(cats[0])
        else:
            placeholders = ", ".join(["?"] * len(cats))
            where.append(f'"{COL_CAT}" IN ({placeholders})')
            params.extend(cats)

    def expand_brand_terms(user_term: str) -> List[str]:
        """
        Expand brand aliases in either direction:
        - if user typed alias ("فابر") => include canonical ("faber castell") too
        - if user typed canonical => include aliases too
        """
        t_norm = normalize_persian(user_term).lower()
        for canonical, aliases in BRAND_ALIASES.items():
            all_terms = [canonical] + aliases
            all_norm = [normalize_persian(x).lower() for x in all_terms]
            if t_norm in all_norm:
                return all_terms
        return [user_term]

    def apply_text_filter(col: str, tf: TextFilter):
        # include
        for t in tf.include:
            t = normalize_persian(t)
            if not t:
                continue

            terms = [t]
            if col == COL_PUBLISHER and BRAND_ALIASES:
                terms = expand_brand_terms(t)

            or_block = []
            for term in terms:
                term_n = normalize_persian(term)
                or_block.append(f'"{col}" LIKE ?')
                params.append(like_pattern(term_n))
            where.append("(" + " OR ".join(or_block) + ")")

        # exclude
        for t in tf.exclude:
            t = normalize_persian(t)
            if t:
                where.append(f'("{col}" IS NULL OR "{col}" NOT LIKE ?)')
                params.append(like_pattern(t))

    apply_text_filter(COL_NAME, intent.filters.name)
    apply_text_filter(COL_AUTHOR, intent.filters.author_or_type_or_age)
    apply_text_filter(COL_TRANSLATOR, intent.filters.translator_or_playtime)
    apply_text_filter(COL_PUBLISHER, intent.filters.publisher_or_brand)
    apply_text_filter(COL_GROUP, intent.filters.group_main)

    gf = intent.filters.groupfamily_tags

    # AND
    for t in gf.include_all:
        t = normalize_persian(t)
        if t:
            where.append(f'"{COL_GROUPFAMILY}" LIKE ?')
            params.append(like_pattern(t))

    # OR
    any_terms = [normalize_persian(t) for t in gf.include_any if normalize_persian(t)]
    if any_terms:
        or_block = []
        for t in any_terms:
            or_block.append(f'"{COL_GROUPFAMILY}" LIKE ?')
            params.append(like_pattern(t))
        where.append("(" + " OR ".join(or_block) + ")")

    # exclude
    for t in gf.exclude:
        t = normalize_persian(t)
        if t:
            where.append(f'("{COL_GROUPFAMILY}" IS NULL OR "{COL_GROUPFAMILY}" NOT LIKE ?)')
            params.append(like_pattern(t))

    # query_text fallback only if user did NOT provide explicit filters
    user_filters_present = (
        intent.category_code is not None
        or bool(intent.filters.name.include or intent.filters.name.exclude)
        or bool(intent.filters.author_or_type_or_age.include or intent.filters.author_or_type_or_age.exclude)
        or bool(intent.filters.translator_or_playtime.include or intent.filters.translator_or_playtime.exclude)
        or bool(intent.filters.publisher_or_brand.include or intent.filters.publisher_or_brand.exclude)
        or bool(intent.filters.group_main.include or intent.filters.group_main.exclude)
        or bool(gf.include_all or gf.include_any or gf.exclude)
    )

    qtext = normalize_query(intent.query_text)
    if qtext and not user_filters_present:
        or_cols = [COL_NAME, COL_AUTHOR, COL_TRANSLATOR, COL_PUBLISHER, COL_GROUP, COL_GROUPFAMILY]
        or_block = []
        for c in or_cols:
            or_block.append(f'"{c}" LIKE ?')
            params.append(like_pattern(qtext))
        where.append("(" + " OR ".join(or_block) + ")")

    where_sql = " AND ".join(where) if where else "1=1"

    sql = f"""
        SELECT
            "{COL_NAME}"        AS name,
            "{COL_QTY}"         AS qty,
            "{COL_PRICE}"       AS price,
            "{COL_CAT}"         AS category_code,
            "{COL_AUTHOR}"      AS author_or_type_or_age,
            "{COL_TRANSLATOR}"  AS translator_or_playtime,
            "{COL_PUBLISHER}"   AS publisher_or_brand,
            "{COL_GROUP}"       AS group_main,
            "{COL_SYS}"         AS system_code,
            "{COL_GROUPFAMILY}" AS groupfamily
        FROM {TABLE_NAME}
        WHERE {where_sql}
        LIMIT 1200
    """
    return sql, params


# ==========================================================
# RANKING + POST-FILTER
# ==========================================================

def compute_relevance_score(intent: SearchIntent, row: Dict[str, Any]) -> int:
    q = normalize_query(intent.query_text or "")
    if not q:
        return 0

    name = normalize_persian(row.get("name", ""))
    author = normalize_persian(row.get("author_or_type_or_age", ""))
    translator = normalize_persian(row.get("translator_or_playtime", ""))
    pub = normalize_persian(row.get("publisher_or_brand", ""))
    group = normalize_persian(row.get("group_main", ""))
    gf = normalize_persian(row.get("groupfamily", ""))

    s1 = fuzz.partial_ratio(q, name)
    s2 = fuzz.partial_ratio(q, author)
    s3 = fuzz.partial_ratio(q, translator)
    s4 = fuzz.partial_ratio(q, pub)
    s5 = fuzz.partial_ratio(q, gf)
    s6 = fuzz.partial_ratio(q, group)

    # name/groupfamily strongest
    return int(max(
        0.95 * s1,
        0.90 * s5,
        0.75 * s2,
        0.65 * s3,
        0.55 * s4,
        0.50 * s6,
    ))


def passes_price_filter(intent: SearchIntent, row: Dict[str, Any]) -> bool:
    pmin = intent.filters.price.min
    pmax = intent.filters.price.max
    if pmin is None and pmax is None:
        return True
    price_val = parse_price(row.get("price"))
    if price_val is None:
        return False
    if pmin is not None and price_val < pmin:
        return False
    if pmax is not None and price_val > pmax:
        return False
    return True


# ==========================================================
# FASTAPI APP
# ==========================================================

app = FastAPI(
    title="Pasdaran Book City Stock API",
    version="3.1.0",
    description="Deterministic DB search engine for InStock.db using SearchIntent contract",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/ping")
def ping():
    return {
        "status": "ok",
        "db_exists": DB_PATH.exists(),
        "db_path": str(DB_PATH),
        "aliases_loaded": bool(BRAND_ALIASES),
        "aliases_path": str(ALIASES_PATH),
        "table": TABLE_NAME,
    }


@app.post("/search")
def search(intent: SearchIntent) -> Dict[str, Any]:
    # normalize query_text for punctuation, arabic variants, spaces
    intent.query_text = normalize_query(intent.query_text or "")

    sql, params = build_sql(intent)
    limit_final = max(1, min(intent.limit, 50))

    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(sql, params)
        rows = cur.fetchall()

        # --------------------------------------------------
        # FUZZY NAME FALLBACK (typo tolerance)
        # Only if strict SQL returns nothing AND we have a meaningful "name intent"
        # --------------------------------------------------
        name_seed = ""
        if intent.filters.name.include:
            name_seed = normalize_query(intent.filters.name.include[0])
        elif intent.query_text:
            # if user typed a specific title without explicit filters, still allow fallback
            name_seed = intent.query_text

        if not rows and name_seed:
            # fetch a wider candidate set (still respecting uploaded_only)
            where = []
            p = []
            if intent.uploaded_only:
                where.append(f'"{COL_GROUPFAMILY}" IS NOT NULL')
                where.append(f'TRIM("{COL_GROUPFAMILY}") != ""')

            # if category is provided, keep it in candidate fetch too (string or list)
            if intent.category_code:
                if isinstance(intent.category_code, list):
                    cats = [normalize_persian(c).strip().lower() for c in intent.category_code if str(c).strip()]
                else:
                    cats = [normalize_persian(intent.category_code).strip().lower()]
                if cats:
                    placeholders = ", ".join(["?"] * len(cats))
                    where.append(f'"{COL_CAT}" IN ({placeholders})')
                    p.extend(cats)

            where_sql = " AND ".join(where) if where else "1=1"

            fallback_sql = f"""
                SELECT
                    "{COL_NAME}"        AS name,
                    "{COL_QTY}"         AS qty,
                    "{COL_PRICE}"       AS price,
                    "{COL_CAT}"         AS category_code,
                    "{COL_AUTHOR}"      AS author_or_type_or_age,
                    "{COL_TRANSLATOR}"  AS translator_or_playtime,
                    "{COL_PUBLISHER}"   AS publisher_or_brand,
                    "{COL_GROUP}"       AS group_main,
                    "{COL_SYS}"         AS system_code,
                    "{COL_GROUPFAMILY}" AS groupfamily
                FROM {TABLE_NAME}
                WHERE {where_sql}
                LIMIT 2500
            """
            cur.execute(fallback_sql, p)
            candidates = cur.fetchall()

            # fuzzy filter on name
            fuzzy_matches: List[Tuple[float, sqlite3.Row]] = []
            for r in candidates:
                pname = normalize_query(r["name"] or "")
                if not pname:
                    continue
                score = similarity(name_seed, pname)
                if score >= 0.72:
                    fuzzy_matches.append((score, r))

            fuzzy_matches.sort(key=lambda x: x[0], reverse=True)
            rows = [r for _, r in fuzzy_matches[:1200]]

    finally:
        conn.close()

    # Convert to dict + normalize display fields
    items: List[Dict[str, Any]] = []
    category_counts: Dict[str, int] = {}

    for r in rows:
        d = dict(r)

        d["name"] = normalize_persian(d.get("name"))
        d["author_or_type_or_age"] = normalize_persian(d.get("author_or_type_or_age"))
        d["translator_or_playtime"] = normalize_persian(d.get("translator_or_playtime"))
        d["publisher_or_brand"] = normalize_persian(d.get("publisher_or_brand"))
        d["group_main"] = normalize_persian(d.get("group_main"))
        d["groupfamily"] = normalize_persian(d.get("groupfamily"))
        d["category_code"] = normalize_persian(d.get("category_code")).strip().lower()

        d["price_value"] = parse_price(d.get("price"))

        try:
            d["qty_value"] = int(d.get("qty")) if d.get("qty") is not None else None
        except Exception:
            d["qty_value"] = None

        # Price filter
        if not passes_price_filter(intent, d):
            continue

        cc = d.get("category_code") or ""
        category_counts[cc] = category_counts.get(cc, 0) + 1

        d["score"] = compute_relevance_score(intent, d)

        items.append(d)

    # Sorting
    reverse = (intent.sort.direction == "desc")
    if intent.sort.by == "price":
        items.sort(
            key=lambda x: (x.get("price_value") is None, x.get("price_value", 10**18)),
            reverse=reverse,
        )
    elif intent.sort.by == "qty":
        items.sort(
            key=lambda x: (x.get("qty_value") is None, x.get("qty_value", -1)),
            reverse=reverse,
        )
    else:
        items.sort(key=lambda x: x.get("score", 0), reverse=True)

    items = items[:limit_final]

    response = {
        "version": intent.version,
        "query_text": intent.query_text,
        "category_code_used": intent.category_code,
        "uploaded_only": intent.uploaded_only,
        "count": len(items),
        "category_distribution": category_counts,
        "results": [
            {
                "name": it["name"],
                "qty": it.get("qty_value"),
                "price": it.get("price_value"),
                "category_code": it.get("category_code"),
                "author_or_type_or_age": it.get("author_or_type_or_age"),
                "translator_or_playtime": it.get("translator_or_playtime"),
                "publisher_or_brand": it.get("publisher_or_brand"),
                "group_main": it.get("group_main"),
                "groupfamily": it.get("groupfamily"),
                "system_code": it.get("system_code"),
                "score": it.get("score", 0),
            }
            for it in items
        ],
    }

    if intent.debug:
        response["debug"] = {
            "sql": sql.strip(),
            "params": params,
            "db_path": str(DB_PATH),
            "aliases_path": str(ALIASES_PATH),
            "aliases_loaded": bool(BRAND_ALIASES),
        }

    return response
