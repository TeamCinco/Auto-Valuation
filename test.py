"""
dera_fetcher.py
Uses SEC DERA Financial Statement Data Sets as an enhancement layer
for the XBRL parser. NOT a replacement — a lookup/validation tool.

Three use cases:
  1. Label resolver: concept "CashAndCashEquivalentsAtCarryingValue" →
     "Cash and cash equivalents"
  2. Statement classifier: when classify_role() fails, DERA's pre.txt
     tells you exactly which statement (BS/IS/CF) a concept belongs to
  3. Line ordering: gives you the SEC's rendered line order so you don't
     need to DFS-walk the presentation linkbase

Data source: https://www.sec.gov/files/dera/data/financial-statement-data-sets/{year}q{quarter}.zip
Each zip contains tab-delimited txt files: sub.txt, num.txt, pre.txt, tag.txt

Coverage: 2009q1 through present. Updated quarterly.

Schema (from readme.htm):
  sub.txt: adsh, cik, name, form, period, fy, fp, filed, instance, ...
  num.txt: adsh, tag, version, ddate, qtrs, uom, coreg, value, footnote
  pre.txt: adsh, report, line, stmt, inpth, tag, version, plabel, negating
  tag.txt: tag, version, custom, abstract, datatype, iord, crdr, tlabel, doc
"""
import io
import logging
import zipfile
from pathlib import Path

import requests
import pandas as pd

logger = logging.getLogger(__name__)

HEADERS = {"User-Agent": "SEC Filing App contact@example.com"}
BASE_URL = "https://www.sec.gov/files/dera/data/financial-statement-data-sets"

# Statement code mapping
STMT_MAP = {
    "BS": "Balance Sheet",
    "IS": "Income Statement",
    "CF": "Cash Flow",
    "EQ": "Stockholders Equity",
    "CI": "Comprehensive Income",
}


# ═══════════════════════════════════════════════════════════════
#  Fetch & cache
# ═══════════════════════════════════════════════════════════════

def fetch_quarter(year, quarter, cache_dir=None):
    """
    Download a DERA quarterly zip and return extracted DataFrames.
    Caches to disk as parquet if cache_dir is provided.

    Returns:
        dict: {"sub": DataFrame, "num": DataFrame, "pre": DataFrame, "tag": DataFrame}
        or None on failure
    """
    if cache_dir:
        cache_path = Path(cache_dir) / f"{year}q{quarter}"
        if cache_path.exists():
            return _load_cached(cache_path)

    url = f"{BASE_URL}/{year}q{quarter}.zip"
    logger.info(f"Fetching DERA: {url}")

    try:
        r = requests.get(url, headers=HEADERS, timeout=120)
        if r.status_code == 404:
            logger.warning(f"Not found: {year}q{quarter}")
            return None
        r.raise_for_status()
    except Exception as e:
        logger.warning(f"Failed to fetch {year}q{quarter}: {e}")
        return None

    zf = zipfile.ZipFile(io.BytesIO(r.content))
    data = {}
    for name in zf.namelist():
        if name.endswith(".txt"):
            key = name.replace(".txt", "")
            try:
                df = pd.read_csv(
                    zf.open(name), sep="\t",
                    low_memory=False, encoding="utf-8",
                    on_bad_lines="skip"
                )
                data[key] = df
            except Exception as e:
                logger.warning(f"Failed to read {name}: {e}")

    if cache_dir and data:
        cache_path = Path(cache_dir) / f"{year}q{quarter}"
        cache_path.mkdir(parents=True, exist_ok=True)
        for key, df in data.items():
            df.to_parquet(cache_path / f"{key}.parquet", index=False)
        logger.info(f"Cached to {cache_path}")

    return data if data else None


def _load_cached(cache_path):
    data = {}
    for f in cache_path.glob("*.parquet"):
        data[f.stem] = pd.read_parquet(f)
    return data if data else None


# ═══════════════════════════════════════════════════════════════
#  1. Label resolver
# ═══════════════════════════════════════════════════════════════

def resolve_labels(concepts, cik, year, quarter, cache_dir=None):
    """
    Look up human-readable labels, statement assignment, and line order
    for a list of XBRL concept names from DERA pre.txt.

    Args:
        concepts: list of concept names (e.g. ["Assets", "Revenue"])
        cik: company CIK (int)
        year, quarter: which DERA zip to use
        cache_dir: optional cache path

    Returns:
        dict: {concept_name: {
            "label": str,       # plabel from pre.txt
            "stmt": str,        # "Balance Sheet", etc.
            "stmt_code": str,   # "BS", "IS", etc.
            "line": int,        # presentation line order
            "negating": bool,   # should value be negated for display
        }}
    """
    data = fetch_quarter(year, quarter, cache_dir)
    if not data:
        return {}

    sub = data.get("sub")
    pre = data.get("pre")
    if sub is None or pre is None:
        return {}

    cik = int(str(cik).lstrip("0"))
    company_subs = sub[sub["cik"] == cik]
    if company_subs.empty:
        return {}

    accessions = set(company_subs["adsh"])
    company_pre = pre[pre["adsh"].isin(accessions)]
    if company_pre.empty:
        return {}

    concept_set = set(concepts)
    matches = company_pre[company_pre["tag"].isin(concept_set)]

    result = {}
    for _, row in matches.iterrows():
        tag = row["tag"]
        stmt_code = row.get("stmt", "")
        line = int(row.get("line", 999))

        if tag not in result or line < result[tag]["line"]:
            result[tag] = {
                "label": row.get("plabel", tag),
                "stmt": STMT_MAP.get(stmt_code, stmt_code),
                "stmt_code": stmt_code,
                "line": line,
                "negating": bool(row.get("negating", 0)),
            }

    return result


# ═══════════════════════════════════════════════════════════════
#  2. Full statement structure
# ═══════════════════════════════════════════════════════════════

def get_statement_structure(cik, year, quarter, form_type="10-K", cache_dir=None):
    """
    Get complete statement structures (all line items in order)
    for a company from a specific DERA quarter.

    Returns:
        dict: {
            "Balance Sheet": [
                {"tag": "Assets", "label": "Total assets", "line": 1, "negating": False},
                ...
            ],
            ...
        }
    """
    data = fetch_quarter(year, quarter, cache_dir)
    if not data:
        return {}

    sub = data.get("sub")
    pre = data.get("pre")
    if sub is None or pre is None:
        return {}

    cik = int(str(cik).lstrip("0"))
    company_subs = sub[sub["cik"] == cik]
    if form_type:
        company_subs = company_subs[company_subs["form"] == form_type]
    if company_subs.empty:
        return {}

    if "period" in company_subs.columns:
        company_subs = company_subs.sort_values("period", ascending=False)
    latest_adsh = company_subs.iloc[0]["adsh"]

    filing_pre = pre[pre["adsh"] == latest_adsh]
    if filing_pre.empty:
        return {}

    # Skip parenthetical
    if "inpth" in filing_pre.columns:
        filing_pre = filing_pre[filing_pre["inpth"] != 1]

    statements = {}
    for stmt_code, stmt_name in STMT_MAP.items():
        stmt_rows = filing_pre[filing_pre["stmt"] == stmt_code].sort_values("line")
        if stmt_rows.empty:
            continue

        items = []
        for _, row in stmt_rows.iterrows():
            items.append({
                "tag": row["tag"],
                "label": row.get("plabel", row["tag"]),
                "line": int(row.get("line", 0)),
                "negating": bool(row.get("negating", 0)),
            })
        statements[stmt_name] = items

    return statements


# ═══════════════════════════════════════════════════════════════
#  3. Get numeric values with statement context
# ═══════════════════════════════════════════════════════════════

def get_company_facts(cik, year, quarter, form_type="10-K", cache_dir=None):
    """
    Get all numeric facts for a company from a DERA quarter,
    enriched with statement assignment and labels from pre.txt.

    Returns:
        DataFrame: concept, date, value, qtrs, stmt, label, line, negating
        Consolidated USD only.
    """
    data = fetch_quarter(year, quarter, cache_dir)
    if not data:
        return pd.DataFrame()

    sub = data.get("sub")
    num = data.get("num")
    pre = data.get("pre")
    if sub is None or num is None:
        return pd.DataFrame()

    cik = int(str(cik).lstrip("0"))
    company_subs = sub[sub["cik"] == cik]
    if form_type:
        company_subs = company_subs[company_subs["form"] == form_type]
    if company_subs.empty:
        return pd.DataFrame()

    accessions = set(company_subs["adsh"])

    # Filter facts
    facts = num[num["adsh"].isin(accessions)].copy()
    if "coreg" in facts.columns:
        facts = facts[facts["coreg"].isna() | (facts["coreg"] == "")]
    if "uom" in facts.columns:
        facts = facts[facts["uom"] == "USD"]
    if facts.empty:
        return pd.DataFrame()

    facts["date"] = facts["ddate"].astype(str).apply(_format_ddate)

    # Enrich with presentation info
    if pre is not None:
        company_pre = pre[pre["adsh"].isin(accessions)]
        if not company_pre.empty:
            merge_on = [c for c in ["adsh", "tag", "version"]
                        if c in facts.columns and c in company_pre.columns]
            if merge_on:
                pre_cols = merge_on + [c for c in ["stmt", "line", "plabel", "negating"]
                                       if c in company_pre.columns]
                pre_info = company_pre[pre_cols].drop_duplicates(subset=merge_on)
                facts = facts.merge(pre_info, on=merge_on, how="left")

    # Map stmt codes to names
    if "stmt" in facts.columns:
        facts["stmt"] = facts["stmt"].map(STMT_MAP).fillna(facts["stmt"])

    # Rename and select
    rename = {"tag": "concept", "plabel": "label"}
    facts = facts.rename(columns=rename)
    keep = [c for c in ["concept", "date", "value", "qtrs", "stmt", "label", "line", "negating"]
            if c in facts.columns]
    return facts[keep]


# ═══════════════════════════════════════════════════════════════
#  Integration: enhance XBRL parser output
# ═══════════════════════════════════════════════════════════════

def enhance_statements(xbrl_result, cik, year, quarter, cache_dir=None):
    """
    Enhance extract_statements() output with DERA labels.
    Adds a "labels" key: {concept_name: "Human Readable Label"}.

    Does NOT change statement structure or values — just adds
    the label mapping for the Excel exporter to use.
    """
    all_concepts = set()
    for stmt_name, pivot_df in xbrl_result.get("statements", {}).items():
        all_concepts.update(pivot_df.index.tolist())

    if not all_concepts:
        xbrl_result["labels"] = {}
        return xbrl_result

    labels = resolve_labels(list(all_concepts), cik, year, quarter, cache_dir)
    label_map = {concept: info["label"] for concept, info in labels.items()}

    xbrl_result["labels"] = label_map
    logger.info(f"DERA resolved {len(label_map)}/{len(all_concepts)} concept labels")

    return xbrl_result


def _format_ddate(ddate_str):
    """YYYYMMDD → YYYY-MM-DD"""
    try:
        s = str(int(float(ddate_str)))
        if len(s) == 8:
            return f"{s[:4]}-{s[4:6]}-{s[6:]}"
        return s
    except (ValueError, TypeError):
        return str(ddate_str)