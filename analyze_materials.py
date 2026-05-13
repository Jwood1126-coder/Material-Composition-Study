#!/usr/bin/env python3
"""
NetSuite Material Composition Analyzer
=======================================
Analyzes material composition data from a NetSuite saved search CSV export.

Checks performed:
  1. Name suffix → material composition consistency
       -SS (or -SS-<anything>) → must be Stainless Steel
       -B (exact final segment) → must be Brass
       -BB is Brennan Black (a program/class), NOT Brass — not flagged
  2. Class name must reflect material composition
       e.g. "Hydraulic Fittings - 37° JIC - Steel" should contain "Steel"
       Brennan Black parts are exempt (material won't appear in that class by design)
  3. Matrix Material field must match Material Composition (when populated)
  4. Missing material composition
  5. Matrix parents detected (same Internal ID repeated → multiple compositions)

Sanity / data-quality guards baked in:
  - Column detection is case-insensitive and whitespace-tolerant
  - All string comparisons are case-insensitive and stripped
  - Encoding fallback (utf-8-sig → latin-1) for Excel-exported CSVs
  - -BB vs -B exact segment matching (no partial suffix collisions)
  - "Steel" in class does NOT satisfy "Stainless Steel" requirement
  - Multi-value composition fields (comma / semicolon / pipe separated)
  - Duplicate Internal ID detection → matrix parent flagging
  - Boolean flag columns clearly separated from informational notes
"""

import os
import re
import sys
import argparse
from pathlib import Path
from datetime import datetime

# In PyInstaller --windowed builds on Windows, stdout/stderr are None.
# Redirect to devnull so print() calls don't crash; CLI usage from a real
# terminal is unaffected because stdout/stderr exist there.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

import pandas as pd

__version__ = "1.12.1"

try:
    import xlsxwriter
    from xlsxwriter.utility import xl_col_to_name
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


# ── Constants ──────────────────────────────────────────────────────────────────

# Maps a normalized material composition string → keywords that MUST appear in
# the class name for the class to be considered consistent with the material.
MATERIAL_CLASS_KEYWORDS: dict[str, list[str]] = {
    "stainless steel":  ["stainless"],
    "steel":            ["steel"],          # "stainless" in class would be wrong — handled separately
    "brass":            ["brass"],
    "aluminum":         ["aluminum", "aluminium"],
    "aluminum alloy":   ["aluminum", "aluminium"],
    "carbon steel":     ["carbon"],
    "cast iron":        ["cast iron"],
    "copper":           ["copper"],
    "titanium":         ["titanium"],
    "plastic":          ["plastic", "nylon", "polymer"],
    "rubber":           ["rubber"],
    "zinc":             ["zinc"],
    "chrome":           ["chrome"],
    "nickel":           ["nickel"],
}

# Name suffix segment → expected Material Composition value.
# Keys are uppercase; values are the canonical display string.
SUFFIX_MATERIAL_MAP: dict[str, str] = {
    "SS": "Stainless Steel",
    "B":  "Brass",
    "D":  "Aluminum",
}

# Segments that look like a suffix but are NOT material indicators.
# Prevents false-positive material detection.
NON_MATERIAL_SEGMENTS: set[str] = {
    "BB",   # Brennan Black program
    "LW",   # Light weight — structural modifier, not material
    "SPL",  # Special
    "UNP",  # Unpainted
    "ZN",   # Zinc plated
}

# Classes that are explicitly exempt from the "class must contain material" check.
EXEMPT_CLASS_SUBSTRINGS: list[str] = [
    "brennan black",
]

# Issue flags — these contribute to any_flag and warrant human review.
# Every flag here represents POSITIVE evidence that something is wrong.
# Evidence is ranked by reliability per the user's directive:
#   1. Suffix / name pattern  — most authoritative
#   2. Legacy ERP material    — second
#   3. Class string           — third (often wrong)
# Higher-priority confirmation suppresses lower-priority disagreement
# flags, and the recommendation column always reflects the highest-
# priority signal in play.
# We deliberately do NOT flag *absence of evidence* (e.g. Brass material
# with no -B suffix and a generic class). Many parts use legitimate
# alternate naming conventions; flagging those would be noise.
ISSUE_FLAGS: dict[str, str] = {
    "flag_suffix_mismatch":            "Name Suffix ↔ Material Mismatch",
    "flag_legacy_disagrees":           "Legacy ERP Disagrees with Material Composition",
    "flag_name_legacy_conflict":       "Name Signal and Legacy ERP Disagree (Review)",
    "flag_catsy_disagrees":            "Catsy PIM Disagrees with Material Composition",
    "flag_name_catsy_conflict":        "Name Signal and Catsy Disagree (Review)",
    "flag_class_material_mismatch":    "Class Names Different Material",
    "flag_empty_material_composition": "Missing Material Composition",
}

# Informational flags — context only, no analysis note, do NOT trigger any_flag.
INFO_FLAGS: dict[str, str] = {
    "flag_is_matrix_parent":     "Matrix Parent (Multiple Compositions)",
    "flag_is_bb_part":           "Brennan Black Part (Class Check Exempt)",
    "flag_unknown_composition":  "Unknown Composition (Limited Validation)",
}

# Combined ordered dict used for display/export.
FLAG_META: dict[str, str] = {**ISSUE_FLAGS, **INFO_FLAGS}
FLAG_COLS = list(FLAG_META.keys())


# ── Name Parsing ───────────────────────────────────────────────────────────────

def get_name_segments(name: str) -> list[str]:
    """Split an item name on '-' and return cleaned, uppercased segments."""
    if not name or str(name).strip() in ("", "nan", "None"):
        return []
    return [s.strip().upper() for s in str(name).split("-") if s.strip()]


def detect_suffix_material(name: str) -> str | None:
    """
    Infer expected material composition from name segments.

    Rules (applied in order):
      - Any segment == 'SS' → Stainless Steel
        (handles -SS-LW, -SS-SPL, etc. — SS anywhere still means stainless)
      - Final segment == 'B' exactly → Brass
        (-BB is NOT brass; exact match guards against this)
      - Final segment == 'D' exactly → Aluminum

    Returns the canonical material string, or None if no suffix match.
    """
    segments = get_name_segments(name)
    if not segments:
        return None

    # SS can appear anywhere in the segment chain
    if "SS" in segments:
        return "Stainless Steel"

    # B / D must be the final segment exactly (no partial collisions like BB)
    if segments[-1] == "B":
        return "Brass"
    if segments[-1] == "D":
        return "Aluminum"

    return None


def is_bb_part(name: str, class_val: str) -> bool:
    """
    Return True if this part belongs to the Brennan Black program.
    BB parts are exempt from the class-contains-material check because
    the class IS 'Brennan Black' and won't contain the material by design.
    """
    segments = get_name_segments(name)
    class_lower = str(class_val).strip().lower() if class_val else ""

    if "BB" in segments:
        return True
    if any(sub in class_lower for sub in EXEMPT_CLASS_SUBSTRINGS):
        return True
    return False


# ── Material / Class Comparison ────────────────────────────────────────────────

def normalize(value: str) -> str:
    """Lowercase + strip a string; return '' for missing values."""
    if not value or str(value).strip() in ("", "nan", "None"):
        return ""
    return str(value).strip().lower()


def material_matches_class(material_comp: str, class_val: str) -> bool:
    """
    Return True if the class name is consistent with the material composition.

    Logic:
      1. Look up the material in MATERIAL_CLASS_KEYWORDS.
      2. At least one keyword must appear in the class string.
      3. Special guard: if material is plain 'steel' but class contains
         'stainless', that's a mismatch (wrong grade of steel).
      4. Unknown materials fall back to a simple substring check.
    """
    mat = normalize(material_comp)
    cls = normalize(class_val)

    if not mat or not cls:
        return True     # Cannot evaluate — skip, do not flag

    keywords = MATERIAL_CLASS_KEYWORDS.get(mat)

    if keywords is None:
        # Unknown material: simple substring check
        return mat in cls

    # Guard: plain "steel" should not accept a "stainless" class as a match
    if mat == "steel" and "stainless" in cls:
        return False

    return any(kw in cls for kw in keywords)


def split_multi_composition(value: str) -> list[str]:
    """
    Split a material composition field that may hold multiple values.
    Handles comma, semicolon, pipe, and newline delimiters.
    Returns a list of stripped, non-empty strings.
    """
    if not value or str(value).strip() in ("", "nan", "None"):
        return []
    parts = re.split(r"[,;\n|]+", str(value))
    return [p.strip() for p in parts if p.strip()]


# ── Core Analysis ──────────────────────────────────────────────────────────────

def _resolve_column(col_map: dict[str, str], *candidates: str) -> str | None:
    """Find the first matching column name (case-insensitive)."""
    for c in candidates:
        if c.lower() in col_map:
            return col_map[c.lower()]
    return None


# ── Legacy ERP Cross-Check ─────────────────────────────────────────────────────

# Sentinel placeholder text treated as missing in legacy CSVs (same set as the
# main sanitizer below — kept aligned so behavior is consistent).
_LEGACY_SENTINELS = {
    "nan", "NaN", "None", "none",
    "tbd", "TBD", "n/a", "N/A", "na", "NA",
    "unknown", "Unknown", "UNKNOWN",
    "?", "—", "-",
}


def build_legacy_lookup(legacy_df: pd.DataFrame) -> dict[str, str]:
    """
    Build a part-number → material lookup from a legacy ERP CSV.

    Required columns (matched case-insensitively, with common aliases):
      - Part Number  (or 'Part #', 'PartNum', 'Item Number', 'Item #')
      - Material     (or 'Material Composition')

    Returns a dict mapping lowercased part numbers → cleaned material strings.
    Blank part numbers, blank materials, and sentinel placeholders are skipped.
    On duplicate keys with conflicting materials, the FIRST occurrence wins
    (legacy data is treated as immutable; later rows are presumed to be
    corrections OR clerical re-entries — we don't know which, so the safer
    default is to surface conflicts as duplicates rather than silently
    overwrite).

    Sanity checks:
      - Both required columns must be present (otherwise raises ValueError)
      - Trailing/leading whitespace stripped from both columns
      - Sentinel placeholders ("TBD", "N/A", etc.) treated as blank
      - At least one valid (part, material) pair required
    """
    if legacy_df is None or len(legacy_df) == 0:
        return {}

    cols = {c.strip().lower(): c for c in legacy_df.columns}
    pn_col = next(
        (cols[k] for k in ("part number", "part #", "partnum", "part_number",
                           "item number", "item #", "item")
         if k in cols),
        None
    )
    mat_col = next(
        (cols[k] for k in ("material", "material composition",
                           "material_composition", "materialcomposition")
         if k in cols),
        None
    )
    if not pn_col or not mat_col:
        raise ValueError(
            f"Legacy ERP CSV must have a Part Number column and a Material column. "
            f"Found columns: {list(legacy_df.columns)}"
        )

    pn = (legacy_df[pn_col].fillna("").astype(str).str.strip())
    mat = (legacy_df[mat_col].fillna("").astype(str).str.strip())

    lookup: dict[str, str] = {}
    skipped_blank = 0
    conflicts = 0
    for raw_p, raw_m in zip(pn, mat):
        if raw_p in _LEGACY_SENTINELS or raw_m in _LEGACY_SENTINELS:
            skipped_blank += 1
            continue
        if not raw_p or not raw_m:
            skipped_blank += 1
            continue
        key = raw_p.lower()
        if key in lookup:
            if lookup[key].lower() != raw_m.lower():
                conflicts += 1
            # First-occurrence wins — don't overwrite
            continue
        lookup[key] = raw_m

    if not lookup:
        raise ValueError(
            "Legacy ERP CSV produced 0 usable (Part Number, Material) pairs — "
            "every row was blank or a placeholder."
        )
    return lookup


# ── Catsy PIM Cross-Check ──────────────────────────────────────────────────────

def build_catsy_lookup(catsy_df: pd.DataFrame) -> dict[str, str]:
    """
    Build a part-number → material lookup from a Catsy PIM CSV export.

    Expected columns (case-insensitive, with common aliases):
      - Items            ('item', 'sku', 'part number', 'part #')
      - Primary Material ('material', 'material composition', 'primarymaterial')

    Returns a dict mapping lowercased part numbers → cleaned material strings.
    Blank entries and sentinel placeholders (TBD, N/A, ?, —) are skipped.
    On duplicate part-number keys, the FIRST occurrence wins (same policy as
    the legacy ERP lookup — surfaces conflicts via the row count rather than
    silently overwriting).
    """
    if catsy_df is None or len(catsy_df) == 0:
        return {}

    cols = {c.strip().lower(): c for c in catsy_df.columns}
    item_col = next(
        (cols[k] for k in ("items", "item", "sku", "part number", "part #",
                           "partnum", "part_number", "item number", "item #")
         if k in cols),
        None
    )
    mat_col = next(
        (cols[k] for k in ("primary material", "primarymaterial",
                           "primary_material", "material",
                           "material composition", "material_composition")
         if k in cols),
        None
    )
    if not item_col or not mat_col:
        raise ValueError(
            f"Catsy CSV must have an 'Items' column and a 'Primary Material' "
            f"column. Found columns: {list(catsy_df.columns)}"
        )

    pn  = catsy_df[item_col].fillna("").astype(str).str.strip()
    mat = catsy_df[mat_col].fillna("").astype(str).str.strip()

    lookup: dict[str, str] = {}
    for raw_p, raw_m in zip(pn, mat):
        if raw_p in _LEGACY_SENTINELS or raw_m in _LEGACY_SENTINELS:
            continue
        if not raw_p or not raw_m:
            continue
        key = raw_p.lower()
        if key in lookup:
            continue        # first-occurrence wins
        lookup[key] = raw_m

    if not lookup:
        raise ValueError(
            "Catsy CSV produced 0 usable (Items, Primary Material) pairs — "
            "every row was blank or a placeholder."
        )
    return lookup


def analyze_dataframe(
    df: pd.DataFrame,
    enabled_checks: set[str] | None = None,
    progress_callback=None,
    legacy_lookup: dict[str, str] | None = None,
    catsy_lookup: dict[str, str] | None = None,
) -> pd.DataFrame:
    """
    Vectorized analysis pipeline. ~10–50× faster than per-row iteration on
    large exports.

    enabled_checks   : ISSUE_FLAGS keys to evaluate; None → all checks.
                       Disabled checks remain as columns but stay False.
    progress_callback: optional callable(pct: float, message: str) called at
                       phase boundaries (pct in [0.0, 1.0]).
    legacy_lookup    : optional dict mapping lowercased part-number → material
                       (built via build_legacy_lookup from a legacy ERP CSV).
                       When provided, each row is matched against External ID
                       first, then Name. Matches act as authoritative
                       confirmation; mismatches fire flag_legacy_disagrees.
    catsy_lookup     : optional dict mapping lowercased part-number → material
                       (built via build_catsy_lookup from a Catsy PIM CSV).
                       Matched by Name (Catsy "Items" column = part number).
                       Tier-2 equivalent to legacy_lookup — parallel cross-
                       reference. Fires flag_catsy_disagrees on disagreements;
                       used to fill blank NetSuite compositions too.
    """
    if enabled_checks is None:
        enabled_checks = set(ISSUE_FLAGS.keys())

    def report(pct: float, msg: str) -> None:
        if progress_callback:
            try:
                progress_callback(pct, msg)
            except Exception:
                pass    # never let UI updates break analysis

    report(0.0, "Preparing data…")

    df = df.copy()

    # ── Resolve / validate columns ─────────────────────────────────────────
    df.columns = [c.strip() for c in df.columns]
    col_map = {c.lower(): c for c in df.columns}

    col_id       = _resolve_column(col_map, "internal id", "internalid", "internal_id", "id")
    col_name     = _resolve_column(col_map, "name", "item name", "item_name")
    col_ext_id   = _resolve_column(
        col_map, "external id", "externalid", "external_id", "ext id",
        "external", "name (external id)"
    )
    col_class    = _resolve_column(col_map, "class", "item class", "item_class")
    col_material = _resolve_column(col_map, "material")
    col_mat_comp = _resolve_column(
        col_map, "material composition", "material_composition", "materialcomposition"
    )

    missing = [
        label for label, col in [
            ("Name", col_name),
            ("Class", col_class),
            ("Material Composition", col_mat_comp),
        ]
        if col is None
    ]
    if missing:
        raise ValueError(
            f"Required column(s) not found: {', '.join(missing)}\n"
            f"Columns in file: {list(df.columns)}"
        )

    n = len(df)
    idx = df.index

    # ── Sanitize string fields (vectorized) ────────────────────────────────
    # Treat common placeholder text as missing — real NetSuite exports often
    # contain these instead of true blanks.
    sentinel_map = {
        "nan": "", "NaN": "", "None": "", "none": "",
        "tbd": "", "TBD": "", "n/a": "", "N/A": "", "na": "", "NA": "",
        "unknown": "", "Unknown": "", "UNKNOWN": "",
        "?": "", "—": "", "-": "",
    }
    for col in filter(None, [col_id, col_name, col_ext_id, col_class, col_material, col_mat_comp]):
        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace(sentinel_map)
        )

    name      = df[col_name]
    klass     = df[col_class]     if col_class    else pd.Series([""] * n, index=idx)
    matc      = df[col_mat_comp]  if col_mat_comp else pd.Series([""] * n, index=idx)
    material  = df[col_material]  if col_material else pd.Series([""] * n, index=idx)
    int_id    = df[col_id]        if col_id       else pd.Series([""] * n, index=idx)
    ext_id    = df[col_ext_id]    if col_ext_id   else pd.Series([""] * n, index=idx)

    klass_lower = klass.str.lower()
    matc_lower  = matc.str.lower()
    mat_lower   = material.str.lower()

    # NetSuite matrix children carry the parent prefix in their Name —
    # e.g. "NS2404:NS2404-12-06-SS". Catsy and legacy ERP typically store
    # just the child SKU ("NS2404-12-06-SS"), so we also try the post-colon
    # portion of Name when looking up cross-references. rsplit with n=1
    # takes everything AFTER the last colon; for rows without a colon the
    # result equals the full name, so the fallback is a no-op there.
    name_lower            = name.str.lower()
    name_child_lower      = name_lower.str.rsplit(":", n=1).str[-1]

    # ── Legacy ERP lookup ──────────────────────────────────────────────────
    # Tried in order: External ID, full Name, matrix-child portion of Name.
    if legacy_lookup:
        ext_lower  = ext_id.str.lower()
        legacy_via_ext   = ext_lower.map(legacy_lookup).fillna("")
        legacy_via_name  = name_lower.map(legacy_lookup).fillna("")
        legacy_via_child = name_child_lower.map(legacy_lookup).fillna("")
        legacy_mat = legacy_via_ext.where(legacy_via_ext   != "", legacy_via_name)
        legacy_mat = legacy_mat.where(legacy_mat           != "", legacy_via_child)
    else:
        legacy_mat = pd.Series([""] * n, index=idx, dtype=object)
    legacy_lower = legacy_mat.str.lower()
    has_legacy   = legacy_lower != ""

    # ── Catsy PIM lookup ───────────────────────────────────────────────────
    # Parallel cross-reference to Legacy ERP. Tried in order: full Name,
    # then the matrix-child portion of Name (text after the last colon)
    # so matrix items in NetSuite ("PARENT:CHILD") match Catsy rows that
    # store only the child SKU.
    if catsy_lookup:
        catsy_via_name  = name_lower.map(catsy_lookup).fillna("")
        catsy_via_child = name_child_lower.map(catsy_lookup).fillna("")
        catsy_mat = catsy_via_name.where(catsy_via_name != "", catsy_via_child)
    else:
        catsy_mat = pd.Series([""] * n, index=idx, dtype=object)
    catsy_lower = catsy_mat.str.lower()
    has_catsy   = catsy_lower != ""

    report(0.10, "Parsing names…")

    # ── Vectorized name segment analysis ───────────────────────────────────
    segments = (
        name.str.upper()
            .str.split("-")
            .map(lambda s: [seg.strip() for seg in s if seg.strip()] if isinstance(s, list) else [])
    )
    # ── Material signals must appear in the LAST 3 segments only ───────────
    # Per user directive: "Materials usually are near the end of the part,
    # perhaps in the last 3 segments max, never earlier." This eliminates
    # false positives like "D-0306-08" or "STEEL-COUPLER-IS-BEST-08" where
    # a material-like word appears too early to be the material indicator.
    last_3_segs  = segments.map(lambda s: s[-3:] if s else [])
    last_3_text  = last_3_segs.map(lambda s: " ".join(s).lower() if s else "")
    last_seg     = segments.map(lambda s: s[-1] if s else "")

    has_ss        = last_3_segs.map(lambda s: "SS" in s)
    has_d         = last_3_segs.map(lambda s: "D"  in s)
    has_zn        = last_3_segs.map(lambda s: "ZN" in s)   # zinc plated → Steel
    has_bb        = segments.map(lambda s: "BB" in s)      # BB anywhere — class exemption
    has_b_suffix  = (last_seg == "B")    # exact -B (not -BB)

    # Expected material from NAME signals. Priority:
    #   SS segment > -B last seg > D segment > -ZN segment (→ Steel) > keywords
    # All restricted to last-3 scope.
    expected_mat = pd.Series([""] * n, index=idx, dtype=object)
    expected_mat = expected_mat.mask(has_ss, "Stainless Steel")
    expected_mat = expected_mat.mask(~has_ss & has_b_suffix, "Brass")
    expected_mat = expected_mat.mask(~has_ss & ~has_b_suffix & has_d, "Aluminum")
    expected_mat = expected_mat.mask(
        ~has_ss & ~has_b_suffix & ~has_d & has_zn, "Steel"  # zinc plated = steel base
    )

    # `recommended_mat` is finalized after class_confirms / legacy / forward
    # mismatch are all known — placeholder; real value is set below.

    # ── Multi-composition + matrix parent detection ────────────────────────
    has_multi_comp = matc.str.contains(r"[,;|\n]", regex=True, na=False)

    if col_id:
        nonblank = int_id != ""
        id_counts = int_id[nonblank].value_counts()
        matrix_ids = set(id_counts[id_counts > 1].index)
        is_matrix_parent = nonblank & int_id.isin(matrix_ids)
    else:
        is_matrix_parent = pd.Series([False] * n, index=idx)

    is_bb_class = klass_lower.str.contains("brennan black", na=False)
    is_bb_flag  = has_bb | is_bb_class
    is_empty    = (matc_lower == "")

    # ── Class & name keyword extraction ─────────────────────────────────────
    # Each pattern is broken out individually so we can report exactly which
    # signal fired on each row (see the "Matched Signal" column).
    name_lower = name.str.lower()

    # Class signals — all single-purpose, no compound regexes
    contains_stainless = klass_lower.str.contains("stainless", na=False)
    contains_steel     = klass_lower.str.contains("steel",     na=False)
    contains_brass     = klass_lower.str.contains("brass",     na=False)
    contains_alum      = klass_lower.str.contains(r"alumin(?:um|ium)", regex=True, na=False)
    # Aerospace is informational only — surfaced in matched_signal but does
    # NOT confirm Aluminum (aerospace fittings can be Al, Ti, steel, etc.).
    contains_aerospace = klass_lower.str.contains("aerospace", na=False)

    # ── Name keyword signals — STRICT and restricted to last 3 segments ────
    # Per user directives:
    #   - Stainless: ONLY a literal "-SS" segment. Period.
    #   - Brass: ONLY a literal "-B" final segment. B must stand alone, not
    #     combined with other letters (no "BR", "HNBR", "30B", etc.).
    #   - Aluminum: literal "-D" segment OR the "alum" keyword (in last 3 segs).
    #   - Steel: literal "-ZN" segment (zinc plated) OR the "steel" keyword
    #     (in last 3 segs).
    # All keyword matches use last_3_text so a part name like "STEEL-IS-X-Y-08"
    # would NOT match (STEEL is too early).
    name_alum_word  = last_3_text.str.contains(r"\balum",   regex=True, na=False)
    name_steel_word = last_3_text.str.contains(r"\bsteel\b", regex=True, na=False)

    name_says_alum      = name_alum_word
    name_says_brass     = pd.Series([False] * n, index=idx)   # only has_b_suffix
    name_says_stainless = pd.Series([False] * n, index=idx)   # only has_ss
    name_says_steel     = name_steel_word

    # Fall-through: when no suffix-derived material was set, use the name
    # patterns. Order matters when multiple patterns match — we apply most
    # specific first (Stainless > Aluminum > Brass > Steel).
    expected_mat = expected_mat.mask(
        (expected_mat == "") & name_says_stainless, "Stainless Steel"
    )
    expected_mat = expected_mat.mask(
        (expected_mat == "") & name_says_alum, "Aluminum"
    )
    expected_mat = expected_mat.mask(
        (expected_mat == "") & name_says_brass, "Brass"
    )
    expected_mat = expected_mat.mask(
        (expected_mat == "") & name_says_steel, "Steel"
    )

    is_alum_comp = matc_lower.isin(["aluminum", "aluminium", "aluminum alloy"])

    # ── Priority-based confirmation signals ────────────────────────────────
    # Evidence ranking per the user's directive (most → least authoritative):
    #   1. Suffix / name pattern (-SS, -B, -D, "ALUM", "316", "BR", etc.)
    #   2. Legacy ERP material lookup
    #   3. Class string (often wrong, partial evidence only)
    #
    # Each tier confirms the composition independently — and a higher-tier
    # confirmation suppresses lower-tier disagreement flags. Class never
    # suppresses a suffix or legacy mismatch — class is the least reliable
    # source and can't override the others.
    expected_lower      = expected_mat.str.lower()
    name_confirms_comp  = (
        (expected_mat != "") & (matc_lower != "") & (expected_lower == matc_lower)
    )
    legacy_confirms_comp = has_legacy & (legacy_lower == matc_lower)
    catsy_confirms_comp  = has_catsy  & (catsy_lower  == matc_lower)

    report(0.30, "Checking suffix rules…")

    # ── FLAG: empty material composition ───────────────────────────────────
    if "flag_empty_material_composition" in enabled_checks:
        flag_empty = is_empty
    else:
        flag_empty = pd.Series([False] * n, index=idx)

    # ── FLAG: suffix / name mismatch (highest priority) ─────────────────────
    # Fires whenever the name (suffix or pattern) implies a different
    # material than the current composition. The name is the most
    # authoritative signal — a class confirmation does NOT suppress this,
    # because class is the least reliable evidence source and shouldn't
    # override the name.
    forward_mismatch = (
        (expected_mat != "") & (matc_lower != "") & (expected_lower != matc_lower)
    )
    if "flag_suffix_mismatch" in enabled_checks:
        flag_suffix = forward_mismatch
    else:
        flag_suffix = pd.Series([False] * n, index=idx)
        forward_mismatch = pd.Series([False] * n, index=idx)

    # ── FLAG: legacy ERP disagrees (second-priority) ───────────────────────
    # Fires when the legacy lookup disagrees with the current composition,
    # UNLESS the name (higher priority) confirms the composition. In that
    # case the legacy entry is most likely the wrong one.
    if "flag_legacy_disagrees" in enabled_checks:
        flag_legacy_disagrees = (
            has_legacy & (matc_lower != "") & (legacy_lower != matc_lower)
            & ~name_confirms_comp
        )
    else:
        flag_legacy_disagrees = pd.Series([False] * n, index=idx)

    # ── FLAG: Catsy PIM disagrees (parallel to legacy, tier 2) ─────────────
    # Fires when Catsy has a different material than NetSuite, UNLESS the
    # name (tier 1) confirms the NetSuite composition is right. Legacy and
    # Catsy are parallel — Catsy is NOT silenced by a legacy confirmation
    # and vice versa, so genuine PIM/ERP disagreements stay visible.
    if "flag_catsy_disagrees" in enabled_checks:
        flag_catsy_disagrees = (
            has_catsy & (matc_lower != "") & (catsy_lower != matc_lower)
            & ~name_confirms_comp
        )
    else:
        flag_catsy_disagrees = pd.Series([False] * n, index=idx)

    report(0.50, "Checking class consistency…")

    # ── Class-says-which-material? ─────────────────────────────────────────
    # The class string may name a specific material (e.g. "Hyd Fittings -
    # Stainless"). Used for two purposes:
    #   (1) flag_class_material_mismatch: when class names a SPECIFIC
    #       material that disagrees with the current Material Composition,
    #       flag the row as a candidate composition fix (NOT a class fix —
    #       this version's focus is correcting Material Composition).
    #   (2) Recommended Material override: when the class names a different
    #       material than the current composition, the class becomes the
    #       suggested replacement.
    # Generic class strings (no recognized material keyword) produce "" and
    # do NOT contribute a flag — absence of evidence is not evidence.
    class_material_named = pd.Series([""] * n, index=idx, dtype=object)
    class_material_named = class_material_named.mask(contains_stainless, "Stainless Steel")
    class_material_named = class_material_named.mask(
        contains_steel & ~contains_stainless & (class_material_named == ""),
        "Steel"
    )
    class_material_named = class_material_named.mask(
        contains_brass & (class_material_named == ""),
        "Brass"
    )
    class_material_named = class_material_named.mask(
        contains_alum & (class_material_named == ""),
        "Aluminum"
    )

    # ── FLAG: class names a different material than composition ───────────
    flag_class = pd.Series([False] * n, index=idx)
    class_note_text = pd.Series([""] * n, index=idx, dtype=object)

    if "flag_class_material_mismatch" in enabled_checks:
        # Eligible: class is non-empty AND non-BB AND composition is set AND
        # class actually names a specific material (skip generic classes).
        # Suppressed when a higher-priority signal already confirms comp:
        #   - name (suffix/pattern) says comp is right → class is wrong, ignore
        #   - legacy says comp is right → class is wrong, ignore
        eligible = (
            ~is_bb_flag & (matc_lower != "") & (klass_lower != "")
            & (class_material_named != "")
            & ~name_confirms_comp
            & ~legacy_confirms_comp
            & ~catsy_confirms_comp
        )

        # Single-composition rows: simple inequality
        single_eligible = eligible & ~has_multi_comp
        single_disagree = (
            single_eligible &
            (class_material_named.str.lower() != matc_lower)
        )
        flag_class |= single_disagree
        class_note_text = class_note_text.mask(
            single_disagree,
            "Class indicates '" + class_material_named
            + "' but Material Composition is '" + matc + "'"
        )

        # Multi-composition rows: flag when class-named material is NOT
        # among the row's compositions
        multi_eligible = eligible & has_multi_comp
        if multi_eligible.any():
            def _multi_class_check(row):
                cmn = class_material_named.at[row.name].lower()
                if not cmn:
                    return False
                comps = [normalize(c) for c in split_multi_composition(str(row[col_mat_comp]))]
                return cmn not in comps
            multi_flag = df[multi_eligible].apply(_multi_class_check, axis=1)
            flag_class.loc[multi_flag.index] |= multi_flag.fillna(False)
            class_note_text.loc[multi_flag.index] = class_note_text.loc[multi_flag.index].mask(
                multi_flag.fillna(False),
                "Class indicates '" + class_material_named.loc[multi_flag.index]
                + "' but Material Composition is '" + matc.loc[multi_flag.index] + "'"
            )

    # ── FLAG: name vs legacy conflict ──────────────────────────────────────
    # Fires when the name signal (suffix or keyword) and the legacy ERP
    # disagree about the material, regardless of which (if any) agrees with
    # the current composition. Surfaces the case where both sources are
    # present but say different things — review which is correct.
    name_signal_present = (expected_mat != "")
    name_legacy_conflict = (
        name_signal_present & has_legacy & (expected_lower != legacy_lower)
    )
    if "flag_name_legacy_conflict" in enabled_checks:
        flag_name_legacy_conflict = name_legacy_conflict
    else:
        flag_name_legacy_conflict = pd.Series([False] * n, index=idx)

    # ── FLAG: name vs Catsy conflict ───────────────────────────────────────
    # Same shape as name↔legacy. Two trusted signals disagreeing — worth a
    # human's attention regardless of what the composition currently says.
    name_catsy_conflict = (
        name_signal_present & has_catsy & (expected_lower != catsy_lower)
    )
    if "flag_name_catsy_conflict" in enabled_checks:
        flag_name_catsy_conflict = name_catsy_conflict
    else:
        flag_name_catsy_conflict = pd.Series([False] * n, index=idx)

    # any_flag = OR of all issue flags (positive evidence only).
    any_flag = (
        flag_suffix | flag_legacy_disagrees | flag_name_legacy_conflict
        | flag_catsy_disagrees | flag_name_catsy_conflict
        | flag_class | flag_empty
    )

    # ── Recommended Material (final) ────────────────────────────────────────
    # Apply overrides in REVERSE priority order so the highest-priority
    # signal lands last (mask = last write wins). Priority per the user's
    # directive:  Suffix/Name > Legacy > Class.
    #
    # Each override only fires when the corresponding flag fires — so
    # confirming sources don't overwrite the current composition with a
    # value the row already agrees with.
    recommended_mat = matc.copy()
    # ── Blank composition: conservative estimate from the SAME rules ───────
    # No generic "Steel" default. We only estimate when there's actual
    # evidence to draw on. Priority for blanks (least → most authoritative):
    #   1. Class names a material
    #   2. Legacy ERP has a value
    #   3. Name (suffix or keyword) implies a material
    # If NO signal is available, leave the recommendation BLANK — per user
    # directive: "Be ok with not giving a recommendation. If there is no
    # clear signal, leave it blank."
    blank = (matc == "")
    blank_with_class = blank & (class_material_named != "")
    recommended_mat = recommended_mat.mask(blank_with_class, class_material_named)
    blank_with_catsy = blank & has_catsy
    recommended_mat = recommended_mat.mask(blank_with_catsy, catsy_mat)
    blank_with_legacy = blank & has_legacy
    recommended_mat = recommended_mat.mask(blank_with_legacy, legacy_mat)
    blank_with_name = blank & (expected_mat != "")
    recommended_mat = recommended_mat.mask(blank_with_name, expected_mat)
    # Remember which rows ended up with no estimate (for confidence + fix text)
    no_estimate = (recommended_mat == "")

    # ── Override-based recommendations on rows where the composition IS set
    # but a higher-priority signal says it's wrong ──────────────────────
    # Apply in reverse priority order (last write wins).
    recommended_mat = recommended_mat.mask(flag_class, class_material_named)
    recommended_mat = recommended_mat.mask(flag_catsy_disagrees, catsy_mat)
    recommended_mat = recommended_mat.mask(flag_legacy_disagrees, legacy_mat)
    recommended_mat = recommended_mat.mask(forward_mismatch, expected_mat)

    # ── Matched Signal text (for transparency / false-positive auditing) ──
    # Build a per-row description of which signals fired. Lets a human
    # eyeball a flagged row and immediately see why — catches the worst
    # regex-misfire false positives.
    name_signal_text = pd.Series([""] * n, index=idx, dtype=object)
    # Literal suffix signals (highest priority — applied first)
    name_signal_text = name_signal_text.mask(has_ss, "-SS segment")
    name_signal_text = name_signal_text.mask(
        has_b_suffix & ~has_ss & (name_signal_text == ""), "-B segment"
    )
    name_signal_text = name_signal_text.mask(
        has_d & ~has_ss & ~has_b_suffix & (name_signal_text == ""), "-D segment"
    )
    name_signal_text = name_signal_text.mask(
        has_zn & ~has_ss & ~has_b_suffix & ~has_d & (name_signal_text == ""),
        "-ZN segment (zinc plated)"
    )
    # Keyword signals (last 3 segments only)
    name_signal_text = name_signal_text.mask(
        (name_signal_text == "") & name_alum_word,  "'alum' keyword"
    )
    name_signal_text = name_signal_text.mask(
        (name_signal_text == "") & name_steel_word, "'steel' keyword"
    )

    # Boolean masks used by confidence scoring below
    literal_suffix_signal = (has_ss | has_b_suffix | has_d | has_zn)
    keyword_signal = (name_alum_word | name_steel_word)

    # Class signal — material the class names (or "aerospace" as informational)
    class_signal_text = class_material_named.copy()
    class_signal_text = class_signal_text.mask(
        (class_signal_text == "") & contains_aerospace,
        "aerospace (informational only)"
    )

    # Compose the Matched Signal column
    SIGSEP = " | "
    matched_signal = pd.Series([""] * n, index=idx, dtype=object)
    matched_signal = matched_signal.mask(
        name_signal_text != "",
        matched_signal + SIGSEP + "name: " + name_signal_text
    )
    matched_signal = matched_signal.mask(
        has_legacy,
        matched_signal + SIGSEP + "legacy: " + legacy_mat
    )
    matched_signal = matched_signal.mask(
        has_catsy,
        matched_signal + SIGSEP + "catsy: " + catsy_mat
    )
    matched_signal = matched_signal.mask(
        class_signal_text != "",
        matched_signal + SIGSEP + "class: " + class_signal_text
    )
    matched_signal = matched_signal.str.replace(r"^ \| ", "", regex=True)
    matched_signal = matched_signal.where(matched_signal != "", "—")

    # ── Confidence column (concordance-based) ──────────────────────────────
    # The recommended material is supported by some combination of three
    # signals: name (suffix or keyword), legacy ERP, class. Confidence is
    # a function of WHICH supported it and HOW MANY agreed.
    #
    #   HIGH:   2+ signals agree on the recommendation, OR a literal suffix
    #           (-SS, -B, -D, -ZN) supports it with no contradicting source.
    #   MEDIUM: A single legacy match OR a keyword name signal stands alone.
    #           Also: HIGH downgraded when name and legacy disagree.
    #   LOW:    Class is the only supporting signal.
    #   blank:  No flag fires (row is OK) or no signal at all (blank row
    #           with nothing to estimate from).
    rec_lower = recommended_mat.str.lower()
    name_signal_present  = (expected_mat != "")
    name_supports_rec    = name_signal_present & (expected_lower == rec_lower)
    legacy_supports_rec  = has_legacy & (legacy_lower == rec_lower)
    catsy_supports_rec   = has_catsy  & (catsy_lower  == rec_lower)
    class_supports_rec   = (class_material_named != "") & (class_material_named.str.lower() == rec_lower)
    n_supporting = (
        name_supports_rec.astype(int)
        + legacy_supports_rec.astype(int)
        + catsy_supports_rec.astype(int)
        + class_supports_rec.astype(int)
    )

    confidence = pd.Series([""] * n, index=idx, dtype=object)
    # Rows that get a confidence rating: anything flagged (we made a
    # recommendation) AND the recommendation is not blank.
    needs = any_flag & (recommended_mat != "")

    # Apply lowest → highest (last write wins).
    # LOW: class is the only supporting signal
    confidence = confidence.mask(
        needs & class_supports_rec
        & ~legacy_supports_rec & ~catsy_supports_rec & ~name_supports_rec,
        "Low"
    )
    # MEDIUM: legacy alone, OR catsy alone, OR keyword name alone
    confidence = confidence.mask(
        needs & legacy_supports_rec & ~name_supports_rec & ~catsy_supports_rec,
        "Medium"
    )
    confidence = confidence.mask(
        needs & catsy_supports_rec & ~name_supports_rec & ~legacy_supports_rec,
        "Medium"
    )
    confidence = confidence.mask(
        needs & name_supports_rec & keyword_signal & ~literal_suffix_signal
        & ~legacy_supports_rec & ~catsy_supports_rec,
        "Medium"
    )
    # HIGH: literal suffix supporting, OR multiple sources agreeing
    confidence = confidence.mask(
        needs & literal_suffix_signal & name_supports_rec,
        "High"
    )
    confidence = confidence.mask(
        needs & (n_supporting >= 2),
        "High"
    )
    # Conflict downgrade: when the name signal disagrees with a tier-2
    # source (legacy or catsy), one of them is wrong → confidence drops.
    confidence = confidence.mask(
        needs & (flag_name_legacy_conflict | flag_name_catsy_conflict),
        "Medium"
    )

    # (name_legacy_conflict and flag_name_legacy_conflict were computed
    # earlier — before any_flag — so the flag can contribute to it.)

    # ── Unknown-composition info flag (data quality) ───────────────────────
    # When the composition isn't one of our known materials, only the
    # substring class check applied — limited validation. Surface this so
    # the user knows which rows had weaker scrutiny.
    KNOWN_MATERIALS = {
        "stainless steel", "steel", "brass",
        "aluminum", "aluminium", "aluminum alloy",
    }
    flag_unknown_composition = (
        (matc_lower != "") & ~matc_lower.isin(KNOWN_MATERIALS)
    )

    report(0.85, "Composing analysis notes…")

    # ── Compose analysis_notes (vectorized string assembly) ────────────────
    parts = pd.Series([""] * n, index=idx, dtype=object)
    SEP = " | "

    if "flag_empty_material_composition" in enabled_checks:
        parts = parts.mask(flag_empty, parts + SEP + "Material Composition is blank")

    # Legacy ERP disagreement note — high-priority, high-confidence
    if "flag_legacy_disagrees" in enabled_checks:
        legacy_disagree_text = (
            "Legacy ERP says '" + legacy_mat
            + "' but Material Composition is '" + matc + "'"
        )
        parts = parts.mask(flag_legacy_disagrees, parts + SEP + legacy_disagree_text)

    # Catsy PIM disagreement note — parallel to legacy
    if "flag_catsy_disagrees" in enabled_checks:
        catsy_disagree_text = (
            "Catsy says '" + catsy_mat
            + "' but Material Composition is '" + matc + "'"
        )
        parts = parts.mask(flag_catsy_disagrees, parts + SEP + catsy_disagree_text)

    if "flag_suffix_mismatch" in enabled_checks:
        # Name (suffix or keyword) implies a different material than comp
        fwd_text = (
            "Name implies '" + expected_mat
            + "' but Material Composition is '" + matc + "'"
        )
        parts = parts.mask(forward_mismatch, parts + SEP + fwd_text)

    # Name vs Legacy conflict — fires whenever the two disagree, even when
    # neither is in conflict with the composition itself. Legacy is usually
    # right too, so the user should look at these.
    if "flag_name_legacy_conflict" in enabled_checks:
        conflict_warning = (
            "Name signal says '" + expected_mat
            + "' but Legacy ERP says '" + legacy_mat
            + "' — investigate which is correct"
        )
        parts = parts.mask(
            flag_name_legacy_conflict,
            parts + SEP + conflict_warning
        )

    if "flag_name_catsy_conflict" in enabled_checks:
        catsy_conflict_warning = (
            "Name signal says '" + expected_mat
            + "' but Catsy says '" + catsy_mat
            + "' — investigate which is correct"
        )
        parts = parts.mask(
            flag_name_catsy_conflict,
            parts + SEP + catsy_conflict_warning
        )

    if "flag_class_material_mismatch" in enabled_checks:
        parts = parts.mask(flag_class, parts + SEP + class_note_text)

    # Strip leading separator and replace empty → "OK"
    notes = parts.str.replace(r"^ \| ", "", regex=True)
    notes = notes.where(notes != "", "OK")

    # ── Suggested Fix column ───────────────────────────────────────────────
    # Templated, actionable text per row. Picked in priority order — legacy
    # ERP first (most authoritative), then suffix, class, matrix, empty, soft.
    fix = pd.Series([""] * n, index=idx, dtype=object)

    # Legacy ERP disagrees — strong signal, recommend the legacy value
    fix = fix.mask(
        flag_legacy_disagrees,
        "Legacy ERP says '" + legacy_mat
        + "' — verify and update Material Composition to match (or correct legacy if NetSuite is right)"
    )
    # Catsy PIM disagrees — parallel to legacy
    fix = fix.mask(
        flag_catsy_disagrees & (fix == ""),
        "Catsy says '" + catsy_mat
        + "' — verify and update Material Composition to match (or correct Catsy if NetSuite is right)"
    )
    # Name vs Legacy conflict — two trusted signals disagree with each other
    fix = fix.mask(
        flag_name_legacy_conflict & (fix == ""),
        "Name says '" + expected_mat + "' but Legacy ERP says '" + legacy_mat
        + "' — investigate which source is correct before changing composition"
    )
    # Name vs Catsy conflict — same shape
    fix = fix.mask(
        flag_name_catsy_conflict & (fix == ""),
        "Name says '" + expected_mat + "' but Catsy says '" + catsy_mat
        + "' — investigate which source is correct before changing composition"
    )

    # Composition empty — give an evidence-based estimate per priority:
    # name > legacy > catsy > class > (no estimate)
    fix = fix.mask(
        flag_empty & no_estimate & (fix == ""),
        "Material Composition is blank and no name/legacy/catsy/class signal — manual review needed"
    )
    fix = fix.mask(
        flag_empty & name_signal_present & (fix == ""),
        "Material Composition is blank; name suggests '" + expected_mat
        + "' — verify and set to this value"
    )
    fix = fix.mask(
        flag_empty & has_legacy & ~name_signal_present & (fix == ""),
        "Material Composition is blank; legacy ERP says '" + legacy_mat
        + "' — set to this value"
    )
    fix = fix.mask(
        flag_empty & has_catsy & ~name_signal_present & ~has_legacy & (fix == ""),
        "Material Composition is blank; Catsy says '" + catsy_mat
        + "' — set to this value"
    )
    fix = fix.mask(
        flag_empty & blank_with_class & ~name_signal_present & ~has_legacy & ~has_catsy & (fix == ""),
        "Material Composition is blank; class indicates '" + class_material_named
        + "' — verify and set to this value"
    )

    # Name implies a different material → name is the most authoritative
    # signal, suggest changing Material Composition to match the name
    fix = fix.mask(
        forward_mismatch & (fix == ""),
        "Name suggests material is '" + expected_mat
        + "' — verify and update Material Composition to '" + expected_mat + "'"
    )
    # Class mismatch: class names a different material than composition.
    # Focus is fixing Material Composition (this version's scope), so the
    # suggestion is to update the composition to match what the class says.
    fix = fix.mask(
        flag_class & (fix == ""),
        "Class indicates '" + class_material_named
        + "' — verify and update Material Composition to '" + class_material_named + "'"
    )
    # Missing composition (no legacy)
    fix = fix.mask(
        flag_empty & (fix == ""),
        "Populate the Material Composition field"
    )
    fix = fix.where(fix != "", "—")

    report(0.95, "Finalizing…")

    # ── Assemble result ────────────────────────────────────────────────────
    result = df.copy()
    result["flag_suffix_mismatch"]            = flag_suffix.astype(bool)
    result["flag_legacy_disagrees"]           = flag_legacy_disagrees.astype(bool)
    result["flag_name_legacy_conflict"]       = flag_name_legacy_conflict.astype(bool)
    result["flag_catsy_disagrees"]            = flag_catsy_disagrees.astype(bool)
    result["flag_name_catsy_conflict"]        = flag_name_catsy_conflict.astype(bool)
    result["flag_class_material_mismatch"]    = flag_class.astype(bool)
    result["flag_empty_material_composition"] = flag_empty.astype(bool)
    result["flag_is_matrix_parent"]           = is_matrix_parent.astype(bool)
    result["flag_is_bb_part"]                 = is_bb_flag.astype(bool)
    result["flag_unknown_composition"]        = flag_unknown_composition.astype(bool)
    result["legacy_material"]                 = legacy_mat
    result["catsy_material"]                  = catsy_mat
    result["recommended_material"]            = recommended_mat
    result["expected_material_from_name"]     = expected_mat
    result["matched_signal"]                  = matched_signal
    result["confidence"]                      = confidence
    result["analysis_notes"]                  = notes
    result["suggested_fix"]                   = fix
    result["any_flag"]                        = any_flag.astype(bool)

    report(1.0, "Analysis complete")
    return result


# ── Terminal Report ────────────────────────────────────────────────────────────

def print_summary(df: pd.DataFrame) -> None:
    total   = len(df)
    flagged = int(df["any_flag"].sum())
    pct     = flagged / total * 100 if total else 0.0

    width = 62
    sep   = "═" * width

    print(f"\n{sep}")
    print("  NETSUITE MATERIAL COMPOSITION — ANALYSIS REPORT")
    print(sep)
    print(f"  {'Total records analyzed':<40} {total:>8,}")
    print(f"  {'Records with issues':<40} {flagged:>8,}  ({pct:.1f}%)")
    print(f"  {'Clean records':<40} {total - flagged:>8,}")
    print()
    print(f"  {'── Issue Breakdown':─<{width - 4}}")
    for col, label in FLAG_META.items():
        if col in df.columns:
            count = int(df[col].sum())
            marker = "⚠" if count > 0 else " "
            print(f"  {marker} {label:<42} {count:>6,}")
    print(sep)


def print_flagged_details(df: pd.DataFrame, col_name: str) -> None:
    flagged = df[df["any_flag"] == True]
    if flagged.empty:
        print("\nNo flagged records.")
        return

    print(f"\nFlagged Records ({len(flagged):,}):\n")
    cols = [c for c in [col_name, "analysis_notes"] if c in flagged.columns]
    print(flagged[cols].to_string(index=False, max_colwidth=80))


# ── Excel Export (xlsxwriter — streams, fast on 100K+ rows) ────────────────────

# Color palette (no leading '#' because xlsxwriter accepts both forms)
_C = {
    "header_bg":   "#1F3864",
    "header_fg":   "#FFFFFF",
    "flag_red":    "#FFB3B3",
    "row_alt":     "#F2F4F8",
    "blue_accent": "#CCE5FF",
    "yes_text":    "#990000",
}

_FRIENDLY_NAMES: dict[str, str] = {
    "flag_suffix_mismatch":            "Name Mismatch",
    "flag_legacy_disagrees":           "Legacy ERP Mismatch",
    "flag_name_legacy_conflict":       "Name ↔ Legacy Conflict",
    "flag_catsy_disagrees":            "Catsy PIM Mismatch",
    "flag_name_catsy_conflict":        "Name ↔ Catsy Conflict",
    "flag_class_material_mismatch":    "Class Mismatch",
    "flag_empty_material_composition": "Missing Composition",
    "flag_is_matrix_parent":           "Matrix Parent",
    "flag_is_bb_part":                 "BB Part (Exempt)",
    "flag_unknown_composition":        "Unknown Composition",
    "any_flag":                        "Has Issue",
    "legacy_material":                 "Legacy ERP Material",
    "catsy_material":                  "Catsy PIM Material",
    "recommended_material":            "Recommended Material",
    "expected_material_from_name":     "Suffix-Detected Material",
    "matched_signal":                  "Matched Signal",
    "confidence":                      "Confidence",
    "analysis_notes":                  "Analysis Notes",
    "suggested_fix":                   "Suggested Fix",
}


def export_excel(df: pd.DataFrame, output_path: str, progress_callback=None) -> None:
    """
    Stream a formatted .xlsx report to disk using xlsxwriter.

    Optimizations vs. the previous openpyxl version:
      • constant_memory mode: each row flushed immediately → constant memory,
        no in-memory cell graph
      • per-cell formatting eliminated; row colors and YES highlighting are
        applied via Excel conditional formatting rules instead
      • boolean flag columns are converted to "YES"/"—" strings in pandas
        (vectorized) before writing, so no second pass is needed
    """
    if not EXCEL_AVAILABLE:
        print(
            "\nERROR: xlsxwriter is not installed.\n"
            "Install it with:  pip install xlsxwriter\n"
        )
        sys.exit(1)

    def report(pct: float, msg: str) -> None:
        if progress_callback:
            try:
                progress_callback(pct, msg)
            except Exception:
                pass

    # ── Column ordering ────────────────────────────────────────────────────
    derived_cols   = ["legacy_material", "catsy_material",
                      "recommended_material",
                      "expected_material_from_name", "matched_signal",
                      "confidence", "suggested_fix", "analysis_notes"]
    source_cols = [
        c for c in df.columns
        if c not in FLAG_COLS + ["any_flag"] + derived_cols
    ]
    flag_disp_cols = [c for c in FLAG_COLS if c in df.columns] + ["any_flag"]
    ordered_cols   = source_cols + derived_cols + flag_disp_cols

    # ── Convert boolean flag columns to display strings (vectorized) ──────
    report(0.02, "Preparing display data…")
    display_df = df[ordered_cols].copy()
    for c in flag_disp_cols:
        display_df[c] = display_df[c].map({True: "YES", False: "—"}).fillna("—")
    # Ensure no NaN/Inf reaches xlsxwriter (which rejects them by default).
    # Object columns → empty string; numeric NaN/±Inf → empty string too,
    # which forces the column to object dtype but is harmless.
    import numpy as _np
    for c in display_df.columns:
        col = display_df[c]
        if col.dtype == object:
            display_df[c] = col.fillna("").astype(str)
        else:
            display_df[c] = (
                col.replace([_np.inf, -_np.inf], _np.nan)
                   .where(col.notna(), "")
            )

    # Mask of issue rows (drives conditional row-fill rule)
    issue_mask = df["any_flag"].astype(bool).values

    # ── Sheet plan: Summary + a single Detail sheet (with autofilter so the
    # user can isolate Has Issue=YES or any individual flag column) ────────
    sheet_plan: list[tuple[str, pd.DataFrame, "any"]] = [
        ("Detail", display_df, issue_mask),
    ]

    total_rows = sum(len(sub) for _, sub, _ in sheet_plan) or 1
    written = [0]

    def on_chunk_written(chunk: int) -> None:
        written[0] += chunk
        if progress_callback:
            pct = min(0.99, 0.05 + 0.92 * (written[0] / total_rows))
            report(pct, f"Writing Excel… {written[0]:,} of {total_rows:,} rows")

    # ── Workbook + reusable formats ────────────────────────────────────────
    report(0.04, "Initializing workbook…")
    wb = xlsxwriter.Workbook(output_path, {
        "constant_memory":    True,
        "nan_inf_to_errors":  True,    # graceful fallback if a NaN slips through
    })

    fmt = {
        "title":      wb.add_format({
            "bold": True, "font_size": 15, "font_color": _C["header_fg"],
            "bg_color": _C["header_bg"], "align": "center", "valign": "vcenter"
        }),
        "subtle":     wb.add_format({
            "italic": True, "font_size": 10, "font_color": "#555555", "align": "center"
        }),
        "header":     wb.add_format({
            "bold": True, "font_color": _C["header_fg"], "bg_color": _C["header_bg"],
            "align": "center", "valign": "vcenter", "text_wrap": True, "font_size": 10
        }),
        "section":    wb.add_format({
            "bold": True, "bg_color": _C["blue_accent"], "valign": "vcenter"
        }),
        "label_bold": wb.add_format({"bold": True, "valign": "vcenter"}),
        "value":      wb.add_format({"valign": "vcenter"}),
        "notes":      wb.add_format({"text_wrap": True, "valign": "vcenter"}),
        "center":     wb.add_format({"align": "center", "valign": "vcenter"}),
        # Conditional-format formats (no other styling — applied on top of cell)
        "red_row":    wb.add_format({"bg_color": _C["flag_red"]}),
        "yes_cell":   wb.add_format({
            "bold": True, "font_color": _C["yes_text"], "bg_color": _C["flag_red"],
            "align": "center"
        }),
    }

    _write_summary_sheet_xlsx(wb, df, fmt)

    for sheet_name, sub_df, _ in sheet_plan:
        ws = wb.add_worksheet(sheet_name)
        _write_data_sheet_xlsx(ws, sub_df, fmt, on_chunk_written)

    report(0.99, "Saving file…")
    wb.close()
    report(1.0, "Done")
    print(f"\nExcel report saved to: {output_path}")


def _write_summary_sheet_xlsx(wb, df: pd.DataFrame, fmt: dict) -> None:
    ws = wb.add_worksheet("Summary")
    ws.set_column(0, 0, 50)
    ws.set_column(1, 1, 18)

    # Title banner
    ws.merge_range(0, 0, 0, 1, "NetSuite Material Composition Analysis", fmt["title"])
    ws.set_row(0, 36)

    # Provenance line: when, by whom, which version
    try:
        user = os.environ.get("USERNAME") or os.environ.get("USER") or "unknown"
    except Exception:
        user = "unknown"
    ws.merge_range(
        1, 0, 1, 1,
        f"Generated: {datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}   "
        f"by: {user}   analyzer v{__version__}",
        fmt["subtle"]
    )
    ws.set_row(1, 18)

    total   = len(df)
    flagged = int(df["any_flag"].sum())

    summary_rows = [
        ("Total Records Analyzed",  f"{total:,}"),
        ("Records with Issues",     f"{flagged:,}"),
        ("Clean Records",           f"{total - flagged:,}"),
        ("Issue Rate",              f"{flagged / total * 100:.1f}%" if total else "—"),
    ]
    # Surface legacy match coverage when a legacy ERP cross-check was applied
    if "legacy_material" in df.columns:
        legacy_matches = int((df["legacy_material"] != "").sum())
        if legacy_matches > 0:
            cov = legacy_matches / total * 100 if total else 0.0
            summary_rows.append(
                ("Legacy ERP Matches", f"{legacy_matches:,}  ({cov:.1f}%)")
            )
    # Same for Catsy PIM cross-check
    if "catsy_material" in df.columns:
        catsy_matches = int((df["catsy_material"] != "").sum())
        if catsy_matches > 0:
            cov = catsy_matches / total * 100 if total else 0.0
            summary_rows.append(
                ("Catsy PIM Matches", f"{catsy_matches:,}  ({cov:.1f}%)")
            )
    summary_rows += [
        ("", ""),
        ("Issue Breakdown (a single row may be counted in multiple lines)", "Count"),
    ]
    for flag_col, label in FLAG_META.items():
        count = int(df[flag_col].sum()) if flag_col in df.columns else 0
        summary_rows.append((label, f"{count:,}"))

    # Start writing at row 3 (0-indexed), leaving row 2 as a spacer
    out_row = 3
    bold_keys = {"Total Records Analyzed", "Records with Issues",
                 "Clean Records", "Issue Rate",
                 "Legacy ERP Matches", "Catsy PIM Matches"}
    for label, value in summary_rows:
        if label.startswith("Issue Breakdown"):
            ws.write(out_row, 0, label, fmt["section"])
            ws.write(out_row, 1, value, fmt["section"])
        elif label in bold_keys:
            ws.write(out_row, 0, label, fmt["label_bold"])
            ws.write(out_row, 1, value, fmt["value"])
        else:
            ws.write(out_row, 0, label, fmt["value"])
            ws.write(out_row, 1, value, fmt["value"])
        ws.set_row(out_row, 18)
        out_row += 1


def _write_data_sheet_xlsx(ws, df: pd.DataFrame, fmt: dict, on_chunk_written) -> None:
    """
    Write a DataFrame to a worksheet using xlsxwriter.

    Hot path: ws.write_row(row, 0, values) per row — no per-cell formatting.
    Visual highlighting is done by Excel-side conditional formatting rules
    keyed on the "any_flag" / flag column display strings (YES / —).
    """
    headers = [_FRIENDLY_NAMES.get(c, c) for c in df.columns]
    col_names = list(df.columns)
    n_rows = len(df)
    n_cols = len(col_names)

    # Header
    ws.write_row(0, 0, headers, fmt["header"])
    ws.set_row(0, 28)
    ws.freeze_panes(1, 0)

    # Per-column widths and column-default formats
    flag_positions = {i for i, c in enumerate(col_names) if c.startswith("flag_") or c == "any_flag"}
    notes_pos = col_names.index("analysis_notes") if "analysis_notes" in col_names else None
    any_flag_pos = col_names.index("any_flag") if "any_flag" in col_names else None

    sample_size = min(n_rows, 500)
    for col_idx, c in enumerate(col_names):
        if c == "analysis_notes":
            ws.set_column(col_idx, col_idx, 52, fmt["notes"])
        elif c == "suggested_fix":
            ws.set_column(col_idx, col_idx, 60, fmt["notes"])
        elif col_idx in flag_positions:
            ws.set_column(col_idx, col_idx, 16, fmt["center"])
        elif c in ("recommended_material", "expected_material_from_name",
                   "legacy_material", "catsy_material"):
            ws.set_column(col_idx, col_idx, 26)
        elif c == "matched_signal":
            ws.set_column(col_idx, col_idx, 45, fmt["notes"])
        elif c == "confidence":
            ws.set_column(col_idx, col_idx, 18, fmt["center"])
        else:
            if n_rows > 0:
                lengths = df[c].head(sample_size).astype(str).str.len()
                max_content = int(lengths.max()) if len(lengths) else 0
            else:
                max_content = 0
            header_len = len(_FRIENDLY_NAMES.get(c, c))
            width = min(max(max_content, header_len) + 3, 42)
            ws.set_column(col_idx, col_idx, width)

    # Data rows — pure write_row, no formatting overhead
    CHUNK = 5000
    chunk_count = 0
    for row_idx, values in enumerate(df.itertuples(index=False, name=None), 1):
        ws.write_row(row_idx, 0, values)
        chunk_count += 1
        if chunk_count >= CHUNK:
            if on_chunk_written:
                on_chunk_written(chunk_count)
            chunk_count = 0
    if chunk_count and on_chunk_written:
        on_chunk_written(chunk_count)

    # Conditional formatting (one rule covers all data rows) ────────────────
    if n_rows > 0:
        last_col = xl_col_to_name(n_cols - 1)
        full_range = f"A2:{last_col}{n_rows + 1}"

        # Highlight whole row when its any_flag cell == "YES"
        if any_flag_pos is not None:
            af_col = xl_col_to_name(any_flag_pos)
            ws.conditional_format(full_range, {
                "type":     "formula",
                "criteria": f'=${af_col}2="YES"',
                "format":   fmt["red_row"],
            })

        # Highlight individual YES cells in every flag column
        for fp in flag_positions:
            col_letter = xl_col_to_name(fp)
            cell_range = f"{col_letter}2:{col_letter}{n_rows + 1}"
            ws.conditional_format(cell_range, {
                "type":     "cell",
                "criteria": "==",
                "value":    '"YES"',
                "format":   fmt["yes_cell"],
            })

        # Autofilter on all columns
        ws.autofilter(0, 0, n_rows, n_cols - 1)


# ── CLI ────────────────────────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="analyze_materials",
        description="Analyze NetSuite material composition data from a CSV saved-search export.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python analyze_materials.py items.csv
  python analyze_materials.py items.csv --excel report.xlsx
  python analyze_materials.py items.csv --excel report.xlsx --flagged-only
        """,
    )
    p.add_argument("input",
        help="Path to the CSV file exported from NetSuite"
    )
    p.add_argument("--excel", "-e", metavar="OUTPUT.xlsx",
        help="Also export results to a formatted Excel workbook"
    )
    p.add_argument("--flagged-only", "-f", action="store_true",
        help="Print only flagged records in the terminal output"
    )
    p.add_argument("--encoding", default="utf-8-sig",
        help="CSV file encoding (default: utf-8-sig — handles Excel BOM automatically)"
    )
    p.add_argument(
        "--checks", "-c",
        nargs="+",
        choices=["suffix", "class", "empty", "legacy", "catsy", "all"],
        default=["all"],
        help=(
            "Which checks to run (default: all). "
            "'suffix' = name (suffix or pattern) ↔ material; "
            "'legacy' = legacy ERP cross-check (requires --legacy-csv); "
            "'catsy'  = Catsy PIM cross-check (requires --catsy-csv); "
            "'class' = class names a different material; "
            "'empty' = missing composition. "
            "Example: --checks suffix legacy catsy"
        ),
    )
    p.add_argument(
        "--legacy-csv",
        metavar="LEGACY.csv",
        help="Optional legacy ERP CSV with Part Number + Material columns. "
             "When provided, each NetSuite row is matched against External ID "
             "(falling back to Name) and the legacy material is used as an "
             "authoritative cross-check.",
    )
    p.add_argument(
        "--catsy-csv",
        metavar="CATSY.csv",
        help="Optional Catsy PIM CSV with 'Items' + 'Primary Material' columns. "
             "Each NetSuite row is matched by Name; the Catsy material is used "
             "as a parallel tier-2 cross-reference (alongside legacy ERP).",
    )
    return p


# Maps user-friendly check names → the set of underlying flag columns it
# enables. Some check names (legacy, catsy) cover two related flags — the
# main disagreement flag PLUS the name-vs-source conflict flag — because
# users think of them as one check, not two.
CHECK_NAME_TO_FLAG: dict[str, set[str]] = {
    "suffix": {"flag_suffix_mismatch"},
    "class":  {"flag_class_material_mismatch"},
    "empty":  {"flag_empty_material_composition"},
    "legacy": {"flag_legacy_disagrees", "flag_name_legacy_conflict"},
    "catsy":  {"flag_catsy_disagrees",  "flag_name_catsy_conflict"},
}


def resolve_checks(names: list[str]) -> set[str]:
    """Convert a list like ['suffix','class'] or ['all'] → set of flag column keys."""
    if not names or "all" in names:
        return set(ISSUE_FLAGS.keys())
    out: set[str] = set()
    for n in names:
        if n in CHECK_NAME_TO_FLAG:
            out |= CHECK_NAME_TO_FLAG[n]
    return out


def load_csv(path: Path, encoding: str) -> pd.DataFrame:
    """Load CSV with automatic encoding fallback."""
    try:
        df = pd.read_csv(path, encoding=encoding, dtype=str)
        print(f"Loaded {len(df):,} rows from {path}")
        return df
    except UnicodeDecodeError:
        print(f"  utf-8 decode failed — retrying with latin-1...")
        df = pd.read_csv(path, encoding="latin-1", dtype=str)
        print(f"Loaded {len(df):,} rows from {path} (latin-1)")
        return df


def run_cli(args) -> None:
    csv_path = Path(args.input)
    if not csv_path.exists():
        print(f"ERROR: File not found: {csv_path}")
        sys.exit(1)

    df_raw = load_csv(csv_path, args.encoding)
    print(f"Columns: {list(df_raw.columns)}\n")

    legacy_lookup = None
    if args.legacy_csv:
        legacy_path = Path(args.legacy_csv)
        if not legacy_path.exists():
            print(f"ERROR: Legacy ERP CSV not found: {legacy_path}")
            sys.exit(1)
        legacy_df = load_csv(legacy_path, args.encoding)
        try:
            legacy_lookup = build_legacy_lookup(legacy_df)
            print(f"Loaded {len(legacy_lookup):,} legacy ERP entries\n")
        except ValueError as exc:
            print(f"ERROR: {exc}")
            sys.exit(1)

    catsy_lookup = None
    if args.catsy_csv:
        catsy_path = Path(args.catsy_csv)
        if not catsy_path.exists():
            print(f"ERROR: Catsy CSV not found: {catsy_path}")
            sys.exit(1)
        catsy_df = load_csv(catsy_path, args.encoding)
        try:
            catsy_lookup = build_catsy_lookup(catsy_df)
            print(f"Loaded {len(catsy_lookup):,} Catsy PIM entries\n")
        except ValueError as exc:
            print(f"ERROR: {exc}")
            sys.exit(1)

    try:
        df_result = analyze_dataframe(
            df_raw,
            enabled_checks=resolve_checks(args.checks),
            legacy_lookup=legacy_lookup,
            catsy_lookup=catsy_lookup,
        )
    except ValueError as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)

    col_map  = {c.lower(): c for c in df_result.columns}
    col_name = next(
        (col_map[k] for k in ("name", "item name", "item_name") if k in col_map),
        df_result.columns[0]
    )

    print_summary(df_result)

    if args.flagged_only:
        print_flagged_details(df_result, col_name)
    else:
        flagged_count = int(df_result["any_flag"].sum())
        if flagged_count:
            print_flagged_details(df_result, col_name)

    if args.excel:
        export_excel(df_result, args.excel)


# ── GUI ────────────────────────────────────────────────────────────────────────

def run_gui() -> None:
    """Simple tkinter GUI: pick CSV, click Analyze, get an Excel report."""
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import threading
    import subprocess
    import os

    root = tk.Tk()
    root.title("Material Composition Analyzer")
    root.geometry("760x860")
    root.minsize(680, 760)

    state = {
        "csv_path": None, "xlsx_path": None,
        "legacy_path": None, "catsy_path": None,
    }

    def truncate_filename(name: str, max_len: int = 56) -> str:
        """Middle-ellipsis: 'long-filename-export.csv' → 'long-fi…rt.csv'.

        Keeps a useful prefix + the suffix (including extension) visible so
        the user can still recognize the file without the label getting clipped.
        """
        if len(name) <= max_len:
            return name
        keep = max_len - 1               # 1 char for the ellipsis
        head = keep * 2 // 3
        tail = keep - head
        return f"{name[:head]}…{name[-tail:]}"

    # ── Layout ─────────────────────────────────────────────────────────────
    main_frame = ttk.Frame(root, padding=20)
    main_frame.pack(fill="both", expand=True)

    title_lbl = ttk.Label(
        main_frame,
        text="NetSuite Material Composition Analyzer",
        font=("Segoe UI", 14, "bold"),
    )
    title_lbl.pack(anchor="w")

    subtitle_lbl = ttk.Label(
        main_frame,
        text="Select a CSV exported from your NetSuite saved search, then click Analyze.",
        foreground="#555555",
    )
    subtitle_lbl.pack(anchor="w", pady=(2, 16))

    # File picker row
    file_frame = ttk.LabelFrame(main_frame, text="Input CSV", padding=10)
    file_frame.pack(fill="x")

    file_lbl = ttk.Label(file_frame, text="(no file selected)", foreground="#888888")
    file_lbl.pack(side="left", fill="x", expand=True)

    def pick_file():
        path = filedialog.askopenfilename(
            title="Select NetSuite CSV export",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            state["csv_path"] = path
            file_lbl.config(text=truncate_filename(Path(path).name), foreground="#000000")
            analyze_btn.config(state="normal")
            status_lbl.config(text="Ready to analyze.", foreground="#000000")

    pick_btn = ttk.Button(file_frame, text="Choose CSV…", command=pick_file)
    pick_btn.pack(side="right", padx=(10, 0))

    # ── Optional legacy ERP cross-check ─────────────────────────────────────
    legacy_frame = ttk.LabelFrame(
        main_frame, text="Legacy ERP Cross-Check (optional)", padding=10
    )
    legacy_frame.pack(fill="x", pady=(8, 0))

    legacy_lbl = ttk.Label(
        legacy_frame, text="(no file selected)", foreground="#888888"
    )
    legacy_lbl.pack(side="left", fill="x", expand=True)

    def clear_legacy():
        state["legacy_path"] = None
        legacy_lbl.config(text="(no file selected)", foreground="#888888")

    def pick_legacy():
        path = filedialog.askopenfilename(
            title="Select Legacy ERP CSV (Part Number + Material columns)",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            state["legacy_path"] = path
            legacy_lbl.config(text=truncate_filename(Path(path).name), foreground="#000000")

    legacy_clear_btn = ttk.Button(legacy_frame, text="Clear", command=clear_legacy)
    legacy_clear_btn.pack(side="right", padx=(6, 0))
    legacy_pick_btn = ttk.Button(
        legacy_frame, text="Choose Legacy CSV…", command=pick_legacy
    )
    legacy_pick_btn.pack(side="right", padx=(10, 0))

    # ── Optional Catsy PIM cross-check ─────────────────────────────────────
    catsy_frame = ttk.LabelFrame(
        main_frame, text="Catsy PIM Cross-Check (optional)", padding=10
    )
    catsy_frame.pack(fill="x", pady=(8, 0))

    catsy_lbl = ttk.Label(
        catsy_frame, text="(no file selected)", foreground="#888888"
    )
    catsy_lbl.pack(side="left", fill="x", expand=True)

    def clear_catsy():
        state["catsy_path"] = None
        catsy_lbl.config(text="(no file selected)", foreground="#888888")

    def pick_catsy():
        path = filedialog.askopenfilename(
            title="Select Catsy CSV (Items + Primary Material columns)",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            state["catsy_path"] = path
            catsy_lbl.config(text=truncate_filename(Path(path).name), foreground="#000000")

    catsy_clear_btn = ttk.Button(catsy_frame, text="Clear", command=clear_catsy)
    catsy_clear_btn.pack(side="right", padx=(6, 0))
    catsy_pick_btn = ttk.Button(
        catsy_frame, text="Choose Catsy CSV…", command=pick_catsy
    )
    catsy_pick_btn.pack(side="right", padx=(10, 0))

    # ── Checks scope ───────────────────────────────────────────────────────
    checks_frame = ttk.LabelFrame(main_frame, text="Checks to Run", padding=10)
    checks_frame.pack(fill="x", pady=(12, 0))

    check_vars: dict[str, tk.BooleanVar] = {
        "suffix": tk.BooleanVar(value=True),
        "class":  tk.BooleanVar(value=True),
        "empty":  tk.BooleanVar(value=True),
        "catsy":  tk.BooleanVar(value=True),
    }
    check_labels = {
        "suffix": "Name suffix ↔ material composition",
        "class":  "Class reflects material composition",
        "empty":  "Missing material composition",
        "catsy":  "Catsy PIM cross-check (needs Catsy CSV)",
    }

    for key, label in check_labels.items():
        ttk.Checkbutton(checks_frame, text=label, variable=check_vars[key]).pack(
            anchor="w"
        )

    # Quick scope buttons
    scope_btns = ttk.Frame(checks_frame)
    scope_btns.pack(anchor="w", pady=(8, 0))

    def set_scope(*keys):
        for k, var in check_vars.items():
            var.set(k in keys)

    ttk.Button(scope_btns, text="All",
               command=lambda: set_scope("suffix", "class", "empty", "catsy"),
               width=10).pack(side="left")
    ttk.Button(scope_btns, text="Suffix only",
               command=lambda: set_scope("suffix"),
               width=12).pack(side="left", padx=(6, 0))
    ttk.Button(scope_btns, text="Class only",
               command=lambda: set_scope("class"),
               width=12).pack(side="left", padx=(6, 0))
    ttk.Button(scope_btns, text="Catsy only",
               command=lambda: set_scope("catsy"),
               width=12).pack(side="left", padx=(6, 0))

    # ── Layout: pack the action buttons to the BOTTOM first so they're
    # always visible no matter how cramped the rest of the window gets.
    # Then pack the progress bar / label just above (when shown), and let
    # the status area expand to fill the remaining middle space.

    btn_frame = ttk.Frame(main_frame)
    btn_frame.pack(side="bottom", fill="x", pady=(12, 0))

    # Determinate progress bar with live percentage / phase message label
    # (created here, packed dynamically just above btn_frame when analyze runs)
    progress_lbl = ttk.Label(main_frame, text="", foreground="#444444",
                              font=("Segoe UI", 9))
    progress = ttk.Progressbar(main_frame, mode="determinate", maximum=100)

    # Status / summary area — fills the middle, pushes against btn_frame
    status_frame = ttk.LabelFrame(main_frame, text="Status", padding=10)
    status_frame.pack(side="top", fill="both", expand=True, pady=(12, 0))

    status_lbl = ttk.Label(
        status_frame,
        text="Choose a CSV file to begin.",
        foreground="#555555",
    )
    status_lbl.pack(anchor="w")

    summary_text = tk.Text(
        status_frame,
        height=12,
        wrap="word",
        state="disabled",
        font=("Consolas", 10),
        background="#f7f7f7",
        relief="flat",
    )
    summary_text.pack(fill="both", expand=True, pady=(8, 0))

    def open_path(path: str):
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["xdg-open", path], check=False)
        except Exception as exc:
            messagebox.showerror("Couldn't open", str(exc))

    def with_loading_bar(message: str, action, duration_ms: int = 2500):
        """Reuse the progress bar in indeterminate mode for OS-level launches.

        The OS handles the actual work (opening Excel / a file manager) in
        the background and we get no completion signal back. We just show
        an animated bar for `duration_ms` so the user has visible feedback
        that something is happening, then hide it again. Buttons that
        trigger this are disabled for the same window to prevent runaway
        double-clicks.
        """
        # Disable Open buttons during the launch window
        open_file_btn.config(state="disabled")
        open_folder_btn.config(state="disabled")

        progress.config(mode="indeterminate", value=0)
        progress_lbl.config(text=message)
        progress_lbl.pack(side="bottom", fill="x")
        progress.pack(side="bottom", fill="x", pady=(10, 4))
        progress.start(15)        # ms per tick

        # Kick off the OS-level action in a thread so the UI stays responsive
        threading.Thread(target=action, daemon=True).start()

        def cleanup():
            progress.stop()
            progress.config(mode="determinate", value=0)
            progress.pack_forget()
            progress_lbl.pack_forget()
            # Re-enable Open buttons only if we still have a report
            if state["xlsx_path"]:
                open_file_btn.config(state="normal")
                open_folder_btn.config(state="normal")
        root.after(duration_ms, cleanup)

    def do_open_excel():
        if not state["xlsx_path"]:
            return
        with_loading_bar(
            "Opening Excel report…",
            lambda: open_path(state["xlsx_path"]),
            duration_ms=2800,
        )

    def do_open_folder():
        if not state["xlsx_path"]:
            return
        with_loading_bar(
            "Opening folder…",
            lambda: open_path(str(Path(state["xlsx_path"]).parent)),
            duration_ms=1200,
        )

    open_file_btn = ttk.Button(
        btn_frame,
        text="Open Excel Report",
        command=do_open_excel,
        state="disabled",
    )
    open_folder_btn = ttk.Button(
        btn_frame,
        text="Open Folder",
        command=do_open_folder,
        state="disabled",
    )

    def set_summary(text: str):
        summary_text.config(state="normal")
        summary_text.delete("1.0", "end")
        summary_text.insert("1.0", text)
        summary_text.config(state="disabled")

    def do_analyze():
        csv_path = state["csv_path"]
        if not csv_path:
            return

        analyze_btn.config(state="disabled")
        pick_btn.config(state="disabled")
        open_file_btn.config(state="disabled")
        open_folder_btn.config(state="disabled")
        # Pack progress widgets to the BOTTOM (just above btn_frame).
        # Order matters with side=bottom — packed-first ends up lowest:
        # we want label below the bar, so pack label first, then bar.
        progress_lbl.pack(side="bottom", fill="x")
        progress.pack(side="bottom", fill="x", pady=(10, 4))
        progress.config(value=0)
        progress_lbl.config(text="Starting…")
        status_lbl.config(text="Analyzing…", foreground="#000000")
        set_summary("")

        # Snapshot the user's check selection at click time
        selected = [k for k, v in check_vars.items() if v.get()]
        if not selected:
            messagebox.showwarning(
                "No checks selected",
                "Pick at least one check to run."
            )
            analyze_btn.config(state="normal")
            pick_btn.config(state="normal")
            return
        enabled = resolve_checks(selected)

        # Filename suffix reflecting scope: "all" or e.g. "suffix" or "suffix_class"
        is_all = (set(selected) == set(check_vars.keys()))
        scope_tag = "all" if is_all else "_".join(selected)

        # Throttled UI updater — schedules at most one redraw per ~50 ms
        last_pct_pushed = [-1.0]
        def push_progress(global_pct: float, msg: str) -> None:
            # Only marshal to the UI thread when the value actually changes
            if abs(global_pct - last_pct_pushed[0]) < 0.005 and global_pct < 1.0:
                return
            last_pct_pushed[0] = global_pct
            def update():
                progress.config(value=global_pct * 100)
                progress_lbl.config(text=f"{msg}   ({global_pct * 100:.0f}%)")
            root.after(0, update)

        def make_phase_cb(start: float, end: float):
            span = end - start
            def cb(local_pct: float, msg: str) -> None:
                push_progress(start + span * local_pct, msg)
            return cb

        # Snapshot cross-reference paths at click time too
        legacy_path = state.get("legacy_path")
        catsy_path  = state.get("catsy_path")

        def worker():
            try:
                csv_p = Path(csv_path)
                push_progress(0.01, "Reading CSV…")
                df_raw = load_csv(csv_p, "utf-8-sig")
                push_progress(0.04, f"Loaded {len(df_raw):,} rows")

                # Load legacy ERP lookup if a path was picked
                legacy_lookup = None
                legacy_match_count = 0
                if legacy_path:
                    push_progress(0.05, "Reading legacy ERP CSV…")
                    legacy_df = load_csv(Path(legacy_path), "utf-8-sig")
                    legacy_lookup = build_legacy_lookup(legacy_df)
                    push_progress(
                        0.055,
                        f"Legacy ERP: {len(legacy_lookup):,} entries loaded"
                    )

                # Load Catsy PIM lookup if a path was picked
                catsy_lookup = None
                catsy_match_count = 0
                if catsy_path:
                    push_progress(0.058, "Reading Catsy PIM CSV…")
                    catsy_df = load_csv(Path(catsy_path), "utf-8-sig")
                    catsy_lookup = build_catsy_lookup(catsy_df)
                    push_progress(
                        0.06,
                        f"Catsy PIM: {len(catsy_lookup):,} entries loaded"
                    )

                df_result = analyze_dataframe(
                    df_raw,
                    enabled_checks=enabled,
                    progress_callback=make_phase_cb(0.06, 0.35),
                    legacy_lookup=legacy_lookup,
                    catsy_lookup=catsy_lookup,
                )

                if legacy_lookup is not None:
                    legacy_match_count = int((df_result["legacy_material"] != "").sum())
                if catsy_lookup is not None:
                    catsy_match_count = int((df_result["catsy_material"] != "").sum())

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                xlsx_p = csv_p.with_name(
                    f"{csv_p.stem}_analysis_{scope_tag}_{timestamp}.xlsx"
                )
                export_excel(
                    df_result,
                    str(xlsx_p),
                    progress_callback=make_phase_cb(0.35, 1.0),
                )

                total   = len(df_result)
                flagged = int(df_result["any_flag"].sum())
                pct     = flagged / total * 100 if total else 0.0

                lines = [
                    f"Analyzed:       {total:,} records",
                    f"With issues:    {flagged:,}  ({pct:.1f}%)",
                    f"Clean:          {total - flagged:,}",
                ]
                if legacy_lookup is not None:
                    pct_match = legacy_match_count / total * 100 if total else 0.0
                    lines.append(
                        f"Legacy match:   {legacy_match_count:,} of {total:,} "
                        f"({pct_match:.1f}%)"
                    )
                if catsy_lookup is not None:
                    pct_match = catsy_match_count / total * 100 if total else 0.0
                    lines.append(
                        f"Catsy match:    {catsy_match_count:,} of {total:,} "
                        f"({pct_match:.1f}%)"
                    )
                lines += ["", "Issue breakdown:"]
                for col, label in FLAG_META.items():
                    if col in df_result.columns:
                        cnt = int(df_result[col].sum())
                        marker = "  !" if cnt > 0 else "   "
                        lines.append(f"{marker} {label:<48} {cnt:>6,}")
                lines += ["", f"Saved to:  {xlsx_p.name}"]

                state["xlsx_path"] = str(xlsx_p)

                def on_done():
                    progress.config(value=100)
                    progress_lbl.config(text="Complete   (100%)")
                    progress.pack_forget()
                    progress_lbl.pack_forget()
                    set_summary("\n".join(lines))
                    status_lbl.config(text="Done.", foreground="#006600")
                    analyze_btn.config(state="normal")
                    pick_btn.config(state="normal")
                    open_file_btn.config(state="normal")
                    open_folder_btn.config(state="normal")
                root.after(0, on_done)

            except Exception as exc:
                err_msg = str(exc)
                def on_err():
                    progress.pack_forget()
                    progress_lbl.pack_forget()
                    status_lbl.config(text="Error.", foreground="#990000")
                    set_summary(f"ERROR:\n\n{err_msg}")
                    analyze_btn.config(state="normal")
                    pick_btn.config(state="normal")
                root.after(0, on_err)

        threading.Thread(target=worker, daemon=True).start()

    analyze_btn = ttk.Button(
        btn_frame,
        text="Analyze and Save Excel",
        command=do_analyze,
        state="disabled",
    )
    analyze_btn.pack(side="left")
    open_file_btn.pack(side="left", padx=(8, 0))
    open_folder_btn.pack(side="left", padx=(8, 0))

    quit_btn = ttk.Button(btn_frame, text="Close", command=root.destroy)
    quit_btn.pack(side="right")

    root.mainloop()


def main() -> None:
    # No CLI args → launch GUI (this is what double-clicking the .exe does).
    if len(sys.argv) <= 1:
        run_gui()
        return

    # Otherwise, parse CLI args as before.
    args = build_parser().parse_args()
    run_cli(args)


if __name__ == "__main__":
    main()
