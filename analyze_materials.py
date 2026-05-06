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

__version__ = "1.7.0"

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
# Any row that produces an analysis note belongs here, so the count of
# Has Issue == YES rows always matches the count of non-OK notes.
ISSUE_FLAGS: dict[str, str] = {
    "flag_legacy_disagrees":           "Legacy ERP Disagrees with Material Composition",
    "flag_suffix_mismatch":            "Name Suffix ↔ Material Mismatch",
    "flag_no_suffix_non_steel":        "No Material Suffix — Non-Steel Composition (Review)",
    "flag_class_material_mismatch":    "Class Doesn't Reflect Material",
    "flag_material_field_mismatch":    "Matrix Material Field Mismatch",
    "flag_empty_material_composition": "Missing Material Composition",
}

# Informational flags — context only, no analysis note, do NOT trigger any_flag.
INFO_FLAGS: dict[str, str] = {
    "flag_is_matrix_parent": "Matrix Parent (Multiple Compositions)",
    "flag_is_bb_part":       "Brennan Black Part (Class Check Exempt)",
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


def analyze_dataframe(
    df: pd.DataFrame,
    enabled_checks: set[str] | None = None,
    progress_callback=None,
    legacy_lookup: dict[str, str] | None = None,
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

    # ── Legacy ERP lookup ──────────────────────────────────────────────────
    # Match each row by External ID first, then by Name. legacy_material is
    # the cleaned legacy ERP material string (or "" when no match).
    if legacy_lookup:
        ext_lower  = ext_id.str.lower()
        name_lower_for_lookup = name.str.lower()
        legacy_via_ext  = ext_lower.map(legacy_lookup).fillna("")
        legacy_via_name = name_lower_for_lookup.map(legacy_lookup).fillna("")
        legacy_mat = legacy_via_ext.where(legacy_via_ext != "", legacy_via_name)
    else:
        legacy_mat = pd.Series([""] * n, index=idx, dtype=object)
    legacy_lower = legacy_mat.str.lower()
    has_legacy   = legacy_lower != ""

    report(0.10, "Parsing names…")

    # ── Vectorized name segment analysis ───────────────────────────────────
    segments = (
        name.str.upper()
            .str.split("-")
            .map(lambda s: [seg.strip() for seg in s if seg.strip()] if isinstance(s, list) else [])
    )
    has_ss        = segments.map(lambda s: "SS" in s)
    has_bb        = segments.map(lambda s: "BB" in s)
    has_d         = segments.map(lambda s: "D" in s)   # D anywhere as own segment (e.g. -D-GREEN)
    last_seg      = segments.map(lambda s: s[-1] if s else "")
    has_b_suffix  = (last_seg == "B")    # exact -B (not -BB)

    # Expected material from suffix (priority: SS > B > D)
    expected_mat = pd.Series([""] * n, index=idx, dtype=object)
    expected_mat = expected_mat.mask(has_ss, "Stainless Steel")
    expected_mat = expected_mat.mask(~has_ss & has_b_suffix, "Brass")
    expected_mat = expected_mat.mask(~has_ss & ~has_b_suffix & has_d, "Aluminum")

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

    # ── Class & name keyword extraction (used by class check and class_confirms) ─
    name_lower = name.str.lower()

    # Class signals
    contains_stainless = klass_lower.str.contains("stainless", na=False)
    contains_steel     = klass_lower.str.contains("steel",     na=False)
    contains_brass     = klass_lower.str.contains("brass",     na=False)
    # "aerospace" classes → treat as Aluminum-confirming (industry context)
    contains_alum      = klass_lower.str.contains(
        r"alumin(?:um|ium)|aerospace", regex=True, na=False
    )

    # Name signals — patterns picked to match the user's actual NetSuite
    # naming conventions seen in production exports:
    #   ALUM VENT, BRS, BR$, 30B (digit+B at end), 316/304 alloy codes,
    #   S/S, standalone S, etc.
    name_says_alum = name_lower.str.contains(r"\balum", regex=True, na=False)
    name_says_brass = name_lower.str.contains(
        r"\bbrass\b|\bbrs\b|\bbr\b|br$|\d[bB]$",
        regex=True, na=False
    )
    name_says_stainless = name_lower.str.contains(
        r"\bss\b|\bs/s\b|\b316l?\b|\b304l?\b|\bstainless\b"
        r"|(?:^|[^a-z0-9])s(?:[^a-z0-9]|$)",
        regex=True, na=False
    )
    name_says_steel = name_lower.str.contains(r"\bsteel\b", regex=True, na=False)

    # Class (or name) confirms the material composition?
    # Used to suppress *both* the reverse suffix flags AND the forward suffix
    # mismatch — if class/name explicitly name the actual composition, the
    # name's suffix is the misleading element, not the composition.
    is_alum_comp = matc_lower.isin(["aluminum", "aluminium", "aluminum alloy"])
    class_confirms = (
          ((matc_lower == "stainless steel") & (contains_stainless | name_says_stainless))
        | ((matc_lower == "steel") & (contains_steel | name_says_steel)
           & ~contains_stainless & ~name_says_stainless)
        | ((matc_lower == "brass") & (contains_brass | name_says_brass))
        | (is_alum_comp & (contains_alum | name_says_alum))
    )

    # Legacy ERP confirmation: the strongest possible signal.
    # Matches when the legacy material exactly matches the current composition
    # (case-insensitive). When this fires, all soft suffix flags are suppressed.
    legacy_confirms = has_legacy & (legacy_lower == matc_lower)
    class_confirms  = class_confirms | legacy_confirms
    # Generic substring fallback: any non-empty composition whose lowercased
    # form appears verbatim in the class name (e.g. "Nylon" in "...P.T.C - Nylon").
    # Only relevant for materials not in the known list above.
    known_mat = matc_lower.isin([
        "stainless steel", "steel", "brass", "aluminum", "aluminium", "aluminum alloy"
    ])
    if (~known_mat & (matc_lower != "")).any():
        # NOTE: using `m != "" and (m in c)` (not `m and m in c`) because
        # Python's `and` returns the first falsy operand — `"" and X` is `""`,
        # producing a mixed str/bool Series and blowing up `&` later.
        substring_confirm = pd.Series(
            [(m != "") and (m in c) for m, c in zip(matc_lower, klass_lower)],
            index=idx,
            dtype=bool,
        )
        class_confirms = class_confirms | (substring_confirm & ~known_mat)

    report(0.30, "Checking suffix rules…")

    # ── FLAG: empty material composition ───────────────────────────────────
    if "flag_empty_material_composition" in enabled_checks:
        flag_empty = is_empty
    else:
        flag_empty = pd.Series([False] * n, index=idx)

    # ── FLAG: legacy ERP disagrees with Material Composition ──────────────
    # This is the highest-confidence flag — when the legacy ERP says one
    # material and NetSuite has another, NetSuite is almost certainly wrong.
    if "flag_legacy_disagrees" in enabled_checks:
        flag_legacy_disagrees = (
            has_legacy & (matc_lower != "") & (legacy_lower != matc_lower)
        )
    else:
        flag_legacy_disagrees = pd.Series([False] * n, index=idx)

    # ── FLAG: suffix mismatch (bidirectional) ──────────────────────────────
    # Forward mismatch (suffix actively contradicts the composition) is now
    # SUPPRESSED when class/name/legacy confirms the actual composition —
    # in that case the suffix is the misleading element, not the composition.
    # Reverse mismatch (composition needs a suffix the name lacks) is
    # likewise suppressed when class confirms.
    expected_lower  = expected_mat.str.lower()
    forward_mismatch = (
        (expected_mat != "") & (matc_lower != "") & (expected_lower != matc_lower)
        & ~class_confirms
    )
    is_single = ~has_multi_comp & ~is_matrix_parent

    rev_ss_mismatch    = is_single & (matc_lower == "stainless steel") & ~has_ss & ~class_confirms
    rev_brass_mismatch = is_single & (matc_lower == "brass") & ~has_b_suffix & ~class_confirms
    rev_alum_mismatch  = is_single & is_alum_comp & ~has_d & ~class_confirms

    if "flag_suffix_mismatch" in enabled_checks:
        flag_suffix = forward_mismatch | rev_ss_mismatch | rev_brass_mismatch | rev_alum_mismatch
    else:
        flag_suffix = pd.Series([False] * n, index=idx)
        forward_mismatch    = pd.Series([False] * n, index=idx)
        rev_ss_mismatch     = pd.Series([False] * n, index=idx)
        rev_brass_mismatch  = pd.Series([False] * n, index=idx)
        rev_alum_mismatch   = pd.Series([False] * n, index=idx)

    report(0.50, "Checking class consistency…")

    # ── FLAG: class doesn't reflect material ───────────────────────────────
    flag_class = pd.Series([False] * n, index=idx)
    class_note_text = pd.Series([""] * n, index=idx, dtype=object)

    if "flag_class_material_mismatch" in enabled_checks:
        # Blank class is a separate data-quality issue, not a class-mapping
        # mismatch — exclude blank-class rows from this check so notes don't
        # say "Class '' does not reflect: 'X'".
        eligible = ~is_bb_flag & (matc_lower != "") & (klass_lower != "")

        # Vectorized single-composition path (the common case)
        single_eligible = eligible & ~has_multi_comp

        # (contains_* keyword masks were already computed above for class_confirms)

        # Stainless Steel composition: class needs "stainless"
        ss_mat = single_eligible & (matc_lower == "stainless steel")
        flag_class |= ss_mat & ~contains_stainless

        # Plain Steel composition: class needs "steel" but NOT "stainless"
        steel_mat = single_eligible & (matc_lower == "steel")
        flag_class |= steel_mat & (~contains_steel | contains_stainless)

        # Brass
        brass_mat = single_eligible & (matc_lower == "brass")
        flag_class |= brass_mat & ~contains_brass

        # Aluminum / Aluminium
        alum_mat = single_eligible & (
            (matc_lower == "aluminum") | (matc_lower == "aluminium") | (matc_lower == "aluminum alloy")
        )
        flag_class |= alum_mat & ~contains_alum

        # Generic substring fallback for unknown materials (single-comp only)
        known = (matc_lower.isin([
            "stainless steel", "steel", "brass", "aluminum", "aluminium", "aluminum alloy"
        ]))
        unknown_mat = single_eligible & ~known & (matc_lower != "")
        if unknown_mat.any():
            # row-wise unknown check (small subset)
            def _unknown_check(row):
                return matc_lower.at[row.name] not in klass_lower.at[row.name]
            unknown_mismatch = df[unknown_mat].apply(_unknown_check, axis=1)
            flag_class.loc[unknown_mismatch.index] |= unknown_mismatch.fillna(False)

        # Class note text — single-comp rows
        class_note_text = class_note_text.mask(
            flag_class & single_eligible,
            "Class '" + klass + "' does not reflect: '" + matc + "'"
        )

        # Multi-composition path: per-row apply (small subset)
        multi_eligible = eligible & has_multi_comp
        if multi_eligible.any():
            def _multi_check(row):
                comps = split_multi_composition(str(row[col_mat_comp]))
                bad = [c for c in comps if not material_matches_class(c, str(row[col_class]))]
                return (bool(bad), bad)

            sub = df[multi_eligible].apply(_multi_check, axis=1, result_type="reduce")
            multi_flag = sub.map(lambda t: t[0])
            multi_text = sub.map(
                lambda t: "Class '" + "" + "' does not reflect: " +
                          ", ".join(f"'{c}'" for c in t[1]) if t[0] else ""
            )
            # Re-include actual class value:
            for i, val in sub.items():
                ok, bad_list = val
                if ok:
                    multi_text.at[i] = (
                        f"Class '{klass.at[i]}' does not reflect: "
                        + ", ".join(f"'{c}'" for c in bad_list)
                    )
            flag_class.loc[multi_flag.index] |= multi_flag.fillna(False)
            class_note_text.loc[multi_text.index] = (
                class_note_text.loc[multi_text.index]
                .where(~multi_flag.fillna(False), multi_text)
            )

    report(0.65, "Checking matrix Material field…")

    # ── FLAG: matrix Material field mismatch ───────────────────────────────
    flag_matrix = pd.Series([False] * n, index=idx)
    if "flag_material_field_mismatch" in enabled_checks and col_material is not None:
        eligible = (mat_lower != "") & (matc_lower != "")

        # Single-composition: simple inequality
        single_mismatch = eligible & ~has_multi_comp & (mat_lower != matc_lower)
        flag_matrix |= single_mismatch

        # Multi-composition: membership check (small subset)
        multi_eligible = eligible & has_multi_comp
        if multi_eligible.any():
            def _multi_member(row):
                comps = [normalize(c) for c in split_multi_composition(str(row[col_mat_comp]))]
                return normalize(str(row[col_material])) not in comps
            multi_mismatch = df[multi_eligible].apply(_multi_member, axis=1)
            flag_matrix.loc[multi_mismatch.index] |= multi_mismatch.fillna(False)

    # ── ISSUE: no suffix → non-Steel composition (suppressed if suffix flag
    # already fires, OR the class explicitly confirms the material).
    # Treated as part of the suffix scope — gated by flag_suffix_mismatch in
    # enabled_checks because the two checks are conceptually paired (one for
    # rows WITH a contradicting suffix, one for rows WITHOUT a needed one).
    if "flag_suffix_mismatch" in enabled_checks:
        flag_no_suffix_non_steel = (
            (expected_mat == "") & (matc_lower != "") & (matc_lower != "steel")
            & ~flag_suffix & ~class_confirms
        )
    else:
        flag_no_suffix_non_steel = pd.Series([False] * n, index=idx)

    # any_flag = OR of all issue flags. Must include flag_no_suffix_non_steel
    # so the count of "Has Issue == YES" matches the count of non-OK notes.
    any_flag = (
        flag_legacy_disagrees | flag_suffix | flag_no_suffix_non_steel
        | flag_class | flag_matrix | flag_empty
    )

    # ── Recommended Material (final) ────────────────────────────────────────
    # Priority (most → least authoritative):
    #   1. Legacy ERP material (when present) — authoritative source
    #   2. Suffix-derived material when forward mismatch fires AND no other
    #      signal confirms — the suffix is positive evidence
    #   3. Current Material Composition — trust what's already there
    #   4. "Steel" — only when composition is genuinely empty
    recommended_mat = matc.where(matc != "", "Steel")
    recommended_mat = recommended_mat.mask(forward_mismatch, expected_mat)
    recommended_mat = recommended_mat.mask(has_legacy, legacy_mat)

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

    if "flag_suffix_mismatch" in enabled_checks:
        # Forward
        fwd_text = (
            "Name suffix implies '" + expected_mat
            + "' but Material Composition is '" + matc + "'"
        )
        parts = parts.mask(forward_mismatch, parts + SEP + fwd_text)

        # Reverse SS — name lacks -SS for a Stainless Steel composition,
        # and the class doesn't confirm Stainless either
        parts = parts.mask(
            rev_ss_mismatch,
            parts + SEP +
            "Material Composition is 'Stainless Steel' but name has no -SS suffix "
            "and class doesn't confirm — composition may actually be Steel"
        )

        # Reverse Brass
        parts = parts.mask(
            rev_brass_mismatch,
            parts + SEP +
            "Material Composition is 'Brass' but name does not end in -B "
            "and class doesn't confirm — composition may actually be Steel"
        )

        # Reverse Aluminum
        parts = parts.mask(
            rev_alum_mismatch,
            parts + SEP +
            "Material Composition is 'Aluminum' but name does not end in -D "
            "and class doesn't confirm — composition may actually be Steel"
        )

    if "flag_class_material_mismatch" in enabled_checks:
        parts = parts.mask(flag_class, parts + SEP + class_note_text)

    if "flag_material_field_mismatch" in enabled_checks and col_material is not None:
        matrix_text = (
            "Matrix Material '" + material + "' not in Material Composition '" + matc + "'"
        )
        parts = parts.mask(flag_matrix, parts + SEP + matrix_text)

    # No-suffix non-steel info note
    ns_text = (
        "No material suffix — composition is '" + matc + "' "
        "(most bare parts are Steel; confirm if intentional)"
    )
    parts = parts.mask(flag_no_suffix_non_steel, parts + SEP + ns_text)

    # Strip leading separator and replace empty → "OK"
    notes = parts.str.replace(r"^ \| ", "", regex=True)
    notes = notes.where(notes != "", "OK")

    # ── Suggested Fix column ───────────────────────────────────────────────
    # Templated, actionable text per row. Picked in priority order — legacy
    # ERP first (most authoritative), then suffix, class, matrix, empty, soft.
    fix = pd.Series([""] * n, index=idx, dtype=object)

    # Legacy ERP disagrees — strongest signal, recommend the legacy value
    fix = fix.mask(
        flag_legacy_disagrees,
        "Legacy ERP says '" + legacy_mat
        + "' — verify and update Material Composition to match (or correct legacy if NetSuite is right)"
    )
    # Composition empty but legacy has a value — populate from legacy
    fix = fix.mask(
        flag_empty & has_legacy & (fix == ""),
        "Set Material Composition to legacy ERP value: '" + legacy_mat + "'"
    )

    # Forward suffix: composition is real, name is wrong → rename to add suffix
    fix = fix.mask(
        forward_mismatch & (fix == ""),
        "Either rename the part to match '" + matc + "' "
        "OR change Material Composition to '" + expected_mat + "'"
    )
    # Reverse SS / Brass / Aluminum: missing suffix, class doesn't confirm
    fix = fix.mask(
        rev_ss_mismatch & (fix == ""),
        "Verify composition: if part is plain Steel, change Material Composition. "
        "If genuinely Stainless, add -SS to the name and update class."
    )
    fix = fix.mask(
        rev_brass_mismatch & (fix == ""),
        "Verify composition: if part is plain Steel, change Material Composition. "
        "If genuinely Brass, add -B to the name and update class."
    )
    fix = fix.mask(
        rev_alum_mismatch & (fix == ""),
        "Verify composition: if part is plain Steel, change Material Composition. "
        "If genuinely Aluminum, add -D to the name and update class."
    )
    # Class mismatch: composition is right, class is mis-categorized
    fix = fix.mask(
        flag_class & (fix == ""),
        "Reclassify the item under a class containing '" + matc + "'"
    )
    # Matrix Material field disagreement
    fix = fix.mask(
        flag_matrix & (fix == ""),
        "Update the matrix Material field to match the Material Composition"
    )
    # Missing composition
    fix = fix.mask(
        flag_empty & (fix == ""),
        "Populate the Material Composition field"
    )
    # Soft no-suffix non-Steel (review)
    fix = fix.mask(
        flag_no_suffix_non_steel & (fix == ""),
        "Confirm the composition is correct; if so, no action needed"
    )
    fix = fix.where(fix != "", "—")

    report(0.95, "Finalizing…")

    # ── Assemble result ────────────────────────────────────────────────────
    result = df.copy()
    result["flag_legacy_disagrees"]           = flag_legacy_disagrees.astype(bool)
    result["flag_suffix_mismatch"]            = flag_suffix.astype(bool)
    result["flag_class_material_mismatch"]    = flag_class.astype(bool)
    result["flag_material_field_mismatch"]    = flag_matrix.astype(bool)
    result["flag_empty_material_composition"] = flag_empty.astype(bool)
    result["flag_no_suffix_non_steel"]        = flag_no_suffix_non_steel.astype(bool)
    result["flag_is_matrix_parent"]           = is_matrix_parent.astype(bool)
    result["flag_is_bb_part"]                 = is_bb_flag.astype(bool)
    result["legacy_material"]                 = legacy_mat
    result["recommended_material"]            = recommended_mat
    result["expected_material_from_name"]     = expected_mat
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
    "flag_legacy_disagrees":           "Legacy ERP Mismatch",
    "flag_suffix_mismatch":            "Suffix Mismatch",
    "flag_no_suffix_non_steel":        "No-Suffix Non-Steel (Review)",
    "flag_class_material_mismatch":    "Class Mismatch",
    "flag_material_field_mismatch":    "Matrix Field Mismatch",
    "flag_empty_material_composition": "Missing Composition",
    "flag_is_matrix_parent":           "Matrix Parent",
    "flag_is_bb_part":                 "BB Part (Exempt)",
    "any_flag":                        "Has Issue",
    "legacy_material":                 "Legacy ERP Material",
    "recommended_material":            "Recommended Material",
    "expected_material_from_name":     "Suffix-Detected Material",
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
    derived_cols   = ["legacy_material", "recommended_material",
                      "expected_material_from_name", "suggested_fix",
                      "analysis_notes"]
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
                 "Clean Records", "Issue Rate", "Legacy ERP Matches"}
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
        elif c in ("recommended_material", "expected_material_from_name", "legacy_material"):
            ws.set_column(col_idx, col_idx, 26)
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
        choices=["suffix", "class", "matrix", "empty", "legacy", "all"],
        default=["all"],
        help=(
            "Which checks to run (default: all). "
            "'suffix' = name ↔ material; 'class' = class reflects material; "
            "'matrix' = matrix Material field; 'empty' = missing composition; "
            "'legacy' = legacy ERP cross-check (requires --legacy-csv). "
            "Example: --checks suffix legacy"
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
    return p


# Maps user-friendly check names → the underlying flag column.
CHECK_NAME_TO_FLAG: dict[str, str] = {
    "suffix": "flag_suffix_mismatch",
    "class":  "flag_class_material_mismatch",
    "matrix": "flag_material_field_mismatch",
    "empty":  "flag_empty_material_composition",
    "legacy": "flag_legacy_disagrees",
}


def resolve_checks(names: list[str]) -> set[str]:
    """Convert a list like ['suffix','class'] or ['all'] → set of flag column keys."""
    if not names or "all" in names:
        return set(ISSUE_FLAGS.keys())
    return {CHECK_NAME_TO_FLAG[n] for n in names if n in CHECK_NAME_TO_FLAG}


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

    try:
        df_result = analyze_dataframe(
            df_raw,
            enabled_checks=resolve_checks(args.checks),
            legacy_lookup=legacy_lookup,
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
    root.geometry("680x760")
    root.minsize(620, 700)

    state = {"csv_path": None, "xlsx_path": None, "legacy_path": None}

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
            file_lbl.config(text=Path(path).name, foreground="#000000")
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
            legacy_lbl.config(text=Path(path).name, foreground="#000000")

    legacy_clear_btn = ttk.Button(legacy_frame, text="Clear", command=clear_legacy)
    legacy_clear_btn.pack(side="right", padx=(6, 0))
    legacy_pick_btn = ttk.Button(
        legacy_frame, text="Choose Legacy CSV…", command=pick_legacy
    )
    legacy_pick_btn.pack(side="right", padx=(10, 0))

    # ── Checks scope ───────────────────────────────────────────────────────
    checks_frame = ttk.LabelFrame(main_frame, text="Checks to Run", padding=10)
    checks_frame.pack(fill="x", pady=(12, 0))

    check_vars: dict[str, tk.BooleanVar] = {
        "suffix": tk.BooleanVar(value=True),
        "class":  tk.BooleanVar(value=True),
        "matrix": tk.BooleanVar(value=True),
        "empty":  tk.BooleanVar(value=True),
    }
    check_labels = {
        "suffix": "Name suffix ↔ material composition",
        "class":  "Class reflects material composition",
        "matrix": "Matrix Material field matches composition",
        "empty":  "Missing material composition",
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
               command=lambda: set_scope("suffix", "class", "matrix", "empty"),
               width=10).pack(side="left")
    ttk.Button(scope_btns, text="Suffix only",
               command=lambda: set_scope("suffix"),
               width=12).pack(side="left", padx=(6, 0))
    ttk.Button(scope_btns, text="Class only",
               command=lambda: set_scope("class"),
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
        height=8,
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

    open_file_btn = ttk.Button(
        btn_frame,
        text="Open Excel Report",
        command=lambda: open_path(state["xlsx_path"]) if state["xlsx_path"] else None,
        state="disabled",
    )
    open_folder_btn = ttk.Button(
        btn_frame,
        text="Open Folder",
        command=lambda: open_path(str(Path(state["xlsx_path"]).parent)) if state["xlsx_path"] else None,
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

        # Snapshot legacy path at click time too
        legacy_path = state.get("legacy_path")

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
                        0.06,
                        f"Legacy ERP: {len(legacy_lookup):,} entries loaded"
                    )

                df_result = analyze_dataframe(
                    df_raw,
                    enabled_checks=enabled,
                    progress_callback=make_phase_cb(0.06, 0.35),
                    legacy_lookup=legacy_lookup,
                )

                if legacy_lookup is not None:
                    legacy_match_count = int((df_result["legacy_material"] != "").sum())

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
