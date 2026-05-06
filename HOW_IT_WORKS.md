# How the Material Composition Analyzer Works

A plain-English walkthrough of what the program does, what it flags, and why.

---

## 1. The big picture

You give the analyzer a CSV exported from a NetSuite saved search. For every
item in that export, it looks at:

- **Name** (the part number, e.g. `0306-08-SS`)
- **External ID** (used for matching the legacy ERP file)
- **Class** (the NetSuite categorization, e.g. "Hydraulic Fittings - Stainless")
- **Material Composition** (the field this program is auditing)

It can also optionally take a second CSV from the **legacy ERP** with two
columns: a part number and the material that part was recorded as in the
legacy system.

For each row, the program decides whether the **Material Composition** value
in NetSuite is suspect, and if so, what it should probably be changed to.

> **The focus of this version is correcting Material Composition.**
> Class is treated as supporting evidence — never as a fix target. If the
> class names a material that disagrees with the composition, the
> recommendation is to update the composition (not reclassify the item).

---

## 2. The evidence ranking

The program uses three sources of evidence to decide whether a Material
Composition value is right or wrong. They are ranked by reliability — when
two sources disagree, the higher-priority one wins:

| Tier | Source | Notes |
|---|---|---|
| **1 (highest)** | **Name** — suffix or pattern in the part name | Most authoritative. -SS, -B, -D, "ALUM", "BR", "316", "S/S", standalone S, etc. |
| **2 (middle)** | **Legacy ERP material** | An external authoritative source. Can occasionally be wrong, so name still wins. |
| **3 (lowest)** | **Class** | Often wrong. Used as a hint only — never overrides name or legacy. |

This ranking has two consequences:

1. **A higher-tier confirmation suppresses lower-tier disagreement flags.**
   If the name's suffix says "Brass" and the composition is also Brass, the
   program won't flag the row even if the class says Steel and the legacy
   says Aluminum — the name is the most reliable signal and it confirms
   the composition.
2. **The Recommended Material always reflects the highest-tier signal in
   play.** If the name says X, recommendation is X. If only legacy disagrees,
   recommendation is the legacy value. If only the class names a different
   material, recommendation is the class-named material.

---

## 3. Input requirements

### Required NetSuite CSV columns
| Column | Used for |
|---|---|
| Name (or "Item Name") | Suffix detection, part-number lookup against legacy |
| Class (or "Item Class") | Classification evidence (Tier 3) |
| Material Composition | The field being audited |

### Optional NetSuite columns
| Column | Used for |
|---|---|
| Internal ID | Detecting matrix parents (same ID on multiple rows) |
| External ID | Primary key for legacy ERP lookup |

### Optional legacy ERP CSV columns
| Column | Aliases accepted | Used for |
|---|---|---|
| Part Number | "Part #", "PartNum", "Item Number" | Cross-check identifier |
| Material | "Material Composition" | Authoritative comparison value |

Column names are matched case-insensitively and whitespace is trimmed.
Common placeholder text — `TBD`, `N/A`, `Unknown`, `?`, `—` — is treated as
blank in both files.

---

## 4. The four rules

The analyzer fires a flag (and a row gets marked **Has Issue = YES**) only
when there is **positive evidence** that the Material Composition is wrong.
Absence of evidence — e.g., a Brass part that doesn't have `-B` in its
name — is not flagged. Many parts use legitimate alternate naming
conventions.

### Rule 1 — Name implies a different material (Tier 1)
**Fires when:** the part name has a recognized suffix or pattern AND the
material that name implies disagrees with the current composition.

**Recognized signals (in priority order):**

| Signal | Means | Examples |
|---|---|---|
| `SS` segment | Stainless Steel | `0306-08-SS`, `2106-04-SS-LW` |
| Last segment exactly `B` | Brass | `0306-08-B` (NOT `-BB`, which is Brennan Black) |
| `D` segment | Aluminum | `0306-08-D`, `0318-08-D-GREEN` |
| `\bbrass\b`, `\bbrs\b`, `\bbr\b`, name ending in `BR`, digit+B at end | Brass | `200-24BR`, `202-30B` |
| `\bSS\b`, `\bS/S\b`, `316`, `316L`, `304`, `304L`, `\bstainless\b`, standalone `S` | Stainless Steel | `2-S`, `2/1T MSA 316`, `200-05 316 S/S` |
| `\balum` | Aluminum | `ALUM VENT 3 FPT` |
| `\bsteel\b` | Steel | `STEEL TUBE 06` |

**Example flagged:** `2106-12-12-B` with composition `Stainless Steel`. The
`-B` suffix says Brass; that disagrees with the composition. Even though
the class might say "Stainless", **the name overrides the class** —
flagged. Recommendation: Brass.

**Example NOT flagged:** `200-24BR` with composition `Brass`. The "BR"
pattern matches — the name confirms the composition. No flag.

### Rule 2 — Legacy ERP disagrees (Tier 2)
**Fires when:** the legacy ERP has an entry for this part (matched by
External ID first, then Name) AND the legacy material differs from the
current composition AND the **name doesn't already confirm the composition**.

**Why suppressed by name:** if the suffix/pattern says the composition is
right, we trust that more than the legacy entry — legacy can occasionally
be wrong. The legacy disagreement is treated as a stale legacy record,
not a NetSuite problem.

**Example flagged:** legacy says `Brass`, NetSuite has `Steel`, name has
no relevant suffix. Flagged. Recommendation: Brass (from legacy).

**Example NOT flagged:** name `0306-08-SS`, NetSuite composition
`Stainless Steel`, legacy says `Steel`. The name's `-SS` confirms the
NetSuite composition — the legacy entry is the odd one out. No flag.

### Rule 3 — Class names a different specific material (Tier 3)
**Fires when:** the Class string contains a recognized material keyword
(`stainless`, `steel`, `brass`, `aluminum`, `aluminium`, `aerospace`)
naming a material that disagrees with the composition AND **neither the
name nor the legacy already confirms the composition**.

**Why this is the lowest priority:** class is the least reliable source.
If the name or legacy already confirms the composition, a class
disagreement is treated as a misclassification of the item — not a
composition problem.

**Example flagged:** class `Hydraulic Fittings - Steel - LV`, composition
`Brass`, no name signal, no legacy entry. Flagged. Recommendation: Steel
(from class).

**Example NOT flagged:** class `Hydraulic Fittings - 4-Bolt Flange -
Split Flange`, composition `Brass`. The class is generic — it doesn't
name a specific material — so it offers no evidence the composition is
wrong. No flag.

> "Aerospace" in a class is treated as Aluminum-confirming.

### Rule 4 — Material Composition is empty
**Fires when:** the Material Composition field is blank (after sentinel
cleanup — so `TBD`, `N/A`, etc. count as blank).

If a legacy ERP entry exists for this part, the suggested fix tells you
to use that value. Otherwise, the recommendation defaults to whatever the
class names, or "Steel" as a last resort.

---

## 5. What does NOT get flagged

Intentionally, by design:

- **Brass composition with no `-B` in the name** and a generic or empty
  class. Many parts legitimately use alternate naming conventions without
  the material suffix.
- **Stainless Steel composition with no `-SS`** and no contradicting class.
- **Aluminum composition with no `-D`** and no contradicting class.
- **Generic class** (e.g. `Hydraulic Fittings - 4-Bolt Flange`) with any
  composition. No material is named, so the class can't disagree.
- **Brennan Black parts** (`-BB` segment in the name OR class is "Brennan
  Black"). Exempt from the class check by design — `-BB` is a program
  designation, not a material.

---

## 6. The Recommended Material column

Every row gets a **Recommended Material** value. Priority (highest →
lowest):

1. **Name-implied material** — when Rule 1 fires
2. **Legacy ERP material** — when Rule 2 fires
3. **Class-named material** — when Rule 3 fires
4. **Current Material Composition** — the default; trust what's there
5. **Class-named or legacy material** — when composition is empty
6. **"Steel"** — last-resort default when composition is empty AND no
   other source provides a material

If the row is **Has Issue = NO**, the Recommended Material always equals
the current composition. The program never silently suggests a change on
rows it considers correct.

---

## 7. The Suggested Fix column

A short, templated sentence. Picked in priority order from whichever
flag fires:

| Flag | Suggested Fix |
|---|---|
| Name implies different material | "Name suggests material is 'X' — verify and update Material Composition to 'X'" |
| Legacy ERP disagrees | "Legacy ERP says 'X' — verify and update Material Composition to match (or correct legacy if NetSuite is right)" |
| Empty composition + legacy match | "Set Material Composition to legacy ERP value: 'X'" |
| Class names different material | "Class indicates 'X' — verify and update Material Composition to 'X'" |
| Empty composition (no legacy) | "Populate the Material Composition field" |
| No issue | "—" |

---

## 8. Output workbook

### Summary tab
- Total records, records with issues, clean count, issue rate
- Legacy ERP match coverage (when a legacy file was loaded)
- Issue Breakdown — count per flag type. **A single row may be counted in
  multiple lines** (a row can have both a name mismatch AND a class mismatch).

### Detail tab
- All source columns from the input CSV
- **Legacy ERP Material** — what the legacy file said for this part
- **Recommended Material** — what the composition should be (or keep)
- **Suffix-Detected Material** — what the name implies, if anything
- **Suggested Fix** — actionable guidance
- **Analysis Notes** — explanation of any flags
- One column per flag (YES / —)
- **Has Issue** — overall flag status

The Detail tab has Excel autofilter enabled. Filter on **Has Issue = YES**
to see only flagged rows, or filter individual flag columns for one
issue type.

---

## 9. Sanity guards built in

- **Encoding fallback** — utf-8-sig (handles Excel BOM) → latin-1
- **Whitespace-tolerant** column matching and value comparison
- **Sentinel cleanup** — `TBD`, `N/A`, `Unknown`, `?`, `—`, `-` → blank
- **Case-insensitive** material comparisons
- **`-BB` is not `-B`** — exact segment matching, no partial collisions
- **"Steel" class doesn't satisfy "Stainless Steel"** — separate keyword
  logic
- **Multi-composition rows** (matrix parents with comma/semicolon-
  separated values) — each composition checked independently
- **Duplicate Internal IDs** — flagged as matrix parents (informational
  only)
- **First-occurrence wins** in the legacy lookup when a part appears
  multiple times — surfaced as a count of conflicts but not silently
  overwritten
- **NaN/Inf** in any column → blank before writing Excel

---

## 10. Worked examples

### Example A — `-B` suffix, but composition and class both say Stainless Steel
**Input:** name `2106-12-12-B`, class `Tube Fittings - Stainless - LV`,
composition `Stainless Steel`.

**Logic:**
- The `-B` suffix says Brass — name (Tier 1) implies Brass.
- Name disagrees with composition → **Rule 1 fires**.
- Class confirms composition (Stainless), but class is Tier 3 — it
  cannot suppress a Tier 1 disagreement.

**Result:** Has Issue = YES. Recommendation: Brass. Suggested Fix: "Name
suggests material is 'Brass' — verify and update Material Composition to
'Brass'".

### Example B — Brass part with no suffix and a generic class
**Input:** name `1961-SHAFT`, class blank, composition `Brass`.

**Logic:**
- No name signal (no SS/B/D suffix, no BR/BRASS pattern)
- No class signal
- No legacy entry
- All four rules check for *positive evidence*; none find any

**Result:** Has Issue = NO. Recommendation: Brass.

### Example C — Class names a material when nothing else helps
**Input:** name `0306-09-LP`, class `Hydraulic Fittings - Steel - LV`,
composition `Brass`.

**Logic:**
- Name has no relevant suffix
- No legacy entry
- Class explicitly names "Steel" → Tier 3
- Class disagrees with composition; nothing higher-tier confirms → **Rule 3
  fires**

**Result:** Has Issue = YES. Recommendation: Steel.

### Example D — Legacy disagrees but name confirms composition
**Input:** name `FANCY-SS`, composition `Stainless Steel`, legacy says `Steel`.

**Logic:**
- Name `-SS` confirms the composition (Tier 1)
- Legacy disagrees but is Tier 2 — suppressed by Tier 1 confirmation

**Result:** Has Issue = NO. Recommendation: Stainless Steel.
(Implicitly: the legacy record is the wrong one, but we don't fix that
from this tool.)

### Example E — Three-way disagreement: legacy + class against composition
**Input:** name has no relevant suffix, composition `Stainless Steel`,
legacy says `Brass`, class names `Steel`.

**Logic:**
- No Tier 1 signal
- Legacy disagrees → **Rule 2 fires**
- Class also disagrees, but Rule 3 is suppressed by legacy (Tier 2 > Tier 3)

**Result:** Has Issue = YES. Recommendation: Brass (legacy wins over class).
Notes call out both disagreements so the human can sort it out.

---

## 11. Workflow tips

1. **Run with all checks enabled first** to get the overall picture.
2. **Use the autofilter** on the Detail tab to focus on one flag type at
   a time. Name-mismatch issues are usually the easiest to fix in bulk.
3. **Treat the Suggested Fix column as a hypothesis**, not a command. The
   program is right most of the time, but always sanity-check before
   making changes in NetSuite.
4. **Re-run after each round of corrections** to confirm the issue count
   is going down.
5. **Use the legacy ERP cross-check** when you have it. It's a strong
   second-tier signal and dramatically reduces false positives on
   raw-material-style part numbers that the name patterns can't classify.
