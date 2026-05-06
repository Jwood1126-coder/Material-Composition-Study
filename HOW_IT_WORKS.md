# How the Material Composition Analyzer Works

A plain-English walkthrough of what the program does, what it flags, and why.

---

## 1. The big picture

You give the analyzer a CSV exported from a NetSuite saved search. For every
item in that export, it looks at four pieces of information:

- **Name** (the part number, e.g. `0306-08-SS`)
- **External ID** (used for matching the legacy ERP file)
- **Class** (the NetSuite categorization, e.g. "Hydraulic Fittings - Stainless")
- **Material Composition** (the field this program is auditing)
- **Material** (the matrix-item Material field, only when populated)

It can also optionally take a second CSV from the **legacy ERP** with two
columns: a part number and the material that part was recorded as in the
legacy system.

For each row, the program decides whether the **Material Composition** value
in NetSuite is suspect, and if so, what it should probably be changed to.
The output is an Excel workbook with a **Summary** tab and a **Detail** tab.

> **The focus of this version is correcting Material Composition only.**
> The Class field is used as evidence — a class containing "Stainless"
> tells us the part is probably stainless steel — but the program does
> NOT recommend reclassifying items. If the class names a different
> material than the composition, the assumption is that the composition
> needs updating to match the class, not the other way around.

---

## 2. Input requirements

### Required NetSuite CSV columns

| Column | Used for |
|---|---|
| Name (or "Item Name") | Suffix detection, part-number lookup against legacy |
| Class (or "Item Class") | Classification evidence |
| Material Composition | The field being audited |

### Optional NetSuite columns

| Column | Used for |
|---|---|
| Internal ID | Detecting matrix parents (same ID on multiple rows) |
| External ID | Primary key for legacy ERP lookup |
| Material | Matrix-item field consistency check |

### Optional legacy ERP CSV columns

| Column | Aliases accepted | Used for |
|---|---|---|
| Part Number | "Part #", "PartNum", "Item Number" | Cross-check identifier |
| Material | "Material Composition" | Authoritative comparison value |

Column names are matched case-insensitively and whitespace is trimmed.
Common placeholder text — `TBD`, `N/A`, `Unknown`, `?`, `—` — is treated as
blank in both files.

---

## 3. The five rules

The analyzer fires a flag (and a row gets marked **Has Issue = YES**) only
when there is **positive evidence** that the Material Composition is wrong.
The program will not flag a row just because a suffix or class is "missing"
— absence of evidence is not evidence of a problem, and many parts
legitimately use naming conventions that don't include a material suffix.

The five rules, in priority order:

### Rule 1: Legacy ERP disagrees
**Fires when:** the legacy ERP CSV has an entry for this part number
(matched against External ID first, then Name) AND the legacy material
differs from the current NetSuite Material Composition.

**Example:** Legacy says `Brass`, NetSuite has `Steel`. Flagged.

This is the highest-confidence signal because the legacy ERP is treated
as an external authoritative source.

### Rule 2: Name suffix actively contradicts the composition
**Fires when:** the part name has a recognized material suffix AND the
material that suffix implies disagrees with the current composition AND
the class doesn't independently confirm the composition.

**Recognized suffixes:**
- `-SS` (or `SS` anywhere as its own segment) → Stainless Steel
- `-B` exactly as the last segment (NOT `-BB`, which is the Brennan Black
  program) → Brass
- `-D` anywhere as its own segment (handles `-D-GREEN`, `-D-NS`, etc.) →
  Aluminum

**Example:** part name `0306-08-SS` with composition `Steel` and a generic
class. Flagged because `-SS` says Stainless Steel.

**Example NOT flagged:** part name `2106-12-12-B` with composition
`Stainless Steel` and class `Tube Fittings - Stainless`. The class
confirms Stainless Steel, so the `-B` is treated as a legitimate naming
quirk and the row is OK.

### Rule 3: Class names a different specific material
**Fires when:** the Class string explicitly contains a material keyword
(`stainless`, `steel`, `brass`, `aluminum`, `aluminium`, `aerospace`)
that disagrees with the current composition.

**Example:** class `Hydraulic Fittings - Steel - LV` with composition
`Brass`. Flagged.

**Example NOT flagged:** class `Hydraulic Fittings - 4-Bolt Flange -
Split Flange` with composition `Brass`. The class is generic — no
material is named — so there's no evidence the composition is wrong.

> "Aerospace" in a class is treated as Aluminum-confirming, since
> aerospace fittings are predominantly aluminum.

### Rule 4: Matrix Material field disagrees
**Fires when:** the matrix-item `Material` field is populated AND it
doesn't match the row's Material Composition (or, for matrix parents
with multiple compositions, isn't found among them).

**Example:** matrix `Material` = "Steel", `Material Composition` =
"Brass, Stainless Steel". Flagged.

### Rule 5: Material Composition is empty
**Fires when:** the Material Composition field is blank (after sentinel
cleanup — so `TBD`, `N/A`, etc. count as blank too).

**Example:** Material Composition is blank. Flagged. If a legacy ERP
entry exists for this part, the suggested fix tells you to use that
value.

---

## 4. What does NOT get flagged

These cases are intentionally NOT flagged, because flagging them
created noise without an actionable recommendation:

- **Brass composition with no `-B` in the name** and no contradicting
  class. Many parts legitimately use alternate naming conventions
  without the material suffix. Without independent evidence, the program
  trusts the composition.
- **Stainless Steel composition with no `-SS`** and no contradicting
  class.
- **Aluminum composition with no `-D`** and no contradicting class.
- **Generic class** (e.g. `Hydraulic Fittings - 4-Bolt Flange`) with
  any composition. The class doesn't specifically name a material, so
  it can't disagree.
- **Brennan Black parts** (`-BB` segment in the name OR class is
  "Brennan Black"). These are exempt from the class check by design —
  the class is a program name, not a material category.

---

## 5. The Recommended Material column

Every row gets a **Recommended Material** value. The priority is:

1. **Legacy ERP material** (if matched) — highest authority
2. **Suffix-derived material** — when forward suffix mismatch fires
3. **Class-named material** — when class mismatch fires
4. **Current Material Composition** — the default; trust what's there
5. **Class-named material** — when composition is empty AND class
   names a specific material
6. **"Steel"** — last-resort default when composition is empty AND
   nothing else helps

If the row is **Has Issue = NO**, the Recommended Material always equals
the current composition. The program never silently suggests a change
on rows it considers correct.

---

## 6. The Suggested Fix column

A short, templated sentence telling you what to do. The fix matches the
flag that fired:

| Flag | Suggested Fix |
|---|---|
| Legacy ERP disagrees | "Legacy ERP says 'X' — verify and update Material Composition to match (or correct legacy if NetSuite is right)" |
| Empty composition + legacy match | "Set Material Composition to legacy ERP value: 'X'" |
| Forward suffix mismatch | "Either rename the part to match 'Y' OR change Material Composition to 'X'" |
| Class names different material | "Class indicates 'X' — verify and update Material Composition to 'X'" |
| Matrix Material field disagrees | "Update the matrix Material field to match the Material Composition" |
| Empty composition (no legacy) | "Populate the Material Composition field" |
| No issue | "—" |

---

## 7. Output workbook

### Summary tab
- Total records analyzed
- Records with issues / clean / issue rate
- Legacy ERP match coverage (if a legacy file was loaded)
- Issue Breakdown — count per flag type. **A single row may be counted
  in multiple lines** (e.g., a row can have both a suffix mismatch and
  a class mismatch).

### Detail tab
- All source columns from the input CSV
- **Legacy ERP Material** — what the legacy file said for this part
- **Recommended Material** — what to change composition to (or keep)
- **Suffix-Detected Material** — what the suffix implies, if anything
- **Suggested Fix** — actionable guidance
- **Analysis Notes** — explanation of any flags
- One column per flag (YES / —)
- **Has Issue** — overall flag status

The Detail tab has Excel autofilter enabled. To see only flagged rows,
click the dropdown on **Has Issue** and pick "YES". To see only one
issue type, filter the corresponding flag column instead.

---

## 8. Sanity guards built in

- **Encoding fallback** — utf-8-sig (handles Excel BOM) → latin-1 if
  utf-8 decode fails
- **Whitespace-tolerant** column matching and value comparison
- **Sentinel cleanup** — `TBD`, `N/A`, `Unknown`, `?`, `—`, `-` →
  treated as blank in both NetSuite and legacy CSVs
- **Case-insensitive** material-name comparisons
- **`-BB` is not `-B`** — exact segment matching, no partial collisions
- **"Steel" class doesn't satisfy "Stainless Steel"** — separate
  keyword logic for the two
- **Multi-composition rows** (matrix parents with comma/semicolon-
  separated values) — each composition checked independently
- **Duplicate Internal IDs** — flagged as matrix parents (informational
  only)
- **First-occurrence wins** in the legacy lookup when a part number
  appears multiple times — surfaced as a count of conflicts but not
  silently overwritten
- **NaN/Inf** in any column — converted to blank before writing Excel
  (xlsxwriter rejects them by default)

---

## 9. Worked examples

### Example A — Stainless part with -B in the name
**Input row:** `2106-12-12-B`, class `Tube Fittings - Stainless - LV`,
Material Composition `Stainless Steel`.

**Logic:**
- Suffix `-B` would imply Brass, but…
- Class contains "Stainless", which confirms the composition independently
- Forward suffix mismatch is suppressed by class confirmation

**Result:** Has Issue = NO. Recommended Material = Stainless Steel.

### Example B — Brass part with no suffix and a generic class
**Input row:** `1961-SHAFT`, class is blank, Material Composition `Brass`.

**Logic:**
- No suffix detected
- Class is blank — no signal either way
- No legacy entry
- All rules check for *positive evidence* and find none

**Result:** Has Issue = NO. Recommended Material = Brass.

### Example C — Class actively says one material, composition says another
**Input row:** `0306-09-LP`, class `Hydraulic Fittings - Steel`,
Material Composition `Brass`.

**Logic:**
- Class explicitly names "Steel"
- Class-named material (Steel) ≠ composition (Brass)
- Rule 3 fires

**Result:** Has Issue = YES. Recommended Material = Steel.
**Suggested Fix:** "Class indicates 'Steel' — verify and update
Material Composition to 'Steel'".

### Example D — Legacy ERP disagrees
**Input row:** `RANDO-PT-1`, NetSuite Material Composition `Steel`.
Legacy ERP CSV has this part listed as `Brass`.

**Logic:**
- Legacy lookup matches via External ID
- Legacy material (Brass) ≠ NetSuite composition (Steel)
- Rule 1 fires

**Result:** Has Issue = YES. Recommended Material = Brass.
**Suggested Fix:** "Legacy ERP says 'Brass' — verify and update
Material Composition to match (or correct legacy if NetSuite is right)".

### Example E — Empty composition, legacy fills the gap
**Input row:** `EMPTY-LEG`, Material Composition blank, legacy ERP
has it listed as `Aluminum`.

**Logic:**
- Composition is blank — Rule 5 fires
- Legacy match found

**Result:** Has Issue = YES. Recommended Material = Aluminum.
**Suggested Fix:** "Set Material Composition to legacy ERP value:
'Aluminum'".

---

## 10. Workflow tips

1. **Run with all checks enabled first** to get the overall picture.
2. **Use the autofilter** on the Detail tab to focus on one flag type
   at a time. Suffix issues are usually the easiest to fix in bulk.
3. **Treat the Suggested Fix column as a hypothesis**, not a command.
   The program is right most of the time, but always sanity-check
   against the actual part before making changes in NetSuite.
4. **Re-run after each round of corrections** to confirm the issue
   count is going down and to surface any new ones revealed by the
   fixes.
5. **Use the legacy ERP cross-check** when you have it. It's the
   highest-confidence signal and dramatically reduces false positives
   on your raw-material-style part numbers.
