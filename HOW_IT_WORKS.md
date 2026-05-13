# How the Material Composition Analyzer Works

A plain-English walkthrough of what the program does, what it flags, and why.
Written so a non-technical reader can read this top to bottom and verify the
logic without reading any code.

---

## 1. What the program is for

You give it a CSV exported from a NetSuite saved search. It looks at every
item and decides whether the **Material Composition** field on that item is
probably wrong. It writes the result to an Excel workbook so a human can
review the flagged rows.

It is intentionally conservative. The goal is to do most of the work of
finding wrong compositions, not to be perfect. Whenever the program can't
confidently say what the material is, it says nothing rather than guessing.

> **Scope:** this version only tries to correct **Material Composition**.
> If the Class on an item disagrees with the composition, the program
> recommends fixing the composition — never reclassifying the item.

---

## 2. The three pieces of evidence

For each item, the program looks at up to three sources:

| Tier | Source | Reliability |
|---|---|---|
| **1 — Highest** | **Name** — recognized patterns inside the part number | Most authoritative. If a strict pattern matches, the program trusts it. |
| **2 — Middle** | **Legacy ERP material** — looked up in a second optional CSV | Usually right. Occasionally wrong. |
| **3 — Lowest** | **Class** — the NetSuite class string | Often wrong. Treated only as a weak hint. |

Two rules govern how the tiers interact:

1. **Higher tier wins ties and conflicts.** If the Name pattern says Brass
   and the Class says Stainless, the program goes with Brass. It never lets
   a Class disagreement override a Name signal.
2. **A higher-tier confirmation silences a lower-tier disagreement.** If
   the Name says the composition is right, the program won't flag the row
   just because the Legacy ERP disagrees — it assumes the legacy record is
   the one that's outdated.

---

## 3. How a Name signal is recognized (the strict rules)

The program splits the part number on dashes. Example:
`0306-08-SS-LW` → segments `0306`, `08`, `SS`, `LW`.

It then looks **only at the last 3 segments** of the name. Material
indicators that appear earlier than the last 3 segments are ignored. (A
"STEEL" word at the front of a long name is not treated as a material
signal — materials appear near the end of part numbers, not the start.)

Inside those last 3 segments, only these strict patterns count:

| Pattern | Means | Notes |
|---|---|---|
| A segment exactly equal to `SS` | **Stainless Steel** | Must be the whole segment, between dashes. `SS-LW` qualifies; `HNBR` does not. |
| The **last** segment is exactly `B` | **Brass** | Must be the whole final segment. `HNBR`, `30B`, `-BR`, `-BB` do **not** count. |
| A segment exactly equal to `D` | **Aluminum** | Whole segment only. |
| A segment exactly equal to `ZN` | **Steel** (zinc plated base) | Whole segment only. |
| The word `alum` appearing in the last 3 segments | **Aluminum** | Word match — catches `ALUM`, `ALUMINUM`, `ALUMINIUM`. |
| The word `steel` (whole word) in the last 3 segments | **Steel** | Whole-word match. |

That is the **complete** list. The program does **not** treat any of the
following as a material signal anymore: `BR`, `BRS`, `BRASS` text, `316`,
`304`, `S/S`, a lone `S`, `aerospace`. These caused false positives in
earlier versions and were removed.

If more than one pattern matches in the same name, priority is:
`SS` → `-B` → `D` → `ZN` → `alum` keyword → `steel` keyword.

If nothing matches, the Name simply has no signal — that's normal, and not
itself a flag.

---

## 4. How the Legacy ERP cross-check works

The legacy ERP cross-check is optional. If you provide a second CSV with
two columns — a **part number** column and a **material** column — the
program looks each NetSuite row up in it.

- It matches first by **External ID**, then by **Name**.
- Placeholder text like `TBD`, `N/A`, `Unknown`, `?`, `—`, `-` is treated as
  blank.
- If the same part appears more than once in the legacy file with conflicting
  materials, the **first occurrence wins** — the program won't silently pick
  a "newer" one.

The looked-up value is shown in the **Legacy ERP Material** column.

---

## 5. How the Class signal works

The program scans the Class string for these whole keywords:

- **stainless** → class names Stainless Steel
- **steel** (without "stainless") → class names Steel
- **brass** → class names Brass
- **alumin** (matches Aluminum or Aluminium) → class names Aluminum

Priority is Stainless → Steel → Brass → Aluminum (a class containing
"stainless steel" resolves to Stainless Steel, not Steel).

The word **aerospace** is shown in the **Matched Signal** column as an
informational note, but it does **not** confirm Aluminum any more —
aerospace fittings can be Aluminum, Titanium, Steel, or other alloys.

A class with **no** material keyword (e.g. `Hydraulic Fittings - 4-Bolt
Flange`) provides no Class signal at all, and the program won't fire a
Class flag on those rows.

**Brennan Black parts are exempt from the Class check.** A part is
Brennan Black if `BB` appears as a whole segment in the name, **or** if
the class contains "Brennan Black". `BB` is a program designation, not
a material, and the class on those parts won't (and shouldn't) contain
the material word.

---

## 6. The five things the program flags as issues

A row is marked **Has Issue = YES** if any of these fire. The program only
flags when it has *positive* evidence something is wrong — a Brass part
without `-B` in the name is **not** flagged just because the name doesn't
say Brass.

### Issue 1 — Name Mismatch
**Fires when:** the Name has one of the strict signals from §3, **and** that
signal disagrees with the current Material Composition.

> Example — `2106-12-12-B`, composition `Stainless Steel`, class
> `Tube Fittings - Stainless`. The `-B` last segment says Brass. The
> composition says Stainless. They disagree, so the row is flagged.
> Recommendation: Brass. (Class says Stainless, but Name beats Class.)

### Issue 2 — Legacy ERP Mismatch
**Fires when:** the legacy ERP has an entry for this part, **and** that
entry disagrees with the current Material Composition, **and** the Name
signal hasn't already confirmed the composition is right.

> Example — name `1234-FOO` (no recognized signal), composition `Steel`,
> legacy says `Brass`. Flagged. Recommendation: Brass.

> Example NOT flagged — name `FANCY-SS`, composition `Stainless Steel`,
> legacy says `Steel`. The Name's `-SS` confirms the composition, so the
> legacy disagreement is treated as a stale legacy record and silenced.

### Issue 3 — Name ↔ Legacy Conflict
**Fires when:** the Name has a recognized signal **and** the legacy ERP
has an entry, **and** those two disagree.

This fires even when one of them matches the current composition — because
one of the two trusted sources is wrong, and a human should look.

> Example — name `9999-SS`, legacy says `Brass`, composition `Stainless
> Steel`. The Name says Stainless, the legacy says Brass. Flagged for
> review. (The Name still wins for the recommendation: Stainless Steel.)

### Issue 4 — Class Mismatch
**Fires when:** the Class string names a specific material (Stainless,
Steel, Brass, or Aluminum) **and** that material differs from the current
composition, **and** neither the Name nor the Legacy ERP has already
confirmed the composition is right.

This is the **weakest** flag. Class is the least reliable source, so it's
checked last and silenced first.

> Example — name has no relevant pattern, no legacy entry, class is
> `Hydraulic Fittings - Steel - LV`, composition is `Brass`. Flagged.
> Recommendation: Steel.

> Example NOT flagged — same situation but the legacy ERP says `Brass`.
> Legacy confirms the composition, so the Class disagreement is silenced.

### Issue 5 — Missing Composition
**Fires when:** the Material Composition field is blank (after sentinel
cleanup — `TBD`, `N/A`, etc. count as blank).

---

## 7. Three things the program notes but does NOT flag as issues

These appear in their own columns for context, but they do not count as
issues and don't turn `Has Issue` to YES on their own:

- **Matrix Parent** — the same Internal ID appears on more than one row
  (the parent of a matrix item).
- **BB Part (Exempt)** — `BB` segment or "Brennan Black" in the class.
  Exempt from the Class check.
- **Unknown Composition** — the composition is filled in but isn't one
  of the recognized materials (Stainless Steel, Steel, Brass, Aluminum /
  Aluminium / Aluminum Alloy). The program could only do limited
  validation on these.

---

## 8. The Recommended Material column

Every row gets a value. It is picked like this:

1. **If no flag fired**, the recommendation equals the current composition.
   The program never silently suggests changing a row it considers correct.
2. **If the composition is blank**, an estimate is built from whatever
   evidence exists, in this priority: Name signal → Legacy ERP → Class.
   **If none of those have a signal, the recommendation is left BLANK.**
   The program is willing to say "I don't know" rather than guess "Steel".
3. **If the composition is filled in but a flag fired**, the recommendation
   is the higher-tier signal that disagreed: Name beats Legacy beats Class.

---

## 9. The Confidence column

When the program recommends a change, it scores how strong the evidence is:

| Confidence | Meaning |
|---|---|
| **High** | A literal suffix (`-SS`, `-B`, `-D`, `-ZN`) supports the recommendation, **OR** two or more independent sources (Name, Legacy, Class) agree on it. |
| **Medium** | Either the Legacy ERP is the only supporting source, **OR** a keyword Name signal (`alum`/`steel`) stands alone, **OR** the High case was downgraded because the Name and Legacy disagreed. |
| **Low** | The Class is the only source supporting the recommendation. |
| **blank** | Row has no issue, or the row has an issue but no evidence was strong enough to make a recommendation. |

A practical workflow: **filter Confidence = High first.** Those are the
recommendations the program is most sure about, and the bulk corrections
you do there cost the least review time. Then work down to Medium and Low.

---

## 10. The Matched Signal column

For every row, this column lists which signals fired and what they said.
It's the transparency column — if you don't believe a flag, look here
first. Format:

`name: <name signal> | legacy: <legacy material> | class: <class material>`

Any of the three parts is omitted when that source had nothing to say. A
row with no signals shows `—`.

> Example: `name: -SS segment | legacy: Steel | class: Stainless Steel`
> tells you the Name pattern was the literal `-SS` segment, the legacy ERP
> said Steel, and the Class string contained "stainless". You can now
> judge the flag yourself.

---

## 11. The Suggested Fix column

A short, templated sentence per row, picked in priority order:

| When this fires | Suggested Fix says |
|---|---|
| Legacy ERP disagrees | "Legacy ERP says 'X' — verify and update Material Composition to match (or correct legacy if NetSuite is right)" |
| Name ↔ Legacy Conflict | "Name says 'X' but Legacy ERP says 'Y' — investigate which source is correct before changing composition" |
| Missing composition, no signal at all | "Material Composition is blank and no name/legacy/class signal — manual review needed" |
| Missing composition, only Class names a material | "Material Composition is blank; class indicates 'X' — verify and set to this value" |
| Missing composition, Legacy has a value | "Material Composition is blank; legacy ERP says 'X' — set to this value" |
| Missing composition, Name signal present | "Material Composition is blank; name suggests 'X' — verify and set to this value" |
| Name signal disagrees with composition | "Name suggests material is 'X' — verify and update Material Composition to 'X'" |
| Class names different material | "Class indicates 'X' — verify and update Material Composition to 'X'" |
| Missing composition (fallback) | "Populate the Material Composition field" |
| No issue | "—" |

---

## 12. Output workbook

### Summary tab
- Total rows, rows with issues, clean count, issue rate
- Legacy ERP match coverage (when a legacy file was used)
- Issue breakdown by flag. A single row can be counted in multiple lines
  if it triggered more than one flag.

### Detail tab
- All source columns from the input CSV
- **Legacy ERP Material** — what the legacy file said
- **Recommended Material** — what the composition should be
- **Suffix-Detected Material** — what the Name signal implied, if anything
- **Matched Signal** — which signals fired (see §10)
- **Confidence** — High / Medium / Low (see §9)
- **Suggested Fix** — actionable guidance (see §11)
- **Analysis Notes** — human-readable explanation of any flags
- One column per flag (YES / —)
- **Has Issue** — overall flag status

Flagged rows are highlighted red. YES cells in flag columns are bold red.
Excel autofilter is enabled on every data sheet.

There is also a **Flagged Items** sheet (only the rows with `Has Issue = YES`),
and one sheet per flag type that fired (skipped if a single type would
exceed 20,000 rows).

---

## 13. Safety guards built in

- Encoding fallback: utf-8-sig (handles Excel BOM) → latin-1
- Whitespace-tolerant column matching and value comparison
- Sentinel cleanup: `TBD`, `N/A`, `Unknown`, `?`, `—`, `-` → treated as blank
- All material comparisons are case-insensitive
- `-BB` is **not** `-B` — segment matching is exact
- `-B` must be the **last** segment — `-B-LW` does not qualify
- A class containing "steel" does not satisfy a "stainless steel" composition
- Multi-value compositions (comma/semicolon/pipe/newline separated) are each
  checked independently
- Duplicate Internal IDs are flagged as matrix parents (informational only)
- First-occurrence wins in the legacy lookup on duplicate keys
- NaN / ±Inf values in any column are blanked before writing Excel

---

## 14. Worked examples

### Example A — `-B` suffix, composition and class both say Stainless
**Input:** name `2106-12-12-B`, class `Tube Fittings - Stainless - LV`,
composition `Stainless Steel`, no legacy entry.

- Name signal: `-B` last segment → Brass.
- Legacy: none.
- Class signal: Stainless Steel.
- Name (Tier 1) disagrees with composition → **Name Mismatch fires.**
- Class confirms composition but is Tier 3 — cannot silence Tier 1.

**Result:** Has Issue = YES. Recommendation: **Brass**. Confidence: **High**
(literal `-B` suffix supports it). Suggested Fix: "Name suggests material is
'Brass' — verify and update Material Composition to 'Brass'".

### Example B — Brass part with no `-B` and a generic class
**Input:** name `1961-SHAFT`, class blank, composition `Brass`.

- No Name signal (no SS / B / D / ZN segments; no `alum` or `steel` word).
- No legacy entry.
- No Class signal.

**Result:** Has Issue = NO. Recommendation: Brass. (The program does not
flag the *absence* of a Name signal.)

### Example C — Class is the only signal
**Input:** name `0306-09-LP`, class `Hydraulic Fittings - Steel - LV`,
composition `Brass`, no legacy entry.

- No Name signal.
- No legacy entry.
- Class names Steel, disagrees with composition → **Class Mismatch fires.**

**Result:** Has Issue = YES. Recommendation: **Steel**. Confidence: **Low**
(Class is the only supporting source).

### Example D — Legacy disagrees but Name confirms composition
**Input:** name `FANCY-SS`, composition `Stainless Steel`, legacy says `Steel`.

- Name signal: `-SS` → Stainless Steel. Confirms composition.
- Legacy disagrees, but legacy is Tier 2 — silenced by Tier 1 confirmation.

**Result:** Has Issue = NO. Recommendation: Stainless Steel.

### Example E — Three-way disagreement, Name signal absent
**Input:** name has no recognized pattern, composition `Stainless Steel`,
legacy says `Brass`, class names `Steel`.

- No Name signal.
- Legacy disagrees with composition → **Legacy ERP Mismatch fires.**
- Class also disagrees, but Class is silenced because Legacy already fired
  (we have a higher-tier explanation; we don't need a duplicate).

**Result:** Has Issue = YES. Recommendation: **Brass** (Legacy beats Class).
Confidence: **Medium** (Legacy alone).

### Example F — Name and Legacy disagree
**Input:** name `9999-SS`, composition `Stainless Steel`, legacy says `Brass`.

- Name signal: `-SS` → Stainless Steel. Confirms composition.
- Legacy says Brass. Not silenced by Name → wait, actually it is —
  Legacy Mismatch is silenced because Name confirms composition.
- But the Name and Legacy disagree with **each other** →
  **Name ↔ Legacy Conflict fires.**

**Result:** Has Issue = YES. Recommendation: Stainless Steel (Name wins).
Confidence: **Medium** (would be High for the literal suffix, but the
Name↔Legacy conflict downgrades it). Suggested Fix tells the human to
investigate which of the two trusted sources is wrong.

### Example G — `-ZN` zinc-plated part
**Input:** name `0306-08-ZN`, composition `Stainless Steel`, no legacy.

- Name signal: `-ZN` → Steel (zinc-plated parts have a steel base).
- Name disagrees with composition → **Name Mismatch fires.**

**Result:** Has Issue = YES. Recommendation: **Steel**. Confidence: **High**
(literal suffix).

### Example H — Blank composition with no evidence anywhere
**Input:** name `1961-SHAFT`, class blank, composition blank, no legacy.

- No signal of any kind.
- **Missing Composition fires** (because composition is blank).
- No source can supply an estimate.

**Result:** Has Issue = YES. Recommendation: **(blank)**. Confidence:
**(blank)**. Suggested Fix: "Material Composition is blank and no
name/legacy/class signal — manual review needed".

### Example I — Blank composition, Name signal supplies the estimate
**Input:** name `0306-08-SS`, composition blank, no legacy.

- Name signal: `-SS` → Stainless Steel.
- **Missing Composition fires.**
- Estimate built from Name signal.

**Result:** Has Issue = YES. Recommendation: **Stainless Steel**.
Confidence: **High** (literal suffix). Suggested Fix: "Material Composition
is blank; name suggests 'Stainless Steel' — verify and set to this value".

---

## 15. Workflow tips

1. **Run with all checks enabled first** to get the overall picture.
2. **Sort or filter by Confidence = High** — those are the safest bulk
   corrections.
3. **When in doubt, read the Matched Signal column.** It tells you exactly
   why the program made its recommendation, in plain text.
4. **Treat every Suggested Fix as a hypothesis** rather than a command.
   Sanity-check before changing anything in NetSuite.
5. **Re-run after each round of corrections** to confirm the issue count
   is going down.
6. **Use the legacy ERP cross-check whenever you have it.** It dramatically
   improves coverage on parts whose names don't follow the strict signal
   conventions.
7. **Don't expect zero false positives.** The program is calibrated to be
   helpful to a human reviewer, not to be a robot that can finalize
   changes unattended.
