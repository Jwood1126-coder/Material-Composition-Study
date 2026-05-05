# Material Composition Analyzer

A small tool that audits **Material Composition** data on NetSuite items.
Feed it a CSV exported from a saved search and it produces a formatted Excel
report flagging the rows that look wrong.

---

## Download (Windows)

Grab the latest `.exe` from the
[**Releases page**](https://github.com/Jwood1126-coder/Material-Composition-Study/releases/latest)
and save it anywhere (Desktop is fine).

No Python install required — everything is bundled.

> **First-run note:** Windows SmartScreen may show
> *"Windows protected your PC"* because the binary is unsigned. Click
> **More info → Run anyway**. After the first launch it stops warning.

---

## Using the App

1. Double-click `MaterialCompositionAnalyzer.exe` — a window opens (no
   command prompt).
2. **Choose CSV…** → pick your NetSuite saved-search export.
3. Optionally narrow the **Checks to Run** (see below). Default is all checks.
4. Click **Analyze and Save Excel**.
5. The progress bar shows real moment-to-moment progress through CSV load →
   analyze → Excel write.
6. When done, click **Open Excel Report** or **Open Folder**.

The Excel file is saved next to the CSV with a timestamp, e.g.
`items_analysis_all_20260505_143012.xlsx`.

### Report scopes

The "Checks to Run" section lets you tackle one issue type at a time:

- **All** — every check enabled (default).
- **Suffix only** — just the name-suffix ↔ material rules. Run this first,
  fix the items it flags, then re-export.
- **Class only** — just the class-doesn't-reflect-material rule. Useful
  after suffix issues are clean (classes can't be auto-corrected, only
  flagged for human review).

The output filename includes the scope, so you can keep multiple reports
in the same folder without overwriting.

---

## What Gets Flagged

### Issues (highlighted red)

| Flag | Rule |
|---|---|
| **Suffix Mismatch** | `-SS` (or `-SS-<anything>`) → must be Stainless Steel; `-B` (exact) → must be Brass. Also fires the other way: composition is "Stainless Steel" but no `-SS` segment, or "Brass" but no trailing `-B`. |
| **Class Mismatch** | The Class name must contain a keyword reflecting the Material Composition (e.g. "Steel" composition → class must say "Steel" but not "Stainless"; "Stainless Steel" → class must say "Stainless"; "Brass" → "Brass"; etc.). |
| **Matrix Field Mismatch** | When the matrix `Material` field is populated, it must match (one of) the `Material Composition` value(s). |
| **Missing Composition** | `Material Composition` is blank. |

### Informational (noted but not flagged as errors)

| Flag | Note |
|---|---|
| **No-suffix Non-Steel** | Most bare parts are Steel. If a part has no material suffix and the composition isn't "Steel", it's noted for human review (could be intentional). |
| **Matrix Parent** | Same Internal ID appears on multiple rows → matrix parent. |
| **BB Part (Exempt)** | `-BB` segment or class is "Brennan Black". These are exempt from the class-contains-material check (the class itself is the program name). |

### Built-in safety guards

- `-BB` is **not** treated as `-B` (exact segment matching).
- "Steel" composition does **not** accept a "Stainless" class as a match.
- Matrix parents with multiple compositions are exempt from the reverse
  suffix check.
- Encoding fallback (utf-8-sig → latin-1) for CSVs that include the
  Excel BOM.

---

## Excel Output

The workbook contains:

- **Summary** — counts, issue rate, breakdown by issue type.
- **All Data** — every row, with the source columns plus
  `Recommended Material`, `Suffix-Detected Material`, `Analysis Notes`,
  and a column per flag.
- **Flagged Items** — only the rows with `Has Issue = YES`.
- **Per-flag sheets** — one sheet per issue type that has flagged rows
  (skipped when a single type would exceed 20K rows, to avoid bulk
  duplication).

Issue rows are highlighted red; YES cells in flag columns are bold red.
Auto-filter is enabled on every data sheet.

---

## Performance

On ~111K rows (a real-world export size), end-to-end is around
**30–40 seconds** on typical hardware:

- Analysis: ~1.5 s (vectorized via pandas)
- Excel write: ~30 s (xlsxwriter in constant_memory mode, conditional
  formatting instead of per-cell styling)

---

## Running from Source (optional)

If you want to run the script directly instead of the bundled `.exe`:

```bash
pip install -r requirements.txt

# GUI (same as the .exe)
python analyze_materials.py

# CLI
python analyze_materials.py items.csv --excel report.xlsx
python analyze_materials.py items.csv --excel report.xlsx --checks suffix
```

CLI flags:

- `--excel <path>` — also save a formatted Excel workbook.
- `--checks {suffix|class|matrix|empty|all}` — restrict to specific
  checks. Default `all`.
- `--flagged-only` — print only flagged rows.
- `--encoding <enc>` — override CSV encoding (default `utf-8-sig`).

---

## Building a New Release

Tag the commit and push:

```bash
git tag v1.4.0
git push origin v1.4.0
```

GitHub Actions (`.github/workflows/release.yml`) builds a Windows `.exe`
with PyInstaller and attaches it to a fresh GitHub Release on every `v*`
tag.
