# Mats — Label Formatter

## Running

```
python3 generate_labels.py
```

Run from the `Mats/` directory. Writes `labels_output.docx` in place.

Dependencies: `python-docx`, `openpyxl`

---

## File Access

| File | Path | Access |
|---|---|---|
| `MorphoData.csv` | `../~data/MorphoData.csv` | READ ONLY |
| `CollectionData.csv` | `CollectionData.csv` (local copy) | READ ONLY |
| `HistoricCollData.xlsx` | `../~data/HistoricCollData.xlsx` | READ ONLY |
| `labels_output.docx` | `labels_output.docx` | WRITE (only output) |

---

## Data Sources

### MorphoData.csv
One row per slide. Key columns: `Collection_ID`, `Slide_ID`, `Gen`, `Clade`, `Morph`.
- `Clade`: L, H, or N (H and N are collapsed into the H+N display group)
- `Morph`: `aptera` or `alate` (these are the operative values — not the Excel Morph column)
- `Collection_ID` may have trailing letter suffixes (`23-02a`, `23-02b`) — strip to get base ID
- Historical records appear here too; they are identified by cross-referencing with `HistoricCollData.xlsx`

### CollectionData.csv
One row per collection event. Key columns: `Collection ID`, `Location`, `Date`, `Host`, `Collector`, `Decimal Degrees GPS Coordinates`, `GenBank Accession`.
- `Date`: already in `12.viii.2023` format — use as-is
- `Location`: raw format `USA: VA, Clarke Co., Cool Spring` — normalized by `normalize_location()`

### HistoricCollData.xlsx
One row per historical slide. Columns: `Gen`, `Morph`, `Collection_ID`, `Slide_ID`, `Collector`, `Location`, `Date`, `Host`, `Institution`, `Collection Accession ID`, `Historic ID`, `Special`.
- `Collection_ID` here must match `Collection_ID` in MorphoData.csv — that is the join key
- `Special` holds curation notes (e.g. `"Remounted xi.1968 in balsam"`)
- `Morph` column in this file is NOT used for abbreviations — morph comes from MorphoData.csv

---

## Two Pipelines

### Modern (MorphoData × CollectionData)

1. Count slides per `(base_id, gen, clade, morph)` from MorphoData
2. Look up each base_id in CollectionData (`lookup_cid` handles zero-padded IDs like `23-082 → 23-82` and year-prefix aliases like `01-XX → Jan-XX`)
3. Collapse H/N into H+N display clade
4. Accumulate `gen_label_counts` per `(canonical_id, display_clade)` — e.g. `{'genI': 3, 'genValat': 2}`
5. Merge entries sharing the same county-level location + clade into one label
6. Sub-group by `(date, host)` within each merge bucket
7. Flag discrepancies (collector mismatch, GenBank mismatch) as plain text in the Flags section

**Modern label format:**
```
Location, ex Host, date, Collector, Coll. ID: XX-NN (genI count breakdown); ex Host2, ...
```
- Host italicized inline
- Gen breakdown in parentheses: `(2 genI, 1 genValat)`
- Sorted: country → state by total slides descending, then label slide count desc, then minimum ID

### Historical (HistoricCollData × MorphoData)

1. Build `hist_data` dict keyed by `Collection_ID` from the Excel file
2. While iterating MorphoData, if a base_id has no match in CollectionData but does match `hist_data`, it is a historical record — collect its slide IDs and metadata from the Excel row
3. Group historical raw entries by `(gen, morph, location, date, host, collector)` — accumulate slide IDs, accession numbers, and curation notes per group
4. Sort groups by gen numerically, then by location

**Historical label format:**
```
Location; date; ex Host; Collector; Slide_ID_1, Slide_ID_2; (n Gen N abbrev) (curation note) [INSTITUTION acc. NUM].
```
- Host italicized inline
- Count parenthetical uses gen+morph keyed abbreviation (see below)
- Curation note from `Special` column, lowercased first char, wrapped in parentheses
- If no accession: `[INSTITUTION]`; if accession present: `[INSTITUTION acc. NUM]`

---

## Gen+Morph Labels and Abbreviations

Morph values from MorphoData are `aptera`/`alate`. Gen I aptera = fundatrices (fp), not apt. Both the section headings and count abbreviations use `(gen, morph)` as the joint key.

| Gen | Morph (MorphoData) | Section Heading | Abbrev |
|---|---|---|---|
| 1 | aptera | Generation I — Fundatrices | fp |
| 2 | alate | Generation II — Alate Virginoparae | al. |
| 3 | aptera | Generation III — Coccidiformes | apt. |
| 3 | alate | Generation III — Coccidiformes | al. |
| 4 | aptera | Generation IV — Apterous Virginoparae | apt. |
| 5 | aptera | Generation V — Apterous Virginoparae | apt. |
| 5 | alate | Generation V — Alate Sexuparae | al. |

These are defined in `HIST_GEN_MORPH_HEADINGS` and `HIST_GEN_MORPH_ABBREV`. Add new rows there if new gen/morph combinations appear in the data.

---

## Document Structure

```
Clade L                              ← Heading 1
  [modern label paragraphs]

Clade H+N                            ← Heading 1
  [modern label paragraphs]

Historical Morphometric Specimens    ← Heading 1
  Generation I — Fundatrices         ← Heading 2
    [historical label paragraphs]
  Generation II — Alate Virginoparae ← Heading 2
    ...
  Generation V — Alate Sexuparae     ← Heading 2
    ...

Flags — Data Discrepancies           ← Heading 1 (omitted if no flags)
Flagged — No Match in CollectionData.csv ← Heading 1
```

---

## Clade Grouping

`CLADE_GROUPS = [('L', ['L']), ('H+N', ['H', 'N'])]`

H and N specimens are treated as one display group. The variable `ACTUAL_TO_DISPLAY_CLADE` maps raw clade values to their display label.

---

## Location Normalization (`normalize_location`)

Input format from CollectionData: `USA: VA, Clarke Co., Cool Spring`

Output format: `United States of America: Virginia: Clarke Co., Cool Spring`

- Country abbreviations expanded via `COUNTRY_NAMES` (`USA` → `United States of America`, `CA` → `Canada`)
- State/province abbreviations expanded via `SUBDIVISION_NAMES`
- DC special case: `USA: DC, Site` → `United States of America: Washington D.C., Site`
- County detected by `\bCo\.` pattern

Modern labels are merged at county level (`county_location()` strips everything after `Co.` to form the merge key).
