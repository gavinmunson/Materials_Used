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
- `Historic ID` is the canonical slide identifier; multiple rows with the same `Historic ID` are multiple insects on one slide (letter suffix pattern in `Slide_ID`)
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
Location, ex Host (ital.), date, Collector, Coll. ID: XX-NN (n gen. I fund., m gen. V al. sp.); ex Host2, ...
```

### Historical (HistoricCollData × MorphoData)

1. Build `hist_data` dict keyed by `Collection_ID` from the Excel file
2. While iterating MorphoData, if a base_id has no match in CollectionData but does match `hist_data`, it is a historical record
3. Group by `Historic_ID` to collapse letter-variant rows; count specimens per slide group
4. Detect lectotypes by `Collection_ID` prefix (`'Lectotype:'`) → pulled into separate Type Material section
5. Sort by gen numerically, then by location

**Historical label format:**
```
Location, Host (ital.), date, Collector (n morph; IDs) (curation note) [INSTITUTION acc. NUM].
```
- Count parenthetical uses gen+morph keyed abbreviation (see below)
- Curation note from `Special` column, lowercased first char, wrapped in parentheses — omitted if empty
- If no accession: `[INSTITUTION]`; if accession present: `[INSTITUTION acc. NUM]`

---

## Gen+Morph Labels and Abbreviations

Morph values from MorphoData are `aptera`/`alate`. Gen I aptera = fundatrices. Both section headings and count abbreviations use `(gen, morph)` as the joint key.

| Gen | Morph (MorphoData) | Section Heading | Abbrev |
|---|---|---|---|
| 1 | aptera | Generation I — Fundatrices | gen. I fund. |
| 2 | alate | Generation II — Alate Virginoparae | gen. II al. vp. |
| 4 | aptera | Generation IV — Apterous Virginoparae | gen. IV apt. vp. |
| 5 | aptera | Generation V — Apterous Virginoparae | gen. V apt. vp. |
| 5 | alate | Generation V — Alate Sexuparae | gen. V al. sp. |

These are defined in `HIST_GEN_MORPH_HEADINGS` and `HIST_GEN_MORPH_ABBREV`. Add new rows there if new gen/morph combinations appear in the data.

---

## Document Structure

```
Clade L                              ← Heading 1
  [modern label paragraphs]

Clade H+N                            ← Heading 1
  [modern label paragraphs]

Type Material                        ← Heading 1 (omitted if no lectotypes)
  Lectotype                          ← Heading 2
    [lectotype label paragraph]

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

H and N specimens are treated as one display group.

---

## Location Normalization (`normalize_location`)

Input format from CollectionData: `USA: VA, Clarke Co., Cool Spring`

Output format: `United States of America: Virginia: Clarke Co., Cool Spring`

- Country abbreviations expanded via `COUNTRY_NAMES`
- State/province abbreviations expanded via `SUBDIVISION_NAMES`
- Modern labels merged at county level (`county_location()` strips everything after `Co.`)
