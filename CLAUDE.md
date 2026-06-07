# Label Formatter — Claude Code Prompt

## FILE ACCESS RULES
- `data.csv` → **READ ONLY**
- `CollectionData.csv` → **READ ONLY**
- `HistoricCollData.xlsx` → **READ ONLY**
- Output: create a **new `labels_output.docx`** — this is the ONLY file you may write to. Do not modify any existing file.

---

## OVERVIEW — TWO PARALLEL PIPELINES, ONE OUTPUT

This script processes **two independent datasets** and writes both into `labels_output.docx`:

| Pipeline | Source files | Label source tag |
|---|---|---|
| **Modern** | `data.csv` + `CollectionData.csv` | *(existing, unchanged)* |
| **Historical** | `HistoricCollData.xlsx` | *(new)* |

Both pipelines produce labels in the same format and are organized under the same Clade → Generation → Morph heading structure in the output document. Within each Clade/Gen/Morph section, modern labels appear first (sorted by Base_ID), followed by historical labels (sorted by date, then by their ID). Each historical label is tagged with `[HIST]` in the flagged section only — **not** in the label text itself.

Run both pipelines fully before writing any output. Merge their results into the unified document structure described in Step 6.

---

## MORPH ABBREVIATIONS (Blackman & Eastop conventions)

Use these abbreviations consistently across both pipelines wherever morph values appear:

| Abbreviation | Full term |
|---|---|
| `fp` | fundatrix / fundatrices |
| `fg` | fundatrigeniae |
| `al.` | alatae (winged) |
| `apt.` | apterae (wingless) |
| `sexup.` | sexuparae |
| `ovip.` | oviparae |
| `L1`, `L2`, `L3`, `L4` | larval instars (1st–4th) |

**Normalization rules:**
- If morph column says `"sexuparae alates"` → write `sexuparae al.`
- If morph column says `"migrant sexuparae alates"` → write `migrant sexuparae al.`
- If morph column says `"sexuparae"` alone (no wing state) → write `sexup.`
- If morph column says `"oviparae"` → write `ovip.`
- Keep any generation notation (e.g. `generation V`) alongside the morph abbreviation when present.
- Unrecognized morph values: use as-is but flag in the Flagged — Historical Data Issues section.

---

## PIPELINE A — MODERN DATA (existing, unchanged)

### A-1 — Parse and Group `data.csv`

- Read all rows from `data.csv`.
- `Collection_ID` values come in two formats:
  - **Standard:** `23-01`, `23-105`
  - **Lettered:** `23-02a`, `23-02b`, `23-02c`
- **Group lettered IDs by their base ID** (strip the trailing letter). Count the number of variants = that base ID's slide count.
  - e.g., `23-02a`, `23-02b`, `23-02c` → base `23-02`, slide count = 3
- Standard IDs = slide count of 1.
- Lettered variants of the same base ID are guaranteed to share identical `Gen`, `Clade`, `Morph`, and collection data.
- For each base ID, record: `Base_ID`, `Gen`, `Clade`, `Morph`, `Slide_Count`.

---

### A-2 — Cross-reference `CollectionData.csv`

- Match each `Base_ID` to its row in `CollectionData.csv` via the `Collection ID` column.
- Extract: `Date`, `Location`, `Host`, and all other columns.
- **Date:** already in `12.viii.2023` format — use as-is.
- **Location:** normalize punctuation to exactly:
  - `[Country]: [State]: [County]. [Site name]`
  - Colons after Country and State; period after County; no trailing punctuation after site name.
  - Example: `USA: VA: Clark Co. Shenandoah Univ. River Campus at Cool Spring Battlefield`
  - Special case — no county (e.g., Washington D.C.): `USA: Washington D.C. [Site name]`
- If a `Base_ID` has **no match** in `CollectionData.csv`: skip from main output, log to flagged section (see Step 6).

---

### A-3 — Merge Modern Entries

Group base IDs that share **all five** of the following:
- Same `Location`
- Same `Date`
- Same `Gen`
- Same `Clade`
- Same `Morph`

These become a **single merged label.**

**Within a merged group:**
- Total slide count = sum of all `Slide_Count` values across all base IDs in the group.
- Sort base IDs numerically (lower numbers first; `23-` before `24-` etc.).
- Group base IDs by `Host`. Within each host group, list IDs in numerical order.
- Format host clauses as: `ex [Host]; [ID-1]; [ID-2];` — only write `ex [Host]` once per host group, then list all IDs sharing that host.
- Order host groups by the lowest Base_ID in each group.
- If any non-Host, non-ID column differs unexpectedly between entries being merged (e.g., Collector, GPS coordinates, GenBank Accession, or anything else), attach a **Word comment** to that label noting the discrepancy. Do not block the merge — proceed and flag.

---

### A-4 — Format Modern Label Strings

Use this exact format (one paragraph per label, host name(s) italicized inline):

```
[N] slides: [Location]; [Date]; ex [Host A]; [ID-1]; [ID-2]; ex [Host B]; [ID-3].
```

- Spell out slide count as a word: One, Two, Three … (use numerals above ten).
- **Drop letter suffixes** from all IDs in output — write `23-02`, never `23-02a`.
- Italicize each host plant name inline (e.g., *H. virginiana*).
- Separate all elements with **semicolons**.
- End the entire label with a **period**.

**Examples:**

Single entry, standard ID:
> One slide: USA: VA: Clark Co. Shenandoah Univ. River Campus at Cool Spring Battlefield; 12.viii.2023; ex *H. virginiana*; 23-105.

Merged, multiple hosts:
> Three slides: USA: Washington D.C. US National Arboretum; 10.vii.2023; ex *H. virginiana*; 23-53; ex *H. vernalis*; 23-82; 23-83.

---

## PIPELINE B — HISTORICAL DATA (`HistoricCollData.xlsx`)

### B-1 — Inspect the Excel file first

Before any parsing, do the following:
1. Read all sheet names in `HistoricCollData.xlsx` and print them.
2. Read the header row of each sheet and print column names.
3. Print 3 sample rows from the sheet that appears to contain collection data.
4. **Pause and report** what you found. Map the discovered columns to the fields below
   using your best judgment, then state your mapping explicitly before proceeding.
   If any required field is missing or ambiguous, flag it rather than guessing silently.

**Expected fields to locate** (column names may differ — match by content, not exact name):

| Required field | Likely column name(s) | Notes |
|---|---|---|
| Record/specimen ID | `Slide_ID`, `S#`, `Catalog #`, `ID` | May be prefixed `S1.XXXX` |
| Generation | `Gen`, `Generation` | Integer 1–5 |
| Clade | `Clade`, `Group` | L, H, or N |
| Morph | `Morph`, `Life stage` | e.g. sexuparae, fp, al. |
| Slide count | `Slide_count`, `# slides`, `Count` | May need to be inferred — see B-2 |
| Country | `Country` | |
| State/Province | `State`, `Province` | |
| County | `County`, `Co.` | May be absent |
| Locality/Site | `Locality`, `Site`, `Location` | |
| Host plant | `Host`, `Host plant`, `Host species` | |
| Date | `Date`, `Collection date` | See date normalization below |
| Collector(s) | `Collector`, `Collected by` | May be multiple names |
| Determiner | `Det`, `Determiner`, `Identified by` | May be absent or same as collector |
| Repository | `Repository`, `Collection`, `INHS`, `Depository` | e.g. INHS, CNC, USNM |
| Accession number | `Accession`, `Acc #`, `INHS #` | May be absent |
| Type status | `Type`, `Type status` | holotype/paratype/lectotype/none |
| Curation note | `Note`, `Curation`, `Remarks` | e.g. "remounted xi.1968" |

---

### B-2 — Parse and Group Historical Records

**Slide ID handling:**
- Historical slide IDs may look like: `S1.5791`, `S1.5792`, `INHS-1042`, or plain integers.
- Multiple slide IDs for the same collection event may be stored as:
  - Separate rows with the same collection data but different Slide_IDs, OR
  - A single row with multiple IDs in one cell (comma- or semicolon-separated)
- If stored as **separate rows**: group rows that share identical Location + Date + Gen + Clade + Morph + Host + Collector into one record. Slide count = number of rows grouped. Collect all Slide_IDs.
- If stored as **one cell with multiple IDs**: split on commas/semicolons. Slide count = number of IDs found.
- If slide count column exists and is populated: use it directly. Cross-check against ID count if possible.

**Lettered suffix handling (if present):**
- Same rule as Pipeline A: strip trailing letter to get base ID, count variants.

**Collector name formatting:**
- Format as: `Initials. Surname` (e.g., `F.C. Hottes`, `E. Maw`, `T.H. Frison`).
- Multiple collectors joined with `&`: `T.H. Frison & F.C. Hottes`.
- Do not write a "Coll." prefix.
- Determiner: include only if different from collector, written as `(det. F.C. Hottes)` immediately after the collector name.
- If Determiner is empty or identical to Collector: omit the `(det. X)` entirely.

For each resolved record, store:
`Base_ID`, `Slide_IDs` (list), `Slide_Count`, `Gen`, `Clade`, `Morph`,
`Location_normalized`, `Date_normalized`, `Host`, `Collector`, `Determiner`,
`Repository`, `Accession`, `Type_Status`, `Curation_Note`

---

### B-3 — Normalize Historical Location

Apply the same location format as Pipeline A:

```
[Country]: [State]: [County]. [Site name]
```

- If Country column says "United States of America" or "USA" or "U.S.A." → normalize to `USA`
- If County is absent or empty → omit it: `[Country]: [State]: [Site name]`
- If only one geographic level below country is present → treat as State, omit County
- Preserve site name punctuation (abbreviations like "Univ.", "Co." are fine within the site name)

---

### B-4 — Normalize Historical Dates

Historical dates may arrive in various formats. Normalize ALL of them to `day.roman_month.year`:

| Input format | Normalized output |
|---|---|
| `4.vi.1928` | `4.vi.1928` *(already correct)* |
| `June 4, 1928` | `4.vi.1928` |
| `1928-06-04` | `4.vi.1928` |
| `6/4/1928` | `4.vi.1928` |
| `4 June 1928` | `4.vi.1928` |
| `vi.1928` (no day) | `vi.1928` *(day unknown — omit day only)* |
| `1928` (year only) | `1928` *(use year only)* |

Month → roman numeral map:
i=Jan, ii=Feb, iii=Mar, iv=Apr, v=May, vi=Jun, vii=Jul, viii=Aug, ix=Sep, x=Oct, xi=Nov, xii=Dec

---

### B-5 — Handle Type Specimens

If a historical record has a non-empty `Type_Status` field (holotype, lectotype, paratype, paralectotype, neotype, allotype):

- **Do NOT include it in the main label output sections.**
- Collect all type records into a separate list for the Type Material section (see Step 6).
- Format each type label as:

```
[TYPE_STATUS] ♀ (morph): [Location]; [Date]; ex [Host]; [Slide_IDs]; [Collector] (det. [Determiner]); [Curation_Note if present] [REPOSITORY acc. AccessionNumber].
```

  - Type status word is bold.
  - Sex symbol (♀/♂) immediately follows type status — use ♀ as default if sex not recorded, and attach a Word comment noting sex was not specified in source data.
  - If Type_Status is `lectotype` or `neotype`: append `designated here` before the repository bracket unless a `Designation_Reference` field is present, in which case use `designated by [Author Year]`.
  - Paralectotypes/paratypes from multiple repositories: list each repository's specimens as a sub-entry, separated by semicolons, each with its own repository bracket.

---

### B-6 — Handle Curation Notes

If `Curation_Note` is non-empty, append it as a parenthetical after the specimen count and before the repository bracket:

```
[N] slides: [Location]; [Date]; ex [Host]; [IDs]; (remounted xi.1968) [REPOSITORY].
```

Common curation note patterns to recognize and standardize:
- "remounted [date]" → `(remounted [date])`
- "remounted by [Name], [date]" → `(remounted by [Name], [date])`
- "slide damaged" → `(slide damaged)`
- "originally in alcohol, slide prepared [date]" → keep as-is in parentheses
- "old slide [ID] → new slide [ID]" → `(remounted as [new ID], [date])`

---

### B-7 — Merge Historical Entries

Apply the same merge logic as Pipeline A (Step A-3).

Group records sharing **all five** of:
- Same `Location_normalized`
- Same `Date_normalized`
- Same `Gen`
- Same `Clade`
- Same `Morph`

Within a merged group:
- Total slide count = sum of `Slide_Count` values.
- Collect all `Slide_IDs` across the group.
- Group IDs by Host — format as: `ex [Host]; [ID-1]; [ID-2];`
- Order host groups by first-appearing ID.
- If Curation_Note differs between merged records: include all unique notes, comma-separated in the parenthetical.
- If any other field conflicts (Collector, Repository, Accession): attach a Word comment. Do not block the merge.

---

### B-8 — Format Historical Label Strings

Use the **same label format** as Pipeline A:

```
[N] slides: [Location]; [Date]; ex [Host]; [Slide_IDs]; (Curation_Note if present) [REPOSITORY acc. AccessionNumber].
```

- Slide count spelled out as a word (numerals above ten).
- Host plant italicized inline.
- Slide IDs listed after host (unlike modern labels which use collection field numbers — historical labels use the actual slide IDs since those are the primary identifiers).
- If multiple repositories: `[INHS acc. 1092556; CNC]`
- If no accession number: `[REPOSITORY]`
- Semicolons throughout; period at end.

**Historical label example:**
> Two slides: USA: IL: Carbondale. *Betula nigra*; 4.vi.1928; ex *Betula nigra*; S1.5791; S1.5792; (det. F.C. Hottes) [INHS acc. 1092556].

---

## STEP 6 — Write `labels_output.docx`

### Document Structure

```
[TYPE MATERIAL]               ← if any type specimens exist (Historical only)
  **Holotype**
  **Lectotype**
  **Paratypes / Paralectotypes**

Clade [L / H / N]            ← Heading 1
  Generation [1–5]            ← Heading 2
    Morph [value]             ← Heading 3
      [Modern label paragraph(s)]
      [Historical label paragraph(s)]

Flagged — No Match           ← end of document
```

**Ordering rules:**
- Clade order: **L, H, N**
- Generation order: **1, 2, 3, 4, 5**
- Within each Clade/Gen/Morph section:
  - Modern labels first, sorted by lowest Base_ID numerically
  - Historical labels after, sorted by `Date_normalized` (oldest first), then by Slide_ID
- Blank line between Morph sections for readability
- Omit any Clade/Gen/Morph combination with no entries

**Type Material section** (if present):
- Appears at the very top of the document, before Heading 1 sections
- Heading: `Type Material` (Heading 1 style)
- Sub-sections: `Holotype`, `Lectotype`, `Paratypes`, `Paralectotypes` — each Heading 2
- One label per line; bold type status word inline

### Flagged Section

At the end of the document, add a **"Flagged"** section with two sub-sections:

**"Flagged — No Match in CollectionData.csv"**
- Modern Base_IDs with no matching row in CollectionData.csv, one per line

**"Flagged — Historical Data Issues"**
- Historical records that could not be resolved, one per line
- Format: `[Slide_ID or row number]: [reason]`
  - e.g. `S1.5791: Gen value missing`
  - e.g. `Row 14: Location field empty`
  - e.g. `S2.0042: Clade not recognized (value: "X")`
  - e.g. `S3.0012: Morph value unrecognized (value: "adult winged") — used as-is`

---

## DEPENDENCIES

Use `python-docx` for `.docx` output, including:
- Inline italic runs (host plant names)
- Bold runs (type status words)
- Word comments (merge discrepancies, missing sex on type specimens)
- Heading styles (Heading 1, Heading 2, Heading 3)
- `openpyxl` for reading `HistoricCollData.xlsx`

Install if needed:
```
pip install python-docx openpyxl
```
