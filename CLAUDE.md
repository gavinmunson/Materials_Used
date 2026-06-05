# Label Formatter — Claude Code Prompt

## FILE ACCESS RULES
- `data.csv` → **READ ONLY**
- `CollectionData.csv` → **READ ONLY**
- Output: create a **new `labels_output.docx`** — this is the ONLY file you may write to. Do not modify any existing file.

---

## INPUT FILE STRUCTURE

**`data.csv` columns:**
`Collection_ID` | `Gen` | `Clade` | `Morph` | `Slide_ID`

**`CollectionData.csv` columns:**
`Collection ID` | `GenBank Accession` | `Date` | `Location` | `Decimal Degrees GPS Coordinates` | `Host` | `Collector`

---

## STEP 1 — Parse and Group `data.csv`

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

## STEP 2 — Cross-reference `CollectionData.csv`

- Match each `Base_ID` to its row in `CollectionData.csv` via the `Collection ID` column.
- Extract: `Date`, `Location`, `Host`, and all other columns.
- **Date:** already in `12.viii.2023` format — use as-is.
- **Location:** normalize punctuation to exactly:
  - `[Country]: [State]: [County]. [Site name]`
  - Colons after Country and State; period after County; no trailing punctuation after site name.
  - Example: `USA: VA: Clark Co. Shenandoah Univ. River Campus at Cool Spring Battlefield`
  - Special case — no county (e.g., Washington D.C.): `USA: Washington D.C. [Site name]`
- If a `Base_ID` has **no match** in `CollectionData.csv`: skip from main output, log to flagged section (see Step 4).

---

## STEP 3 — Merge Entries

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

## STEP 4 — Format Each Label String

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

## STEP 5 — Write `labels_output.docx`

### Document Structure

Organize with Word heading styles:

```
Clade [L / H / N]         ← Heading 1
  Generation [1–5]         ← Heading 2
    Morph [value]          ← Heading 3
      [label paragraph]
      [label paragraph]
      ...
```

- Clade order: **L, H, N**.
- Generation order: **1, 2, 3, 4, 5**.
- Within each Clade/Gen/Morph section, sort labels by their lowest Base_ID numerically.
- Each label = its own paragraph.
- Blank line between Morph sections for readability.
- Omit any Clade/Gen/Morph combination with no entries.

### Flagged Section

At the end of the document, add a **"Flagged — No Match in CollectionData.csv"** section listing any Base_IDs that could not be matched, one per line.

---

## DEPENDENCIES

Use `python-docx` for `.docx` output, including inline italic runs and Word comments.
