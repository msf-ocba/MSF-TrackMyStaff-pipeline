# MSF Logistique → TMS Asset Import Pipeline

Automates the weekly process of pulling MSF Logistique's asset extraction
file off SFTP, cleaning/validating it, and producing a final import
workbook for TMS, plus a completion email summarising the batch.

This replaces the previous manual process (download → open in Excel →
Text Import Wizard → manual filtering → manual lookups → manual email).

## How it works, end to end

```
01_download_sftp.py   → downloads this week's extraction file from SFTP
02_read_files.py       → converts it to A1 (csv) → A2 (txt) → A3 (xlsx table)
03_clean_transform.py  → adds Family / Keep? columns           → A4
04_validate.py         → checks Article Codes against TMS      → A5 / A5B
                          (or, if nothing to import: A5B + A5_META only)
05_export.py           → maps to the TMS import template        → A6 / meta.json
                          (skipped if 04 found nothing to import)
06_notify.py           → sends the batch completion email
                          (always runs when there's a batch to report on -
                          either the normal A6 summary, or a shorter
                          "nothing to import" email - see below)
```

`run_pipeline.py` is the production entry point that runs all six steps
in order. The numbered scripts can also be run individually by hand for
debugging, each one auto-detects the most recent input file from the
previous stage if you don't pass a path explicitly.

## File naming convention

Every batch is identified by a **date tag** (`YYMMDD`, taken from the
source filename, e.g. `EXTRACTION_TMS_OCBA_260629.csv` → `260629`). This
tag is preserved in every derived file for that batch:

| File | Produced by | Contents |
|---|---|---|
| `A1_EXTRACTION_TMS_OCBA_YYMMDD.csv` | `02` | Renamed copy of the source extract |
| `A2_EXTRACTION_TMS_OCBA_YYMMDD.txt` | `02` | Text-import-wizard equivalent |
| `A3_EXTRACTION_TMS_OCBA_YYMMDD.xlsx` | `02` | Excel table, Serial Number forced to text |
| `A4_EXTRACTION_TMS_OCBA_YYMMDD.xlsx` | `03` | + `Family` and `Keep?` columns |
| `A5_EXTRACTION_TMS_OCBA_YYMMDD.xlsx` | `04` | Rows with blank `Keep?` + Article Code Check. **Not produced** if zero rows have a blank `Keep?` — see "Nothing to import" below |
| `A5B_EXTRACTION_TMS_OCBA_YYMMDD.xlsx` | `04` | Rows excluded by `Keep?` (kept for visibility). Always produced whenever there's at least one excluded row |
| `A5_META_EXTRACTION_TMS_OCBA_YYMMDD.json` | `04` | **Only** produced when A5 is skipped (nothing to import) — lighter handoff summary for `06_notify.py`, see below |
| `A6_EXTRACTION_TMS_OCBA_YYMMDD.xlsx` | `05` | **Final TMS import file** |
| `A6_META_EXTRACTION_TMS_OCBA_YYMMDD.json` | `05` | Handoff summary used by `06_notify.py` |

### Nothing to import

Some batches have zero rows with a blank `Keep?` — every item was
already excluded as a Kit (`NOT ASSET`), undeployed (`NOT DEPLOYED`), or
both. In that case there's nothing new for TMS, so the pipeline doesn't
try to build an empty A5/A6:

- `04_validate.py` skips the article-code/TMS lookup entirely, still
  writes `A5B` (so the excluded rows remain visible for the audit
  trail), writes `A5_META_....json` instead of `A5`, and exits with
  code `3` ("nothing to import" — not an error).
- `run_pipeline.py` sees exit code `3`, skips `05_export.py` (nothing
  to export), and runs `06_notify.py` directly. Users depend on
  getting a completion email either way, not on checking server logs,
  so the email still goes out — just a shorter one confirming the
  batch ran and explaining why nothing was imported, with A4/A5B
  attached for reference.
- Running `04_validate.py` by hand in this situation also exits `3`;
  running `06_notify.py` afterward (or letting `run_pipeline.py` do it)
  picks up `A5_META_....json` automatically, the same way it normally
  picks up `A6_META_....json`.

### Folder tag vs. file tag

Folders (`downloads/`, `work/`, `output/`) are **not** dated by the
batch's own `YYMMDD`. They're dated by a shared **ISO week tag**
(`compute_week_tag()` in `common_config.py`), e.g. `Y26W25`, which is
always the ISO week *before* the one the pipeline is run in — MSF
Logistique's extract for a given week only lands (and gets processed)
the following week. All stages use the same tag so a batch's files
always land in matching folders:

```
downloads/Y26W25/EXTRACTION_TMS_OCBA_260622.csv
work/Y26W25/A3_EXTRACTION_TMS_OCBA_260622.xlsx
work/Y26W25/A4_EXTRACTION_TMS_OCBA_260622.xlsx
output/Y26W25/A5B_EXTRACTION_TMS_OCBA_260622.xlsx
output/Y26W25/A6_EXTRACTION_TMS_OCBA_260622.xlsx
output/Y26W25/A6_META_EXTRACTION_TMS_OCBA_260622.json
```

Re-running a step later in the same week reuses the same folder rather
than scattering files across timestamped ones.

## Step-by-step detail

### 01 — `01_download_sftp.py`
Connects to the SFTP server, lists files matching `file_prefix`/`file_suffix`
in `remote_dir`, and picks the latest by the `YYMMDD` in its filename.
**Only downloads if that file is from the current ISO week** ,otherwise
it exits cleanly, since this usually just means the weekly extract hasn't
landed yet.

Exit codes: `0` downloaded, `1` real error (bad filename, connection
failure), `2` not ready yet (no matching files, or latest file is stale).

### 02 — `02_read_files.py`
Finds the latest downloaded file (any dated subfolder) and produces
A1 → A2 → A3, reproducing the manual Text Import Wizard steps: semicolon
delimiter, ISO-8859-1 source encoding, and identifier columns
(`PCL_NO_SERIE_LOT`, `FCL_ART_CODE`, etc.) forced to text so leading
zeros / non-numeric values survive.

### 03 — `03_clean_transform.py`
Adds two columns to A3 → A4:

- **Family** — first 4 characters of `FCL_ART_CODE`. Values starting
  with `K` (Kit family) are flagged.
- **Keep?** — records why a row should be excluded:
  - `NOT ASSET` — it's a Kit
  - `NOT DEPLOYED` — customer code doesn't match `customer_code` in
    config, or the mission code prefix isn't in
    `Mission_code_prefixes.txt`
  - Blank = passed all checks
  - A row can carry both reasons at once

A4 is saved to the batch's work folder **and** copied to the output
folder. Rows are colour-coded by `Keep?` status (red = both reasons,
purple = Kit only, yellow = not deployed only, gray = pending/blank).

### 04 — `04_validate.py`
Splits A4 by `Keep?`:
- Blank rows → looked up against `templates/TMS_UniDataArticles.xlsx`
  (via `templates/ArticleCompose.xlsx` for codes that need translating
  first) → **A5**, with `Article Code Check` = `Checked`/`Review`.
- Non-blank rows → **A5B**, keeping A4's original colour-coding.

If there are zero blank rows, A5 and the article-code lookup are
skipped entirely — see "Nothing to import" above. A5B is still written.

Exit codes: `0` A5 produced normally, `3` nothing to import this run
(not an error — see above), and a real failure otherwise (uncaught
exception).

### 05 — `05_export.py`
Maps A5 onto the TMS import template's column layout → **A6**, the file
actually imported into TMS. Also runs three checks (informational only,
none of them block export):

- **Article Code Check** — rows still marked `Review`
- **Location Check** — looks up the mission prefix in
  `templates/location.xlsx`; leaves `Location` blank if no match
- **Duplicate Check** — flags rows sharing both Article Code *and*
  Serial Number with another row in the batch

Rows are colour-coded (green/clean, yellow/review, orange/missing
location, red/duplicate — most severe wins). `Model` is truncated to 50
characters (TMS's own import limit). A companion
`A6_META_EXTRACTION_TMS_OCBA_YYMMDD.json` records the batch stats for
`06_notify.py`. Not run at all if `04_validate.py` found nothing to
import that batch.

### 06 — `06_notify.py`
Reads the most recent meta JSON — either `A6_META_....json` (the normal
case, written by `05_export.py`) or `A5_META_....json` (the "nothing to
import" case, written by `04_validate.py` instead — see above) — and
sends one completion email:

- **Normal case:** attaches A6 (+ A4 and A5B if present) and summarises
  the batch, including a dedicated section for Kit articles that
  reached a deployed mission and still need manual breakdown into
  component assets before import.
- **Nothing-to-import case:** a shorter email confirming the batch ran,
  explaining that every row already had a non-blank `Keep?`, and
  attaching A4/A5B (no A6 exists for this batch).

If `[notifications] smtp_host` is blank in `config.conf`, the email
content is printed to the console instead of sent — safe to run before
SMTP is configured.

## Running it

**Production (cron/systemd):**
```
python run_pipeline.py
```

**Debugging a single step:**
```
python run_pipeline.py --only 03
```

**See the plan without running anything:**
```
python run_pipeline.py --dry-run
```

### `run_pipeline.py` exit codes

| Code | Meaning |
|---|---|
| `0` | Full success (including the completion email, whether it was the normal A6 summary or a "nothing to import" email), OR nothing new to process yet from SFTP, OR another run was already in progress (lock held) — none need a human |
| `1` | A real failure in steps 01–04, or in 05 when there was something to export — A6 may be missing/stale, needs investigation |
| `2` | The data pipeline succeeded (A6 was produced and is importable, or the batch genuinely had nothing to import) but the completion email (06) failed — lower urgency; re-run `python 06_notify.py` once fixed |

Note that a "nothing to import" batch (04_validate.py exit code 3)
still results in `run_pipeline.py` running `06_notify.py` — 05 is the
only step skipped, since there's nothing for it to export.

The pipeline takes an exclusive lock (`.pipeline.lock`) so overlapping
cron runs can't process the same batch twice, and each step gets a
30-minute timeout so a hung SFTP connection can't wedge future runs.
Only step 01 is retried automatically (3 attempts, 30s apart) — steps
02–06 operate on local files, so a repeat failure is almost always a
real problem a retry won't fix. (04's "nothing to import" exit code is
likewise never retried — it's a normal outcome, not a transient error.)

## Setup

1. Install dependencies:
   ```
   pip install paramiko pandas openpyxl
   ```
2. Copy/edit `config.conf` (see below) with real SFTP credentials and
   paths. **Keep this file out of version control.**
3. Populate the reference files under `templates/`:
   - `TMS_UniDataArticles.xlsx` — TMS's current article export, with an
     `Article Code` column
   - `ArticleCompose.xlsx` — maps raw extraction codes to the TMS code
     for "hidden" Kit articles, columns `Article` / `ArticleCompose`
   - `location.xlsx` — one full TMS location path per row (used to
     derive the 2-letter mission prefix → location mapping)
   - `Mission_code_prefixes.txt` — one valid 2-character mission code
     prefix per line (`#` for comments)
4. Fill in `[notifications]` in `config.conf` once ready to send real
   emails; leave `smtp_host` blank to just print the notification.
5. Schedule `run_pipeline.py` via cron/systemd timer (e.g. daily —
   step 01 will simply exit with "not ready yet" on days the weekly
   extract hasn't landed).

### `config.conf` reference

| Section | Key | Purpose |
|---|---|---|
| `[sftp]` | `host`, `port`, `user`, `password` / `private_key_path` | Connection to MSF Logistique's SFTP server |
| | `remote_dir` | Remote folder the extract is dropped in |
| | `file_prefix`, `file_suffix` | Filter for which remote files count as candidates |
| `[local]` | `download_root`, `work_dir`, `output_dir` | Local folders (auto-created) |
| `[extraction_rules]` | `customer_code` | The `FCT_CLI_CODE_FAC` value identifying OCBA |
| | `excluded_art_faa_code` | Article family code to exclude |
| `[templates]` | `tms_article_list_path`, `article_compose_path`, `location_list_path`, `mission_code_prefixes_path` | Paths to the reference files above |
| `[notifications]` | `smtp_host`, `smtp_port`, `smtp_user`, `smtp_password`, `use_tls`, `from_address`, `to_addresses` | Completion email settings |

## Logs

Every `run_pipeline.py` run writes a timestamped log to `./logs/`
(in addition to console output), so there's a permanent record of what
ran and why, independent of cron's own output capture.

## Requirements

- Python 3.8+
- `paramiko` (SFTP)
- `pandas` (data processing)
- `openpyxl` (Excel read/write, tables, conditional formatting)

## Known open items
- Article Code `Review` rows, missing-location rows, and duplicate
  rows are all **included** in A6 rather than blocked — TMS import
  still needs a human to review the highlighted rows afterward.
