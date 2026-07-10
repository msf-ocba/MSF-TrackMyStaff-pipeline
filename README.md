# MSF Logistique -> TMS Asset Import Pipeline

Automates the weekly process of pulling MSF Logistique's asset extract off
SFTP, cleaning/validating it, and producing a TMS-ready import file (plus a
completion email), replacing what used to be a manual Excel workflow.

## Contents

- [Pipeline overview](#pipeline-overview)
- [File & folder naming](#file--folder-naming)
- [Setup](#setup)
- [Running manually](#running-manually)
- [Running via the master script](#running-via-the-master-script)
- [Production / cron deployment](#production--cron-deployment)
- [Troubleshooting](#troubleshooting)
- [Known follow-ups](#known-follow-ups)

## Pipeline overview

Six numbered scripts, each consuming the previous one's output:

```
01_download_sftp.py   SFTP -> downloads/<week_tag>/EXTRACTION_TMS_OCBA_YYMMDD.csv
        |
        v
02_read_files.py       -> work/<week_tag>/A1_*.csv, A2_*.txt, A3_*.xlsx
        |               (spec section 4, steps 3-6: rename, text-import, table)
        v
03_clean_transform.py  -> work/<week_tag>/A4_*.xlsx   (+ copy in output/<week_tag>/)
        |               (adds Family + Keep? columns; spec step 8 pt.1)
        v
04_validate.py         -> work/<week_tag>/A5_*.xlsx  (kept rows, Article Code checked)
                        -> output/<week_tag>/A5B_*.xlsx (excluded rows)
        |
        v
05_export.py           -> output/<week_tag>/A6_*.xlsx        (final TMS import file)
                        -> output/<week_tag>/A6_META_*.json  (handoff for 06)
        |
        v
06_notify.py           -> completion email (or console printout), with
                           A6 + A5B + A4 attached where available
```

`run_pipeline.py` (the master script - see below) runs all six in order and
is what production/cron should actually call.

## File & folder naming

**Folders use a "week_tag" like `Y26W25`, and it is always LAST week, not
the week the script runs in.** MSF Logistique's weekly extract for a given
week only lands on the SFTP server (and gets processed by this pipeline)
during the *following* week. So if this pipeline runs during ISO week 26,
every folder it touches is tagged `Y26W25`, not `Y26W26`.

This is computed in exactly one place - `compute_week_tag()` in
`common_config.py` - by subtracting 7 days from "now" and taking the ISO
week of that. Every script (01, 02, 04, 05) imports this same function, so
there is no risk of the scripts drifting out of sync with each other, and
there is nowhere else in the codebase that recomputes it differently.

Every stage writes into a dated subfolder - **nothing is ever dropped flat**
into `download_root`, `work_dir`, or `output_dir` themselves:

```
downloads/Y26W25/EXTRACTION_TMS_OCBA_260622.csv
                  downloaded_files.txt
work/Y26W25/A1_EXTRACTION_TMS_OCBA_260622.csv
            A2_EXTRACTION_TMS_OCBA_260622.txt
            A3_EXTRACTION_TMS_OCBA_260622.xlsx
            A4_EXTRACTION_TMS_OCBA_260622.xlsx
            A5_EXTRACTION_TMS_OCBA_260622.xlsx
output/Y26W25/A4_EXTRACTION_TMS_OCBA_260622.xlsx   (copy)
              A5B_EXTRACTION_TMS_OCBA_260622.xlsx
              A6_EXTRACTION_TMS_OCBA_260622.xlsx
              A6_META_EXTRACTION_TMS_OCBA_260622.json
```

Note the two different tags in play:
- **`week_tag`** (e.g. `Y26W25`) names the *folder* - shared by every batch
  processed during the same calendar week.
- **`date_tag`** (e.g. `260622`, the YYMMDD embedded in the source
  filename) names the *files* (A1-A6, A5B) - identifies one specific batch,
  in case more than one ever lands in the same week folder.

### File glossary

| File | Produced by | Contents |
|---|---|---|
| `EXTRACTION_TMS_OCBA_YYMMDD.csv` | 01 | Raw MSF Logistique export, as downloaded |
| `A1_*.csv` | 02 | Renamed/re-saved raw extract |
| `A2_*.txt` | 02 | Same data, `.txt` copy (text-import-wizard equivalent) |
| `A3_*.xlsx` | 02 | Excel table, `PCL_NO_SERIE_LOT` forced to text |
| `A4_*.xlsx` | 03 | A3 + `Keep?` / `Family` columns (asset/deployment filter) |
| `A5_*.xlsx` | 04 | A4 rows kept (`Keep?` blank), + Article Code Check |
| `A5B_*.xlsx` | 04 | A4 rows excluded (`Keep?` not blank) |
| `A6_*.xlsx` | 05 | Final TMS import file, mapped to the C template's columns |
| `A6_META_*.json` | 05 | Counts/warnings/paths - read by 06 to build the email |

## Setup

1. **Python 3.10+** and the dependencies in `requirements.txt`:
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt

   python -m venv venv
   venv\Scripts\Activate.ps1
   python -m pip install -r requirements.txt
   ```
2. **`config.conf`** - copy/edit in place. Sections:
   - `[sftp]` - host/port/user/password (or `private_key_path`), remote
     directory, and the filename prefix/suffix used to pick out extract
     files.
   - `[local]` - `download_root` / `work_dir` / `output_dir` (all relative
     paths are resolved from this pipeline's own directory, not wherever
     it's invoked from - see below).
   - `[extraction_rules]` - `customer_code` (OCBA's FCT_CLI_CODE_FAC value).
   - `[templates]` - paths to the four supporting template files under
     `./templates/` (B admin template, TMS article list, location list,
     project-code-prefix allowlist). **Double-check the prefix allowlist
     filename** - see [Known follow-ups](#known-follow-ups).
   - `[notifications]` - SMTP settings for 06's completion email. Leave
     `smtp_host` blank during setup/testing; the email content is printed
     to the console instead of sent.
3. **Keep `config.conf` out of version control** - it holds SFTP and SMTP
   credentials. Add it to `.gitignore`.
4. Populate `./templates/` with the four files `config.conf` points to
   (synced by hand from SharePoint/TMS exports - see the comments above
   each path in `config.conf`).

## Running manually

Each script can still be run and inspected individually (useful when
setting things up, or debugging one stage):

```bash
python 01_download_sftp.py
python 02_read_files.py
python 03_clean_transform.py
python 04_validate.py
python 05_export.py
python 06_notify.py
```

Every script auto-discovers its input from the previous stage's dated
folder, so this order "just works" as long as each step succeeded. See
each script's own module docstring for what it checks and produces.

## Running via the master script

`run_pipeline.py` runs all six in order and is what a scheduled/production
job should call instead of chaining the six scripts by hand:

```bash
python run_pipeline.py                 # run the full pipeline
python run_pipeline.py --only 03       # run just one numbered step (debugging)
python run_pipeline.py --dry-run       # print the plan, run nothing
python run_pipeline.py --no-lock       # skip the lock file (debugging only)
```

It adds, on top of the six scripts:

- **A file lock** (`.pipeline.lock`) so two overlapping scheduled runs can
  never process the same batch concurrently. If a run finds the lock
  already held, it exits immediately and quietly (exit code 0) rather than
  waiting or double-processing.
- **Per-run logging** to both the console and a timestamped file under
  `./logs/` (`pipeline_run_<timestamp>.log`) - every step's full
  stdout/stderr is captured either way.
- **Retries for step 01 only** (SFTP is the most common source of
  transient failures) - 3 attempts, 30s apart - but only on a real error,
  never on the normal "nothing new this week yet" outcome.
- **A 30-minute timeout per step**, so a hung connection can't wedge a
  scheduled job (and its lock) forever.
- **Clean stop-on-failure**: since every step depends entirely on the
  previous step's output file, the master script never tries to "skip and
  continue" - it stops at the first real failure.

### Exit codes (what a cron/monitoring wrapper should check)

| Code | Meaning |
|---|---|
| `0` | Nothing to alert on. Either the full pipeline (including the email) succeeded, OR there was genuinely nothing new to process yet (this week's extract hasn't landed on the SFTP server), OR another run was already in progress. All three are logged distinctly - check the log to tell them apart - but none need a human to act. |
| `1` | A real failure in steps 01-05 (download/processing). **The batch was not fully processed** - investigate before the next scheduled run. |
| `2` | Steps 01-05 all succeeded - **A6 was produced correctly and is safe to import into TMS** - but the completion email (step 06) failed. Lower urgency than `1`; fix `[notifications]` in `config.conf` and re-run `python 06_notify.py` by hand. |

## Production / cron deployment

1. **Always invoke `run_pipeline.py` with an absolute path**, and run it
   from a venv with the dependencies installed - cron's environment is
   minimal (no `PATH` beyond `/usr/bin:/bin`, no shell profile, no
   assumptions about the working directory). `run_pipeline.py` already
   `chdir`s into its own directory internally, so relative paths in
   `config.conf` resolve correctly regardless of cron's own cwd - but the
   Python interpreter itself must still be specified with a full path.

2. **Timezone matters.** `compute_week_tag()` uses the server's local
   clock. Make sure the machine running cron is in the timezone you expect
   (or explicitly set `TZ` in the crontab) - otherwise a run close to a
   week boundary could compute the wrong `week_tag`.

3. **Example crontab entry** (runs every weekday at 07:00 server time):
   ```cron
   TZ=Europe/Madrid
   MAILTO=logistics-tech@example.org
   0 7 * * 1-5 /opt/msf-pipeline/venv/bin/python /opt/msf-pipeline/run_pipeline.py >> /opt/msf-pipeline/logs/cron.out 2>&1
   ```
   Since `run_pipeline.py` exits `0` for both "fully succeeded" and
   "nothing new yet", a plain cron `MAILTO` (which mails on any output,
   depending on your cron daemon's settings) may still be noisier than
   you want. Two common refinements:
   - Redirect stdout to the log file (as above) and only alert based on
     **exit code**, using a tiny wrapper:
     ```bash
     #!/bin/bash
     /opt/msf-pipeline/venv/bin/python /opt/msf-pipeline/run_pipeline.py
     code=$?
     if [ $code -ne 0 ]; then
         mail -s "MSF pipeline exit $code" logistics-tech@example.org < /opt/msf-pipeline/logs/cron.out
     fi
     exit $code
     ```
   - Or use a tool like `moreutils`' `chronic` / `cronic` to only mail
     when a command fails or produces unexpected output.

4. **Don't run multiple cron entries in parallel for the same pipeline.**
   The master script's own lock file protects against genuine overlap, but
   there's no reason to schedule it more than once a day (or a few times a
   week around when the extract typically lands) - each no-op run still
   costs an SFTP round trip.

5. **Log rotation.** `./logs/pipeline_run_*.log` accumulates one file per
   run and is never cleaned up automatically. Add a `logrotate` config, or
   a simple periodic `find ./logs -mtime +90 -delete` job:
   ```
   /opt/msf-pipeline/logs/*.log {
       weekly
       rotate 12
       compress
       missingok
       notifempty
   }
   ```

6. **Disk space for `downloads/`, `work/`, `output/`.** These grow by one
   dated subfolder per week indefinitely - nothing in this pipeline prunes
   old ones. Set up a periodic archive/cleanup job once you have a
   retention policy (e.g. keep 6 months locally, archive older folders
   elsewhere).

7. **Credentials.** `config.conf` holds the SFTP password (or private key
   path) and SMTP credentials in plain text. Restrict its file permissions
   (`chmod 600 config.conf`) and make sure it's excluded from backups that
   might land somewhere less secure, and from version control.

8. **First run in production**: run `python run_pipeline.py --dry-run`
   first to confirm the step list, then a real run by hand (not via cron)
   to confirm SFTP connectivity, templates, and SMTP all work end-to-end
   before trusting it to a schedule.

## Troubleshooting

- **"Refusing to download a stale file" / exit code 2 from step 01**: this
  week's extract genuinely hasn't been uploaded to the SFTP server yet.
  Not an error - just re-run later (or wait for the next scheduled run).
- **`PermissionError` opening/saving an `.xlsx`**: the file is open in
  Excel (which locks it on Windows) - close it and re-run that step.
- **04/05 complain about a missing template file**: re-check the paths
  under `[templates]` in `config.conf` actually point at files that exist
  under `./templates/` - these are synced by hand from SharePoint/TMS and
  aren't part of this repo.
- **03's Keep? check flags everything as "NOT DEPLOYED"**: the project
  code prefix allowlist is empty or pointing at the wrong file - see
  [Known follow-ups](#known-follow-ups) below.
- **06 prints the email instead of sending it**: `[notifications]
  smtp_host` is blank in `config.conf` - this is the safe default until
  you fill in real SMTP settings.

## Known follow-ups

- **`project_code_prefixes_path` filename mismatch**: `config.conf`
  currently points at `./templates/project_code_prefixes.txt`, but earlier
  versions of `03_clean_transform.py` had a hardcoded, differently-named
  default (`Mission_code_prefixes.txt`) that silently ignored the config
  value entirely. This has been fixed - the script now reads the path from
  `config.conf` (falling back to the old hardcoded name only if the config
  key is absent) - **but you should confirm which filename the real
  allowlist file on disk actually uses**, and make sure `config.conf` and
  the file on disk agree, since this file previously might have been
  silently ignored.
- **No automated retention/cleanup** for `downloads/`, `work/`, `output/`,
  or `./logs/` - see the cron section above.
- **06_notify.py's failure doesn't retry** - if SMTP is down when it runs,
  re-run `python 06_notify.py` by hand once it's fixed (it will find the
  same batch's meta JSON and re-send).
