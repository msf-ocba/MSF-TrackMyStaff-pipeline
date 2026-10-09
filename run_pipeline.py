#!/usr/bin/env python3
"""
run_pipeline.py
----------------
Master script for the MSF Logistique -> TMS asset-import pipeline.
Runs 00_export_odoo.py through 06_notify.py, in order, as a single
automated job - this is the script a cron job (or systemd timer)
should actually call. Running the seven numbered scripts by hand is
still supported for debugging (see each script's own docstring), but
in production this file is the entry point.

WHAT IT DOES, IN ORDER
  1. Chdir into this file's own directory, so relative paths in
     config.conf (download_root, work_dir, output_dir, templates/...)
     always resolve the same way regardless of where/how cron invokes
     this script (cron's working directory is usually NOT this folder).
  2. Takes an exclusive file lock (see LOCKING below) so two overlapping
     cron runs can never process the same batch at once.
  3. Sets up logging to both the console (for `cron`'s own output
     capture / MAILTO) and a timestamped file under ./logs/, so every
     run leaves a permanent record.
  4. Runs 00_export_odoo.py, which refreshes the reference files under
     templates/ (TMS article list + location list) from Odoo, so this
     batch is validated against current data. If this step still fails
     after its retries, the pipeline does NOT stop: it logs a warning,
     emails the [notifications] recipients the error (and how old the
     existing templates are), and carries on with step 01 using the
     existing (stale) templates. The overall exit code is not affected
     by a step 00 failure - the email is the alert. (With --only 00 a
     failure returns exit code 1 and sends no email.)
  4b. Runs 01_download_sftp.py. This step gets a few retries with a
     short backoff, since SFTP/network hiccups are the single most
     common transient failure in this pipeline (see RETRY POLICY).
     If step 01 does not produce a file - exit code 1 (real error, after
     its retries) OR exit code 2 (this week's extract not on the SFTP
     server yet) - an alert email is sent to the [notifications]
     recipients (short, non-technical message) before the pipeline stops.
     No email is sent when running with --only.
  5. If 01 downloaded a new file, runs 02 -> 03 -> 04 -> 05 in strict
     order, STOPPING IMMEDIATELY if any of them fail (see EXIT CODES) -
     each step's output depends entirely on the previous step's output
     file, so there is no reasonable way to "skip and continue".
     EXCEPTION: if 04_validate.py finds that every row in this batch
     already has a non-blank 'Keep?' (nothing new to import), 05 is
     skipped (there is nothing to export) but 06_notify.py is still
     run directly - see step 6 - so a "nothing to import" completion
     email still goes out. Users depend on that email arriving; they
     shouldn't have to check server logs to learn the batch ran with
     nothing new.
  6. If 01-05 all succeed, runs 06_notify.py to send the completion
     email (this also happens, out of the normal 01->06 order, right
     after 04 in the "nothing to import" case described in step 5). A
     failure here is logged and reported distinctly (see EXIT CODES)
     but does NOT roll back or repeat steps 1-5, since A6 (or, in the
     nothing-to-import case, A5B) has already been produced.
  7. Releases the lock and prints a final summary (what ran, exit
     codes, durations) before exiting with the appropriate code.

EXIT CODES (this is the contract a cron wrapper / monitoring tool
should check)
  0  - Normal, nothing to alert on. Covers FOUR distinct situations,
       all logged clearly so you can tell them apart by reading the
       log, but none of which need a human to act:
         a) full pipeline ran successfully end-to-end, including the
            notification email
         b) 01_download_sftp.py determined there is nothing new to
            process yet (this week's extract hasn't landed on the
            SFTP server) - normal/expected if cron runs more often
            than the weekly extract actually arrives
         c) 04_validate.py determined that every row in this batch
            already has a non-blank 'Keep?' (e.g. all "NOT DEPLOYED" /
            "NOT ASSET") - there is nothing new to import into TMS
            this run. A5B is still produced for the audit trail,
            05_export.py is skipped since there is no A5/A6 to
            export, but 06_notify.py IS still run (out of the normal
            step order) to send a distinct "nothing to import"
            completion email - see step 5/6 above. If that email
            itself fails to send, this is reported as exit code 2,
            same as any other 06_notify.py failure.
         d) another instance of this pipeline was already running
            (lock held) - this run exited immediately without doing
            anything, to avoid double-processing the same batch
  1  - A REAL FAILURE in steps 01-05 (data download and
       processing). The
       batch was NOT fully processed - A6 may not exist, or may be
       incomplete/stale. Needs investigation before the next run.
  2  - Steps 01-05 all succeeded (A6 was produced correctly, and IS
       safe to import into TMS), but 06_notify.py (the completion
       email) failed. Needs investigation, but is lower urgency than
       exit code 1 since the actual data pipeline completed.

  A cron wrapper that only wants to be paged on exit code 1 (and treat
  2 as a lower-priority ticket, and 0 as silence) can branch on these
  directly - see the README's "Cron / production" section for a
  ready-to-use crontab entry and wrapper script.

LOCKING
  Uses a plain lock file (./.pipeline.lock) plus an OS-level advisory
  lock (fcntl.flock on Linux/Mac). On platforms without fcntl (e.g.
  Windows), falls back to a PID-file check (is the PID recorded in the
  lock file still alive?) - slightly weaker (small race window) but
  good enough for a once-a-week batch job. Either way, a run that finds
  the lock already held exits immediately with code 0 (see EXIT CODES)
  rather than waiting or erroring.

RETRY POLICY
  Only 00_export_odoo.py and 01_download_sftp.py are retried
  automatically (both talk to a remote server), and only when they
  fail with a real error (exit code 1 / a crash / a timeout) - NOT
  when it exits with code 2 ("not ready yet", e.g. the weekly extract
  genuinely isn't on the server yet), since retrying that immediately
  will just get the same answer. Default: 3 attempts, 30s apart. Steps
  02-06 are never retried automatically - they operate on local files
  and a repeat failure almost always means a real, non-transient
  problem (bad data, missing template, disk full, etc.) that a retry
  won't fix; see the log / re-run the specific step by hand once fixed.
  (04_validate.py's exit code 3, "nothing to import", is likewise
  never retried - it's a normal outcome, not a transient error.)

TIMEOUTS
  Every step gets a generous but finite timeout (STEP_TIMEOUT_SECONDS,
  default 30 minutes) so a hung SFTP connection or a stuck subprocess
  can't wedge a cron job forever and block every future scheduled run
  via the lock file.

USAGE
    python run_pipeline.py                 # run the full pipeline
    python run_pipeline.py --only 00       # refresh templates from Odoo only
    python run_pipeline.py --only 03       # run just one numbered step
    python run_pipeline.py --dry-run       # print the plan, run nothing
    python run_pipeline.py --no-lock       # skip locking (debugging only)
"""
import argparse
import datetime
import logging
import os
import smtplib
import subprocess
import sys
import time
from email.mime.text import MIMEText

from common_config import load_config

try:
    import fcntl
    HAVE_FCNTL = True
except ImportError:
    HAVE_FCNTL = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(BASE_DIR, "logs")
LOCK_PATH = os.path.join(BASE_DIR, ".pipeline.lock")

STEP_TIMEOUT_SECONDS = 30 * 60  # 30 minutes per step

# Exit codes this master script can return - see module docstring.
EXIT_OK = 0
EXIT_PIPELINE_FAILED = 1
EXIT_NOTIFY_FAILED = 2

# Exit codes a numbered script can return that this master script
# specifically understands (see the relevant script's own contract).
STEP_EXIT_OK = 0
STEP_EXIT_NOT_READY = 2       # only meaningful for 01_download_sftp.py
STEP_EXIT_NOTHING_TO_IMPORT = 3  # only meaningful for 04_validate.py

# Step exit codes that mean "stop the pipeline here, but this is a
# normal/expected outcome, not a failure" - see EXIT CODES above.
STEP_EXIT_STOP_CODES = (STEP_EXIT_NOT_READY, STEP_EXIT_NOTHING_TO_IMPORT)

STEPS = [
    # 00 refreshes templates/ from Odoo. A failure here is NON-FATAL in a
    # full run: main() emails an alert and continues with stale templates.
    {"id": "00", "script": "00_export_odoo.py", "max_attempts": 3, "retry_delay_s": 30},
    {"id": "01", "script": "01_download_sftp.py", "max_attempts": 3, "retry_delay_s": 30},
    {"id": "02", "script": "02_read_files.py", "max_attempts": 1, "retry_delay_s": 0},
    {"id": "03", "script": "03_clean_transform.py", "max_attempts": 1, "retry_delay_s": 0},
    {"id": "04", "script": "04_validate.py", "max_attempts": 1, "retry_delay_s": 0},
    {"id": "05", "script": "05_export.py", "max_attempts": 1, "retry_delay_s": 0},
    {"id": "06", "script": "06_notify.py", "max_attempts": 1, "retry_delay_s": 0},
]

logger = logging.getLogger("run_pipeline")


def setup_logging():
    os.makedirs(LOG_DIR, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(LOG_DIR, f"pipeline_run_{timestamp}.log")

    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(fmt)
    logger.addHandler(console_handler)

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(fmt)
    logger.addHandler(file_handler)

    return log_path


class PipelineLock:
    """
    Simple cross-run lock so two cron-triggered pipeline runs can never
    process the same batch concurrently. Prefers an OS-level advisory
    lock (fcntl.flock) when available; falls back to a PID-liveness
    check on platforms without fcntl (e.g. Windows). Either way, if the
    lock can't be acquired, `acquire()` returns False rather than
    blocking - the caller is expected to exit cleanly (exit code 0,
    see EXIT CODES in the module docstring) rather than queue up.
    """

    def __init__(self, path):
        self.path = path
        self._fh = None

    def acquire(self):
        if HAVE_FCNTL:
            return self._acquire_fcntl()
        return self._acquire_pidfile()

    def _acquire_fcntl(self):
        self._fh = open(self.path, "a+")
        try:
            fcntl.flock(self._fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            self._fh.close()
            self._fh = None
            return False
        self._fh.seek(0)
        self._fh.truncate()
        self._fh.write(str(os.getpid()))
        self._fh.flush()
        return True

    def _acquire_pidfile(self):
        if os.path.exists(self.path):
            try:
                with open(self.path) as f:
                    old_pid = int(f.read().strip() or -1)
            except (ValueError, OSError):
                old_pid = -1
            if old_pid > 0 and _pid_is_alive(old_pid):
                return False
        with open(self.path, "w") as f:
            f.write(str(os.getpid()))
        return True

    def release(self):
        if HAVE_FCNTL and self._fh is not None:
            try:
                fcntl.flock(self._fh.fileno(), fcntl.LOCK_UN)
            except OSError:
                pass
            self._fh.close()
            self._fh = None
        else:
            try:
                os.remove(self.path)
            except FileNotFoundError:
                pass


def _pid_is_alive(pid):
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True  # exists, just owned by someone else
    except OSError:
        return False
    return True


def run_step(step):
    """
    Runs one numbered script as a subprocess, with the retry policy
    from STEPS. Returns (returncode, duration_seconds, output) where
    output is the combined stdout/stderr of the final attempt (used for
    the step 00 failure email). Always logs
    full stdout/stderr - at INFO level for a clean exit, at ERROR level
    for anything else - so the log file has everything needed to debug
    a failure without re-running by hand.
    """
    script = step["script"]
    max_attempts = step["max_attempts"]
    retry_delay_s = step["retry_delay_s"]

    attempt = 0
    while True:
        attempt += 1
        logger.info(f"--- Step {step['id']} ({script}): attempt {attempt}/{max_attempts} ---")
        start = time.monotonic()
        try:
            result = subprocess.run(
                [sys.executable, script],
                cwd=BASE_DIR,
                capture_output=True,
                text=True,
                timeout=STEP_TIMEOUT_SECONDS,
            )
            returncode = result.returncode
            stdout, stderr = result.stdout, result.stderr
        except subprocess.TimeoutExpired as e:
            duration = time.monotonic() - start
            logger.error(
                f"Step {step['id']} ({script}) TIMED OUT after {STEP_TIMEOUT_SECONDS}s "
                f"(attempt {attempt}/{max_attempts})."
            )
            if e.stdout:
                logger.error("--- stdout up to timeout ---\n" + e.stdout.decode(errors="replace"))
            if e.stderr:
                logger.error("--- stderr up to timeout ---\n" + e.stderr.decode(errors="replace"))
            returncode = -1  # sentinel: treated as a real failure, eligible for retry
            stdout, stderr = "", f"Step timed out after {STEP_TIMEOUT_SECONDS}s."
            duration_recorded = duration
        else:
            duration = time.monotonic() - start
            duration_recorded = duration

        if returncode == STEP_EXIT_OK:
            logger.info(f"Step {step['id']} ({script}) succeeded in {duration_recorded:.1f}s.")
            if stdout.strip():
                logger.info(f"--- {script} stdout ---\n{stdout.rstrip()}")
            return returncode, duration_recorded, stdout

        if returncode in STEP_EXIT_STOP_CODES:
            # A normal, expected "nothing to do" outcome - never
            # retried. Meaningful for 01_download_sftp.py (code 2,
            # "not ready yet") and 04_validate.py (code 3, "nothing to
            # import this run").
            if returncode == STEP_EXIT_NOT_READY:
                reason = "NOT READY (nothing new to process this run)"
            else:
                reason = "NOTHING TO IMPORT (no rows required import this run)"
            logger.info(f"Step {step['id']} ({script}) reports {reason} - exit code {returncode}.")
            if stdout.strip():
                logger.info(f"--- {script} stdout ---\n{stdout.rstrip()}")
            return returncode, duration_recorded, stdout

        # Anything else is a real failure - log everything, retry if
        # attempts remain, otherwise return the failing code.
        logger.error(f"Step {step['id']} ({script}) FAILED (exit code {returncode}).")
        if stdout.strip():
            logger.error(f"--- {script} stdout ---\n{stdout.rstrip()}")
        if stderr.strip():
            logger.error(f"--- {script} stderr ---\n{stderr.rstrip()}")

        if attempt < max_attempts:
            logger.warning(
                f"Retrying step {step['id']} ({script}) in {retry_delay_s}s "
                f"(attempt {attempt + 1}/{max_attempts})..."
            )
            time.sleep(retry_delay_s)
            continue

        return returncode, duration_recorded, (stdout + "\n" + stderr).strip()


def _templates_last_updated(cfg):
    """
    Returns (oldest_modified_datetime or None, any_missing) across the two
    Odoo-generated template files, for plain-language alert emails.
    """
    modified_times, any_missing = [], False
    for key, default in (
        ("tms_article_list_path", "./templates/TMS_UniDataArticles.xlsx"),
        ("location_list_path", "./templates/location.xlsx"),
    ):
        full = os.path.join(BASE_DIR, cfg.get("templates", key, fallback=default))
        if os.path.exists(full):
            modified_times.append(datetime.datetime.fromtimestamp(os.path.getmtime(full)))
        else:
            any_missing = True
    return (min(modified_times) if modified_times else None), any_missing


def send_alert_email(subject, body):
    """
    Emails the [notifications] recipients. NEVER raises: a broken alert
    must not take the pipeline down. If SMTP isn't configured (or the
    send fails) the would-be email is written to the log instead.
    """
    try:
        cfg = load_config(os.path.join(BASE_DIR, "config.conf"))
        smtp_host = cfg.get("notifications", "smtp_host", fallback="").strip()
        recipients = [r.strip() for r in cfg.get("notifications", "to_addresses", fallback="").split(",") if r.strip()]
        if not smtp_host or not recipients:
            logger.warning(
                "Alert email NOT sent: [notifications] smtp_host / to_addresses "
                f"not configured. Would have sent:\nSubject: {subject}\n{body}"
            )
            return

        smtp_port = cfg.getint("notifications", "smtp_port", fallback=587)
        smtp_user = cfg.get("notifications", "smtp_user", fallback="")
        smtp_password = cfg.get("notifications", "smtp_password", fallback="")
        use_tls = cfg.getboolean("notifications", "use_tls", fallback=True)
        sender = cfg.get("notifications", "from_address", fallback=smtp_user) or smtp_user

        msg = MIMEText(body)
        msg["Subject"] = subject
        msg["From"] = sender
        msg["To"] = ", ".join(recipients)

        with smtplib.SMTP(smtp_host, smtp_port, timeout=60) as server:
            if use_tls:
                server.starttls()
            if smtp_user:
                server.login(smtp_user, smtp_password)
            server.sendmail(sender, recipients, msg.as_string())
        logger.info(f"Alert email sent to: {', '.join(recipients)} (Subject: {subject})")
    except Exception as exc:  # noqa: BLE001 - alert must never crash the pipeline
        logger.error(f"Could not send the alert email '{subject}' ({exc}). See the log for the original error.")


def send_odoo_failure_email(log_path):
    """Step 00 failed: pipeline continues with the previous lists; tell the functional team."""
    try:
        cfg = load_config(os.path.join(BASE_DIR, "config.conf"))
        last_updated, any_missing = _templates_last_updated(cfg)
    except Exception:  # noqa: BLE001
        last_updated, any_missing = None, False

    if any_missing or last_updated is None:
        impact = ("No previous article/location lists were found, so this week's "
                  "import will most likely fail as well.")
    else:
        impact = (f"The import is continuing with the previous lists (last updated "
                  f"{last_updated:%d %b %Y}). Articles or locations added or changed in "
                  "Odoo since then may not be recognised, so please review this "
                  "week's results with extra care.")

    body = "\n".join([
        "Hello,",
        "",
        "The article and location lists could not be refreshed from Odoo.",
        impact,
        "",
        "Action: please ask the technician to check the connection to Odoo.",
        "",
        f"(Technician: details are in the pipeline log, {log_path})",
    ])
    send_alert_email("Action needed: TMS article/location lists could not be updated from Odoo", body)


def send_download_alert_email(returncode, log_path):
    """
    Step 01 did not produce a file this run. Exit code 2 = the weekly
    extract isn't on the SFTP server yet; anything else = a real error.
    In both cases the pipeline stops here (02-06 have nothing to process).
    """
    if returncode == STEP_EXIT_NOT_READY:
        subject = "This week's extract has not arrived yet"
        lines = [
            "Hello,",
            "",
            "This week's extract from MSF Logistique is not available yet, so "
            "nothing was imported.",
            "",
            "Action: none for now - the import will run automatically once the "
            "file arrives. If it is still missing later in the week, please check "
            "with MSF Logistique.",
        ]
    else:
        subject = "Action needed: this week's extract could not be downloaded"
        lines = [
            "Hello,",
            "",
            "The import could not download this week's extract, so this week's "
            "batch was NOT processed.",
            "",
            "Action: please ask the technician to check the connection to the "
            "file server. The import can be re-run once it is fixed.",
        ]
    lines += ["", f"(Technician: details are in the pipeline log, {log_path})"]
    send_alert_email(subject, "\n".join(lines))


def main():
    parser = argparse.ArgumentParser(description="Run the full MSF Logistique -> TMS import pipeline.")
    parser.add_argument(
        "--only", metavar="STEP_ID", choices=[s["id"] for s in STEPS],
        help="Run only one numbered step (e.g. --only 03), for debugging. Skips locking by default.",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Print the plan (which steps would run, in what order) and exit without running anything.",
    )
    parser.add_argument(
        "--no-lock", action="store_true",
        help="Skip the lock file entirely. For debugging only - never use this in the real cron job.",
    )
    args = parser.parse_args()

    os.chdir(BASE_DIR)

    steps_to_run = STEPS if not args.only else [s for s in STEPS if s["id"] == args.only]

    if args.dry_run:
        print("Dry run - the following steps would execute in order:")
        for s in steps_to_run:
            print(f"  {s['id']}: {s['script']} (max_attempts={s['max_attempts']}, "
                  f"retry_delay={s['retry_delay_s']}s)")
        return EXIT_OK

    log_path = setup_logging()
    logger.info(f"=== Pipeline run starting. Log file: {log_path} ===")

    use_lock = not args.no_lock and args.only is None
    lock = PipelineLock(LOCK_PATH)
    if use_lock:
        if not lock.acquire():
            logger.warning(
                "Another pipeline run appears to already be in progress "
                f"(lock file: {LOCK_PATH}). Exiting immediately without doing "
                "anything, to avoid double-processing the same batch."
            )
            return EXIT_OK
    else:
        logger.info("Locking skipped (--only or --no-lock given).")

    overall_start = time.monotonic()
    exit_code = EXIT_OK
    try:
        for i, step in enumerate(steps_to_run):
            returncode, _, step_output = run_step(step)

            if step["id"] == "00" and returncode != STEP_EXIT_OK:
                if args.only == "00":
                    # Debugging 00 by hand: report the failure plainly, no email.
                    logger.error("00_export_odoo.py failed (see output above).")
                    exit_code = EXIT_PIPELINE_FAILED
                    break
                # Full run: stale templates are acceptable. Alert a human
                # by email, then carry on with step 01 and the rest.
                logger.warning(
                    "00_export_odoo.py failed - templates were NOT refreshed from Odoo. "
                    "CONTINUING the pipeline with the existing (stale) templates and "
                    "sending an alert email."
                )
                send_odoo_failure_email(log_path)
                continue

            if step["id"] == "01" and returncode != STEP_EXIT_OK and args.only is None:
                # Exit 1 (real error, after retries) or exit 2 (not ready):
                # either way no file was downloaded, so tell a human by
                # email. The normal stop/failure handling below still runs.
                send_download_alert_email(returncode, log_path)

            if returncode == STEP_EXIT_OK:
                continue

            if returncode == STEP_EXIT_NOT_READY:
                # Only 01_download_sftp.py should ever produce this.
                # Stop the pipeline here - there is nothing for 02-06
                # to process yet - and treat the whole run as a
                # normal, silent no-op (exit code 0).
                logger.info("Stopping pipeline: nothing new to process this run.")
                exit_code = EXIT_OK
                break

            if returncode == STEP_EXIT_NOTHING_TO_IMPORT:
                # Only 04_validate.py should ever produce this: every
                # row in the batch already had a non-blank 'Keep?', so
                # there is nothing to export. 05_export.py is skipped,
                # but users depend on getting a completion email either
                # way - not on checking server logs - so 06_notify.py
                # is still run directly here (out of the normal step
                # order) whenever it's part of the planned run. If this
                # was invoked with "--only 04" (debugging a single
                # step), that plan doesn't include 06, so it's left
                # alone rather than run behind the user's back.
                if any(s["id"] == "06" for s in steps_to_run):
                    logger.info(
                        "04_validate.py found no rows to import this run (every "
                        "row already had a non-blank 'Keep?'). A5B was still "
                        "produced for the audit trail. Skipping 05_export.py "
                        "(nothing to export), but running 06_notify.py now so a "
                        "'nothing to import' completion email still goes out."
                    )
                    notify_step = next(s for s in STEPS if s["id"] == "06")
                    notify_returncode, _, _ = run_step(notify_step)
                    if notify_returncode == STEP_EXIT_OK:
                        exit_code = EXIT_OK
                    else:
                        logger.error(
                            "06_notify.py failed after a 'nothing to import' batch - "
                            "the notification email was NOT sent. Check the log "
                            "above, fix [notifications] in config.conf if needed, "
                            "and re-run 'python 06_notify.py' by hand once resolved."
                        )
                        exit_code = EXIT_NOTIFY_FAILED
                else:
                    logger.info(
                        "04_validate.py found no rows to import this run (every "
                        "row already had a non-blank 'Keep?'). A5B was still "
                        "produced for the audit trail. Stopping here since this "
                        "run's plan (--only 04) doesn't include 06_notify.py."
                    )
                    exit_code = EXIT_OK
                break

            # Any other non-zero code is a real failure.
            if step["id"] == "06":
                # Data pipeline (01-05) already succeeded; only the
                # notification step failed. Surface this distinctly
                # (exit code 2) rather than as a full pipeline failure -
                # A6 was still produced correctly and is importable.
                logger.error(
                    "06_notify.py failed, but steps 01-05 completed successfully - "
                    "A6 was produced and IS safe to import into TMS. The "
                    "completion email was NOT sent; check the log above for why, "
                    "fix [notifications] in config.conf if needed, and re-run "
                    "'python 06_notify.py' by hand once resolved."
                )
                exit_code = EXIT_NOTIFY_FAILED
            else:
                logger.error(
                    f"Stopping pipeline: step {step['id']} ({step['script']}) failed "
                    f"and every later step depends on its output. See the error "
                    f"details above. Fix the underlying issue, then either re-run "
                    f"the whole pipeline or that step directly with "
                    f"'python {step['script']}'."
                )
                exit_code = EXIT_PIPELINE_FAILED
            break
    finally:
        if use_lock:
            lock.release()

    total_duration = time.monotonic() - overall_start
    logger.info(f"=== Pipeline run finished in {total_duration:.1f}s. Exit code: {exit_code} ===")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
