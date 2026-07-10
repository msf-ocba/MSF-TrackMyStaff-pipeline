#!/usr/bin/env python3
"""
run_pipeline.py
----------------
Master script for the MSF Logistique -> TMS asset-import pipeline.
Runs 01_download_sftp.py through 06_notify.py, in order, as a single
automated job - this is the script a cron job (or systemd timer)
should actually call. Running the six numbered scripts by hand is
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
  4. Runs 01_download_sftp.py. This step gets a few retries with a
     short backoff, since SFTP/network hiccups are the single most
     common transient failure in this pipeline (see RETRY POLICY).
  5. If 01 downloaded a new file, runs 02 -> 03 -> 04 -> 05 in strict
     order, STOPPING IMMEDIATELY if any of them fail (see EXIT CODES) -
     each step's output depends entirely on the previous step's output
     file, so there is no reasonable way to "skip and continue".
  6. If 01-05 all succeed, runs 06_notify.py to send the completion
     email. A failure here is logged and reported distinctly (see EXIT
     CODES) but does NOT roll back or repeat steps 1-5, since A6 has
     already been produced and is importable into TMS regardless of
     whether the email went out.
  7. Releases the lock and prints a final summary (what ran, exit
     codes, durations) before exiting with the appropriate code.

EXIT CODES (this is the contract a cron wrapper / monitoring tool
should check)
  0  - Normal, nothing to alert on. Covers THREE distinct situations,
       all logged clearly so you can tell them apart by reading the
       log, but none of which need a human to act:
         a) full pipeline ran successfully end-to-end, including the
            notification email
         b) 01_download_sftp.py determined there is nothing new to
            process yet (this week's extract hasn't landed on the
            SFTP server) - normal/expected if cron runs more often
            than the weekly extract actually arrives
         c) another instance of this pipeline was already running
            (lock held) - this run exited immediately without doing
            anything, to avoid double-processing the same batch
  1  - A REAL FAILURE in steps 01-05 (data download/processing). The
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
  Only 01_download_sftp.py is retried automatically, and only when it
  fails with a real error (exit code 1 / a crash / a timeout) - NOT
  when it exits with code 2 ("not ready yet", e.g. the weekly extract
  genuinely isn't on the server yet), since retrying that immediately
  will just get the same answer. Default: 3 attempts, 30s apart. Steps
  02-06 are never retried automatically - they operate on local files
  and a repeat failure almost always means a real, non-transient
  problem (bad data, missing template, disk full, etc.) that a retry
  won't fix; see the log / re-run the specific step by hand once fixed.

TIMEOUTS
  Every step gets a generous but finite timeout (STEP_TIMEOUT_SECONDS,
  default 30 minutes) so a hung SFTP connection or a stuck subprocess
  can't wedge a cron job forever and block every future scheduled run
  via the lock file.

USAGE
    python run_pipeline.py                 # run the full pipeline
    python run_pipeline.py --only 03       # run just one numbered step
    python run_pipeline.py --dry-run       # print the plan, run nothing
    python run_pipeline.py --no-lock       # skip locking (debugging only)
"""
import argparse
import datetime
import logging
import os
import subprocess
import sys
import time

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
# specifically understands (see 01_download_sftp.py's own contract).
STEP_EXIT_OK = 0
STEP_EXIT_NOT_READY = 2  # only meaningful for 01_download_sftp.py

STEPS = [
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
    from STEPS. Returns (returncode, duration_seconds). Always logs
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
            stdout, stderr = "", ""
            duration_recorded = duration
        else:
            duration = time.monotonic() - start
            duration_recorded = duration

        if returncode == STEP_EXIT_OK:
            logger.info(f"Step {step['id']} ({script}) succeeded in {duration_recorded:.1f}s.")
            if stdout.strip():
                logger.info(f"--- {script} stdout ---\n{stdout.rstrip()}")
            return returncode, duration_recorded

        if returncode == STEP_EXIT_NOT_READY:
            # Only meaningful for 01_download_sftp.py - a normal,
            # expected "nothing new yet" outcome, never retried.
            logger.info(
                f"Step {step['id']} ({script}) reports NOT READY (exit code 2) - "
                f"nothing new to process this run."
            )
            if stdout.strip():
                logger.info(f"--- {script} stdout ---\n{stdout.rstrip()}")
            return returncode, duration_recorded

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

        return returncode, duration_recorded


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
            returncode, _ = run_step(step)

            if returncode == STEP_EXIT_OK:
                continue

            if returncode == STEP_EXIT_NOT_READY:
                # Only 01_download_sftp.py should ever produce this.
                # Stop the pipeline here - there is nothing for 02-06
                # to process yet - and treat the whole run as a normal,
                # silent no-op (exit code 0).
                logger.info("Stopping pipeline: nothing new to process this run.")
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
