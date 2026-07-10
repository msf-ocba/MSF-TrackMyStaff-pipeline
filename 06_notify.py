"""
06_notify.py
--------------
Sends the batch completion email that 05_export.py used to send
itself. This script reads the meta JSON file 05_export.py writes next
to A6 (A6_META_EXTRACTION_TMS_OCBA_YYMMDD.json), locates A4 (the full
pre-filter extract) and A5B (produced earlier by 04_validate.py) if
they exist, and sends ONE completion email - whether the batch was
fully clean or some rows need review.

Run this right after 05_export.py finishes for the same batch.

WHAT THIS SCRIPT DOES NOT DO
  It does not build A6, and it does not re-run any of the Article Code
  / Location / Duplicate checks - all of that already happened in
  05_export.py, and its results (counts, per-row warning details, the
  A6 path) are simply read back out of the meta JSON file. This script
  is purely about finding A4/A5B and sending mail.

A4 - THE FULL PRE-FILTER EXTRACT
  A4_EXTRACTION_TMS_OCBA_YYMMDD.xlsx is created earlier in the
  pipeline (before 04_validate.py splits it into A5/A5B) - it holds
  every row of the batch before the Keep?-based filtering. This script
  looks for that already-existing file in the same dated output folder
  as A6 (given by the meta JSON) and, if found, attaches it to the
  completion email alongside A6 and A5B, so the technician has the
  full original extract, what got imported, and what got excluded all
  in the same email. If no matching A4 is found, the email is still
  sent without it, and a warning is printed/logged noting A4 was
  expected but not found.

A5B - THE "FLAGGED ITEMS" FILE
  A5B_EXTRACTION_TMS_OCBA_YYMMDD.xlsx is created by 04_validate.py - it
  holds the A4 rows that were EXCLUDED before A5/A6 (i.e. every row
  whose 'Keep?' was NOT blank: NOT ASSET / NOT DEPLOYED / both), with
  the same Keep?-based row colouring as A4. This script looks for that
  already-existing file in the same dated output folder as A6 (given
  by the meta JSON) and, if found, attaches it to the completion email
  alongside A6, so both "what got imported" (A6) and "what got
  excluded before that" (A5B) land in the same email. If no matching
  A5B is found (e.g. 04_validate.py hasn't been run for this batch, or
  ran before the output-folder naming changed), the email is still
  sent with just A6 attached, and a warning is printed/logged noting
  A5B was expected but not found.

EMAIL NOTIFICATION
  This script ALWAYS sends one completion email using Python's
  smtplib - whether the batch was fully clean or some rows need
  review:
    - All clean: a short success summary (counts only), with A6 as the
      only required attachment (A4 and A5B are added on top whenever
      they're found - see above).
    - Some rows need review: the same summary, plus the detailed list
      of which rows and why (Article Code Review, missing Location,
      and/or duplicate Article Code), taken from the meta JSON. A4 and
      A5B (found in the output folder, see above) are attached
      whenever they exist, regardless of whether A6 itself has flagged
      rows - they report different, earlier stages of the batch
      (the full pre-filter extract, and the Keep? exclusions from
      04_validate.py).
  The email body always ends with an explicit "Attachments:" section
  listing exactly which file(s) were actually attached (with size), or
  explicitly saying a file was left out and why (missing / too large),
  so the technician never has to guess what they were sent.
  SMTP server/recipient settings live in config.conf under
  [notifications] - fill those in with real values before this can
  actually send. Until then, sending is skipped and the would-be email
  content (including the Attachments section) is printed to the
  console instead, if smtp_host is left blank.

Usage:
    python 06_notify.py [path_to_A6_META_....json]

If the meta-file path is omitted, the script looks in the dated output
subfolders (same "Y26W26"-style convention 05_export.py writes into)
for the most recently modified A6_META_EXTRACTION_TMS_OCBA_*.json.
"""
import os
import sys
import glob
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from common_config import load_config, get_paths

MAX_ATTACHMENT_MB = 20  # most SMTP servers/relays reject attachments above ~20-25MB


def find_latest_meta(output_dir):
    """
    Same dated-subfolder convention as 05_export.py's find_latest_a5():
    looks in the most recently modified dated subfolder(s) under
    output_dir for an A6_META_EXTRACTION_TMS_OCBA_*.json file.
    """
    subfolders = [f for f in glob.glob(os.path.join(output_dir, "*")) if os.path.isdir(f)]
    subfolders.sort(key=os.path.getmtime, reverse=True)
    for folder in subfolders:
        candidates = sorted(glob.glob(os.path.join(folder, "A6_META_EXTRACTION_TMS_OCBA_*.json")))
        if candidates:
            return max(candidates, key=os.path.getmtime)

    direct = sorted(glob.glob(os.path.join(output_dir, "A6_META_EXTRACTION_TMS_OCBA_*.json")))
    if direct:
        chosen = max(direct, key=os.path.getmtime)
        print(f"  WARNING: no dated subfolder under '{output_dir}' contains a meta file. "
              f"Falling back to a flat file found directly in '{output_dir}': {chosen}.")
        return chosen

    raise FileNotFoundError(
        f"No A6_META_EXTRACTION_TMS_OCBA_*.json file found under '{output_dir}' "
        f"(checked dated subfolders and the flat folder). Run 05_export.py first."
    )


def find_existing_a4(output_dir, date_tag):
    """
    A4_EXTRACTION_TMS_OCBA_<date_tag>.xlsx is produced earlier in the
    pipeline (the full extract before 04_validate.py splits it into
    A5/A5B), NOT by 05_export.py or this script. This just looks for
    that already-existing file in the shared output folder so it can
    be attached to the completion email alongside A6 and A5B.

    Returns the path if found, else None (with a console warning -
    the email still goes out without it in that case).
    """
    expected_path = os.path.join(output_dir, f"A4_EXTRACTION_TMS_OCBA_{date_tag}.xlsx")
    if os.path.exists(expected_path):
        return expected_path

    print(f"  WARNING: expected A4 file not found at '{expected_path}'. "
          f"A4 is the full pre-filter extract produced earlier in the pipeline - "
          f"make sure it was written for this batch (date tag '{date_tag}') into "
          f"this same output folder. Continuing without it.")
    return None


def find_existing_a5b(output_dir, date_tag):
    """
    A5B_EXTRACTION_TMS_OCBA_<date_tag>.xlsx is produced by
    04_validate.py earlier in the pipeline (the A4 rows excluded by a
    non-blank Keep?), NOT by 05_export.py or this script. This just
    looks for that already-existing file in the shared output folder
    so it can be attached to the completion email alongside A6.

    Returns the path if found, else None (with a console warning -
    the email still goes out with just A6 attached in that case).
    """
    expected_path = os.path.join(output_dir, f"A5B_EXTRACTION_TMS_OCBA_{date_tag}.xlsx")
    if os.path.exists(expected_path):
        return expected_path

    print(f"  WARNING: expected A5B file not found at '{expected_path}'. "
          f"A5B is produced by 04_validate.py - make sure it was run for this "
          f"batch (date tag '{date_tag}') and wrote into this same output folder. "
          f"Continuing without it - the email will only have A6 attached.")
    return None


def build_completion_email(week_tag, date_tag, total_rows, n_clean, n_review, n_missing_location,
                            n_duplicate, review_rows_summary, missing_location_summary,
                            duplicate_summary, a6_path, a5b_path, a4_path):
    """
    Builds ONE email that always reports the batch's outcome:
      - If everything is clean: a short success summary.
      - If anything needs review: the same summary, plus the detailed
        list of which rows and why (review / missing location /
        duplicate Article Code), so the technician doesn't have to
        open the console output or A6 itself to find them.
    A4's and A5B's paths (or their absence) are always reported,
    regardless of whether A6 itself needs review, since they reflect
    earlier, separate stages of the batch (the full pre-filter extract,
    and 04_validate.py's Keep? exclusions).
    (The actual Attachments section is appended later, in send_email,
    once it's known what was actually attached.)

    The subject leads with week_tag (e.g. "Y26W26") to match the
    output folder naming used throughout the pipeline; the headline
    and body still name the batch's own date_tag (e.g. "260629") for
    precise identification, since more than one batch can land in the
    same week folder.
    """
    needs_review = bool(n_review or n_missing_location or n_duplicate)

    if needs_review:
        subject = f"[MSF TMS Import] {week_tag} - Batch {date_tag} completed - review needed"
        headline = f"Batch {date_tag} (folder {week_tag}) completed: {total_rows} item(s) processed into A6."
    else:
        subject = f"[MSF TMS Import] {week_tag} - Batch {date_tag} completed successfully"
        headline = (
            f"Batch {date_tag} (folder {week_tag}) completed successfully: {total_rows} item(s) processed "
            f"into A6, all clean. No review needed."
        )

    body_lines = [
        headline,
        "",
        f"Total items in A6: {total_rows}",
        f"Clean (no issues): {n_clean}",
        f"Article Code Check = Review: {n_review}",
        f"Missing Location match: {n_missing_location}",
        f"Duplicate Article Code: {n_duplicate}",
        "",
        f"A6 output file: {a6_path}",
    ]

    if a4_path:
        body_lines.append(f"A4 (full pre-filter extract) file: {a4_path}")
    else:
        body_lines.append("A4 (full pre-filter extract): not found for this batch")

    if a5b_path:
        body_lines.append(f"A5B (excluded rows from 04_validate.py) file: {a5b_path}")
    else:
        body_lines.append("A5B (excluded rows from 04_validate.py): not found for this batch")

    if needs_review:
        body_lines.append("")
        if review_rows_summary:
            body_lines.append("Article Code Check = Review - details:")
            body_lines.extend(f"  - {line}" for line in review_rows_summary)
            body_lines.append("")
        if missing_location_summary:
            body_lines.append("Missing Location match - details:")
            body_lines.extend(f"  - {line}" for line in missing_location_summary)
            body_lines.append("")
        if duplicate_summary:
            body_lines.append("Duplicate Article Code - details:")
            body_lines.extend(f"  - {line}" for line in duplicate_summary)

    return subject, body_lines


def send_email(cfg, subject, body_lines, attachment_paths=None):
    """
    Shared SMTP sender used for the batch completion email. Until
    [notifications] smtp_host is filled in in config.conf, this just
    prints the would-be email content (including the Attachments
    section below) to the console instead of sending anything - safe
    to run as-is.

    attachment_paths, if given, is a list of files to attach as-is
    (e.g. A6 and, when found in the output folder, A5B). Each one is
    skipped with a warning if the file is missing or larger than
    MAX_ATTACHMENT_MB. Regardless of what actually gets attached, the
    email body always gets an explicit "Attachments:" section listing
    exactly what was (and wasn't) sent, so the technician never has to
    guess from the subject line alone.
    """
    attachment_paths = attachment_paths or []

    smtp_host = cfg.get("notifications", "smtp_host", fallback="").strip()
    recipients = [r.strip() for r in cfg.get("notifications", "to_addresses", fallback="").split(",") if r.strip()]

    # Work out attachment status up front so it can be reported in the
    # body whether or not sending actually succeeds.
    attachment_status = []  # list of (path, included, reason_if_not)
    for path in attachment_paths:
        if not os.path.exists(path):
            attachment_status.append((path, False, "file not found"))
            continue
        size_mb = os.path.getsize(path) / (1024 * 1024)
        if size_mb > MAX_ATTACHMENT_MB:
            attachment_status.append((path, False, f"{size_mb:.1f} MB exceeds the {MAX_ATTACHMENT_MB} MB guard"))
        else:
            attachment_status.append((path, True, f"{size_mb:.2f} MB"))

    attach_lines = ["", "Attachments:"]
    if not attachment_status:
        attach_lines.append("  (none)")
    for path, included, note in attachment_status:
        name = os.path.basename(path)
        if included:
            attach_lines.append(f"  - {name} ({note}) - attached")
        else:
            attach_lines.append(f"  - {name} - NOT attached ({note})")

    full_body = list(body_lines) + attach_lines

    if not smtp_host:
        print(f"\n[Email notification skipped] No [notifications] smtp_host configured "
              f"in config.conf yet - fill in SMTP settings to enable real emails. "
              f"Printing the notification content instead:\n")
        print(f"--- Subject: {subject} ---")
        for line in full_body:
            print(line)
        print("--- end notification content ---\n")
        return

    if not recipients:
        print("[Email notification skipped] No [notifications] to_addresses configured.")
        return

    smtp_port = cfg.getint("notifications", "smtp_port", fallback=587)
    smtp_user = cfg.get("notifications", "smtp_user", fallback="")
    smtp_password = cfg.get("notifications", "smtp_password", fallback="")
    use_tls = cfg.getboolean("notifications", "use_tls", fallback=True)
    sender = cfg.get("notifications", "from_address", fallback=smtp_user)

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg.attach(MIMEText("\n".join(full_body)))

    attached_names = []
    for path, included, note in attachment_status:
        if not included:
            print(f"  WARNING: attachment '{path}' not attached ({note}).")
            continue
        with open(path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{os.path.basename(path)}"',
        )
        msg.attach(part)
        attached_names.append(os.path.basename(path))

    try:
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            if use_tls:
                server.starttls()
            if smtp_user:
                server.login(smtp_user, smtp_password)
            server.sendmail(sender, recipients, msg.as_string())
        print(f"Notification email sent to: {', '.join(recipients)}"
              + (f" (attachments: {', '.join(attached_names)})" if attached_names else " (no attachments)"))
    except (smtplib.socket.gaierror, ConnectionRefusedError, TimeoutError) as e:
        print(f"  WARNING: could not reach SMTP server '{smtp_host}:{smtp_port}' ({e}).\n"
              f"  This usually means smtp_host in config.conf is wrong, is still a "
              f"placeholder, or isn't reachable from this machine/network. Check "
              f"[notifications] smtp_host in config.conf, or ask IT for the correct "
              f"mail server address.")
        _print_email_fallback(subject, full_body)
    except smtplib.SMTPAuthenticationError as e:
        print(f"  WARNING: SMTP login failed ({e}). Check [notifications] smtp_user "
              f"and smtp_password in config.conf.")
        _print_email_fallback(subject, full_body)
    except Exception as e:
        print(f"  WARNING: failed to send notification email ({e}).")
        _print_email_fallback(subject, full_body)


def _print_email_fallback(subject, body_lines):
    """Used whenever sending fails, so the content is never lost even if SMTP is broken."""
    print(f"\n  --- Email content that could not be sent (Subject: {subject}) ---")
    for line in body_lines:
        print(f"  {line}")
    print("  --- end email content ---\n")


def main():
    cfg = load_config()
    paths = get_paths(cfg)

    meta_arg = sys.argv[1] if len(sys.argv) > 1 else None
    meta_path = meta_arg or find_latest_meta(paths["output_dir"])
    print(f"Reading meta file: {meta_path}")

    with open(meta_path, "r", encoding="utf-8") as f:
        meta = json.load(f)

    required = [
        "week_tag", "date_tag", "output_dir", "total_rows", "n_clean", "n_review",
        "n_missing_location", "n_duplicate", "review_rows_summary",
        "missing_location_summary", "duplicate_summary", "a6_path",
    ]
    missing = [k for k in required if k not in meta]
    if missing:
        raise KeyError(f"Meta file '{meta_path}' is missing expected field(s): {missing}.")

    a4_path = find_existing_a4(meta["output_dir"], meta["date_tag"])
    a5b_path = find_existing_a5b(meta["output_dir"], meta["date_tag"])

    subject, body_lines = build_completion_email(
        meta["week_tag"], meta["date_tag"], meta["total_rows"], meta["n_clean"],
        meta["n_review"], meta["n_missing_location"], meta["n_duplicate"],
        meta["review_rows_summary"], meta["missing_location_summary"],
        meta["duplicate_summary"], meta["a6_path"], a5b_path, a4_path,
    )
    attachment_paths = (
        [meta["a6_path"]]
        + ([a4_path] if a4_path else [])
        + ([a5b_path] if a5b_path else [])
    )
    send_email(cfg, subject, body_lines, attachment_paths=attachment_paths)

    print("\nDone. If SMTP wasn't configured yet, the notification content was "
          "printed above instead of sent - fill in [notifications] in config.conf "
          "to enable real emails.")


if __name__ == "__main__":
    main()
