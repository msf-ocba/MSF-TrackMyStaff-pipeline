"""
06_notify.py
--------------
Sends the batch completion email after 04_validate.py or 05_export.py
has finished.

The script reads a batch meta JSON and uses it to locate the generated
output files, then sends a single completion email summarising the
batch. Two kinds of meta file are recognised:

  - A6_META_EXTRACTION_TMS_OCBA_YYMMDD.json
      Written by 05_export.py after producing A6 - the normal case.
      Triggers the full completion email described below (batch
      statistics, Article Code Review / Location / duplicate warnings,
      Kit-only section, attachments).

  - A5_META_EXTRACTION_TMS_OCBA_YYMMDD.json
      Written by 04_validate.py instead, when a batch has ZERO rows
      with a blank 'Keep?' - i.e. nothing new to import into TMS this
      run (every row was already "NOT ASSET" / "NOT DEPLOYED" / both).
      There is no A6 for this batch. Triggers a distinct, shorter
      "nothing to import" email instead - still attaching A4/A5B so
      the batch is fully accounted for - since technicians rely on
      getting an email either way, not on checking server logs to see
      whether the pipeline ran.

Whichever meta file is most recently modified is used automatically;
see find_latest_meta().

FILES INCLUDED (normal / A6 case)

A6 - Imported Records
  The final import workbook created by 05_export.py. This is the
  primary attachment and contains the records prepared for import into
  TMS.

A4 - Full Pre-filter Extract
  If present, the original batch extract is attached so the technician
  has the complete source data alongside the final import file.

A5B - Excluded Records
  If present, the workbook containing rows excluded during validation
  is attached to provide visibility of records that were not included
  in the import.

KIT-ONLY ITEMS

The script identifies rows in A5B marked only as 'NOT ASSET'. These
represent Kit articles that reached a valid deployed mission but must
be manually broken down into their component assets before being
imported into TMS. Any such rows are listed in a dedicated section of
the email for easy reference.

EMAIL SUMMARY

A single completion email is sent for every batch. The email includes:

  - Overall batch statistics.
  - Any Article Code Review, missing Location, or duplicate Article
    Code warnings recorded during export.
  - A dedicated Kit-only section when applicable.
  - An Attachments section listing every file included with the email,
    together with any expected files that could not be attached.

(For the "nothing to import" case - see A5_META_....json above - the
email is a shorter, distinct summary instead; see
build_nothing_to_import_email().)

SMTP server and recipient settings are read from the
[notifications] section of config.conf.

Usage:
    python 06_notify.py [path_to_meta_json]

If no metadata file is supplied, the script automatically locates the
most recently generated meta file (A6_META_... or A5_META_...) in the
output folders.
"""
import os
import sys
import glob
import json
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from common_config import load_config, get_paths

MAX_ATTACHMENT_MB = 20  # most SMTP servers/relays reject attachments above ~20-25MB

# Mirrors 03_clean_transform.py / 04_validate.py's Keep? reason string.
# Needed here only to pull the "NOT ASSET only" rows back out of A5B
# for the Kit call-out section below. Kept in sync manually, same as
# the rest of this pipeline's per-script duplication.
NOT_ASSET = "NOT ASSET"
NOT_DEPLOYED = "NOT DEPLOYED"

# Glob patterns for the two kinds of meta file this script understands
# - see the module docstring for what each means.
META_PATTERNS = (
    "A6_META_EXTRACTION_TMS_OCBA_*.json",
    "A5_META_EXTRACTION_TMS_OCBA_*.json",
)


def find_latest_meta(output_dir):
    """
    Same dated-subfolder convention as 05_export.py's find_latest_a5():
    looks in the most recently modified dated subfolder(s) under
    output_dir for a meta file. Both A6_META_....json (normal case)
    and A5_META_....json ("nothing to import" case - see module
    docstring) are considered together; whichever individual file is
    most recently modified wins, so a same-week rerun always picks up
    the latest outcome regardless of which kind of batch it was.
    """
    subfolders = [f for f in glob.glob(os.path.join(output_dir, "*")) if os.path.isdir(f)]
    subfolders.sort(key=os.path.getmtime, reverse=True)
    for folder in subfolders:
        candidates = []
        for pattern in META_PATTERNS:
            candidates.extend(glob.glob(os.path.join(folder, pattern)))
        if candidates:
            return max(candidates, key=os.path.getmtime)

    direct = []
    for pattern in META_PATTERNS:
        direct.extend(glob.glob(os.path.join(output_dir, pattern)))
    if direct:
        chosen = max(direct, key=os.path.getmtime)
        print(f"  WARNING: no dated subfolder under '{output_dir}' contains a meta file. "
              f"Falling back to a flat file found directly in '{output_dir}': {chosen}.")
        return chosen

    raise FileNotFoundError(
        f"No A6_META_EXTRACTION_TMS_OCBA_*.json or A5_META_EXTRACTION_TMS_OCBA_*.json "
        f"file found under '{output_dir}' (checked dated subfolders and the flat "
        f"folder). Run 05_export.py (or, for a batch with nothing to import, "
        f"04_validate.py) first."
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
    the email still goes out with just A6 attached, and no Kit call-out
    section, in that case).
    """
    expected_path = os.path.join(output_dir, f"A5B_EXTRACTION_TMS_OCBA_{date_tag}.xlsx")
    if os.path.exists(expected_path):
        return expected_path

    print(f"  WARNING: expected A5B file not found at '{expected_path}'. "
          f"A5B is produced by 04_validate.py - make sure it was run for this "
          f"batch (date tag '{date_tag}') and wrote into this same output folder. "
          f"Continuing without it - the email will only have A6 attached, and "
          f"won't be able to list any Kit-only rows (see find_kit_only_rows()).")
    return None


def find_kit_only_rows(a5b_path):
    """
    Pulls the "NOT ASSET only" rows back out of A5B for the completion
    email's Kit call-out section (see module docstring).

    A5B's 'Keep?' column can read "NOT ASSET", "NOT DEPLOYED", or
    "NOT ASSET; NOT DEPLOYED" (see 03_clean_transform.py /
    04_validate.py). We want ONLY the exact-match "NOT ASSET" rows
    here - Kits that DID reach a valid, deployed TMS mission and so
    still need to be manually exploded into their component assets and
    imported. Rows also flagged "NOT DEPLOYED" are excluded from this
    section: they never reached a valid mission in the first place, so
    there is nothing to explode/import for TMS.

    Returns a list of one-line summaries (e.g. "Article Code 'K1234ABC'
    - Mission 'CF160MES' - Qty 3"), or an empty list if a5b_path is
    None (A5B wasn't found), the file can't be read, or it has none of
    the expected columns / no matching rows.
    """
    if not a5b_path or not os.path.exists(a5b_path):
        return []

    try:
        df = pd.read_excel(a5b_path, dtype=str, keep_default_na=True)
    except Exception as e:
        print(f"  WARNING: could not read '{a5b_path}' to look for Kit-only rows ({e}). "
              f"Skipping the Kit call-out section for this email.")
        return []

    if "Keep?" not in df.columns or "FCL_ART_CODE" not in df.columns:
        print(f"  WARNING: '{a5b_path}' is missing the 'Keep?' and/or 'FCL_ART_CODE' "
              f"column(s) - can't identify Kit-only rows. Skipping the Kit call-out "
              f"section for this email.")
        return []

    keep_col = df["Keep?"].astype(str).str.strip()
    kit_only_mask = keep_col == NOT_ASSET  # exact match only - excludes "NOT ASSET; NOT DEPLOYED"

    lines = []
    for _, row in df.loc[kit_only_mask].iterrows():
        code = row.get("FCL_ART_CODE", "")
        mission = row.get("PCT_CLI_CODE_LIV", "")
        qty = row.get("QTE", "")
        qty_str = f" - Qty {str(qty).strip()}" if pd.notna(qty) and str(qty).strip() else ""
        lines.append(f"Article Code '{code}' - Mission '{mission}'{qty_str}")

    return lines


def build_completion_email(week_tag, date_tag, total_rows, n_clean, n_review, n_missing_location,
                            n_duplicate, review_rows_summary, missing_location_summary,
                            duplicate_summary, a6_path, a5b_path, a4_path, kit_only_summary):
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

    kit_only_summary (see find_kit_only_rows()) is always listed, in
    its own section, whenever it's non-empty - independent of
    needs_review - since these are Kits that reached a deployed
    mission and still need a person to explode them into assets and
    import those into TMS; that's actionable regardless of whether
    anything else in the batch needs attention.

    (The actual Attachments section is appended later, in send_email,
    once it's known what was actually attached.)

    The subject leads with week_tag (e.g. "Y26W26") to match the
    output folder naming used throughout the pipeline; the headline
    and body still name the batch's own date_tag (e.g. "260629") for
    precise identification, since more than one batch can land in the
    same week folder.

    NOTE: this is the "normal" (A6) email. For a batch with zero rows
    to import, see build_nothing_to_import_email() instead.
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

    if kit_only_summary:
        body_lines.append("")
        body_lines.append(
            f"Kit articles excluded from A6, sent to a deployed mission ({len(kit_only_summary)}):"
        )
        body_lines.append(
            "  These are Kits (not individually flagged as undeployed) that reached a "
            "valid TMS mission - explode each one into its component assets and import "
            "those into TMS separately; the Kit line itself was excluded from A6 and is "
            "never imported as a single asset."
        )
        body_lines.extend(f"  - {line}" for line in kit_only_summary)

    return subject, body_lines


def build_nothing_to_import_email(week_tag, date_tag, total_rows, a5b_path, a4_path):
    """
    Completion email for a batch where 04_validate.py found zero rows
    with a blank 'Keep?' - there was nothing new to import into TMS
    this run (every row was already "NOT ASSET" / "NOT DEPLOYED" /
    both). No A6 exists for this batch, so this is a distinct, shorter
    email from build_completion_email() above: it confirms the batch
    ran, explains why nothing was imported, and still attaches
    A4/A5B so every source row is accounted for without anyone having
    to check server logs.
    """
    subject = f"[MSF TMS Import] {week_tag} - Batch {date_tag} - nothing to import"
    headline = (
        f"Batch {date_tag} (folder {week_tag}) ran successfully, but had nothing new "
        f"to import into TMS: all {total_rows} row(s) in this batch already had a "
        f"non-blank 'Keep?' (e.g. 'NOT ASSET' and/or 'NOT DEPLOYED')."
    )

    body_lines = [
        headline,
        "",
        f"Total items in this batch: {total_rows}",
        "Items imported into TMS (A6): 0 - no A6 file was produced this run.",
        "",
    ]

    if a4_path:
        body_lines.append(f"A4 (full pre-filter extract) file: {a4_path}")
    else:
        body_lines.append("A4 (full pre-filter extract): not found for this batch")

    if a5b_path:
        body_lines.append(f"A5B (excluded rows from 04_validate.py) file: {a5b_path}")
    else:
        body_lines.append("A5B (excluded rows from 04_validate.py): not found for this batch")

    body_lines.append("")
    body_lines.append(
        "No action is needed unless this is unexpected - if you were expecting new "
        "items this week, check A4/A5B (attached) for what came in and why every row "
        "was excluded."
    )

    return subject, body_lines


def send_email(cfg, subject, body_lines, attachment_paths=None):
    """
    Shared SMTP sender used for every completion email this script
    sends. Until [notifications] smtp_host is filled in in
    config.conf, this just prints the would-be email content
    (including the Attachments section below) to the console instead
    of sending anything - safe to run as-is.

    attachment_paths, if given, is a list of files to attach as-is
    (e.g. A6 and, when found in the output folder, A5B/A4 - or, for a
    "nothing to import" batch, just A4/A5B). Each one is skipped with
    a warning if the file is missing or larger than MAX_ATTACHMENT_MB.
    Regardless of what actually gets attached, the email body always
    gets an explicit "Attachments:" section listing exactly what was
    (and wasn't) sent, so the technician never has to guess from the
    subject line alone.
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


def _handle_nothing_to_import(cfg, meta, meta_path):
    """
    Handles the A5_META_....json case (see module docstring): a batch
    where 04_validate.py found nothing to import. Sends the shorter
    "nothing to import" email and returns.
    """
    required = ["week_tag", "date_tag", "output_dir", "total_rows", "a5b_path", "a4_path"]
    missing = [k for k in required if k not in meta]
    if missing:
        raise KeyError(f"Meta file '{meta_path}' is missing expected field(s): {missing}.")

    a4_path = meta["a4_path"]
    if a4_path and not os.path.exists(a4_path):
        print(f"  WARNING: A4 file recorded in meta ('{a4_path}') no longer exists. Continuing without it.")
        a4_path = None

    a5b_path = meta["a5b_path"]
    if a5b_path and not os.path.exists(a5b_path):
        print(f"  WARNING: A5B file recorded in meta ('{a5b_path}') no longer exists. Continuing without it.")
        a5b_path = None

    subject, body_lines = build_nothing_to_import_email(
        meta["week_tag"], meta["date_tag"], meta["total_rows"], a5b_path, a4_path
    )
    attachment_paths = ([a4_path] if a4_path else []) + ([a5b_path] if a5b_path else [])
    send_email(cfg, subject, body_lines, attachment_paths=attachment_paths)


def main():
    cfg = load_config()
    paths = get_paths(cfg)

    meta_arg = sys.argv[1] if len(sys.argv) > 1 else None
    meta_path = meta_arg or find_latest_meta(paths["output_dir"])
    print(f"Reading meta file: {meta_path}")

    with open(meta_path, "r", encoding="utf-8") as f:
        meta = json.load(f)

    if meta.get("outcome") == "nothing_to_import":
        _handle_nothing_to_import(cfg, meta, meta_path)
        print("\nDone. If SMTP wasn't configured yet, the notification content was "
              "printed above instead of sent - fill in [notifications] in config.conf "
              "to enable real emails.")
        return

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

    kit_only_summary = find_kit_only_rows(a5b_path)
    if kit_only_summary:
        print(f"  Found {len(kit_only_summary)} Kit-only (NOT ASSET only) row(s) in A5B - "
              f"listing them in their own section of the completion email.")

    subject, body_lines = build_completion_email(
        meta["week_tag"], meta["date_tag"], meta["total_rows"], meta["n_clean"],
        meta["n_review"], meta["n_missing_location"], meta["n_duplicate"],
        meta["review_rows_summary"], meta["missing_location_summary"],
        meta["duplicate_summary"], meta["a6_path"], a5b_path, a4_path,
        kit_only_summary,
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
