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

  - Extraction summary: total rows extracted from MSF Logistique
    (i.e. every row in A4) and how each one was worked on - imported
    via A6, Kit at a deployed mission (review required), Kit not
    deployed, or not deployed. See summarise_extraction().
  - A "review needed" subject/headline whenever anything needs a
    person to look at it - including any Kit article to be assessed.
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
import re
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from common_config import load_config, get_paths, resolve_secret

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


def summarise_extraction(a4_path):
    """
    Counts every row MSF Logistique sent us (= every row in A4) by how
    it was worked on, using A4's 'Keep?' column:

      blank                    -> passed all checks; validated (04) and
                                  exported to A6 (05) for TMS import
      "NOT ASSET"              -> Kit at a deployed mission; excluded
                                  from A6, REVIEW REQUIRED (must be
                                  broken down into component assets)
      "NOT ASSET; NOT DEPLOYED"-> Kit that never reached a valid
                                  mission; excluded, nothing to do
      "NOT DEPLOYED"           -> non-Kit, not a deployed mission;
                                  excluded, nothing to do

    Matching is by substring (same idea as the row colouring in
    03/04), so the wording of the combined value can change without
    breaking the counts. Anything non-blank that matches neither reason
    lands in 'other' so the numbers always add up to 'total'.

    Returns a dict, or None if A4 is missing/unreadable (the email
    still goes out, just without the summary).
    """
    if not a4_path or not os.path.exists(a4_path):
        return None
    try:
        df = pd.read_excel(a4_path, dtype=str, keep_default_na=True)
    except Exception as e:
        print(f"  WARNING: could not read '{a4_path}' for the extraction summary ({e}). "
              f"Sending the email without it.")
        return None
    if "Keep?" not in df.columns:
        print(f"  WARNING: '{a4_path}' has no 'Keep?' column - can't build the "
              f"extraction summary. Sending the email without it.")
        return None

    keep = df["Keep?"].fillna("").astype(str).str.strip().str.upper()
    is_kit = keep.str.contains(NOT_ASSET, regex=False)
    not_dep = keep.str.contains(NOT_DEPLOYED, regex=False)
    blank = keep == ""

    return {
        "total": int(len(df)),
        "imported": int(blank.sum()),
        "kit_deployed": int((is_kit & ~not_dep).sum()),
        "kit_not_deployed": int((is_kit & not_dep).sum()),
        "not_deployed": int((~is_kit & not_dep).sum()),
        "other": int((~blank & ~is_kit & ~not_dep).sum()),
    }


# ---------------------------------------------------------------------------
# Email layout helpers
#
# The completion email is laid out in this order so the reader gets the
# actionable part first and every number appears once:
#
#   1. Headline + one-line "needs attention" status
#   2. WHAT NEEDS YOUR ATTENTION - numbered, grouped (not row by row)
#   3. BATCH BREAKDOWN           - what happened to every extracted row
#   4. FILES                     - output folder + notes on missing files
#      (send_email() then appends the "Attachments:" list)
# ---------------------------------------------------------------------------

MAX_ITEMS_SHOWN = 5   # serials / project codes listed inline before "+N more"

_REVIEW_RE = re.compile(r"Article Code '([^']*)' \(Serial '([^']*)'\)")
_MISSING_LOC_RE = re.compile(r"Article Code '([^']*)' \(Project Code '([^']*)'\).*?prefix '([^']*)'")
_DUPLICATE_RE = re.compile(r"Article Code '([^']*)' \(Serial '([^']*)'\)")


def _group(lines, regex, key_group=1, value_group=2):
    """
    Groups the per-row summary strings written by 05_export.py by one of
    their captured fields, keeping first-seen order.

    Returns (groups, unparsed): groups is {key: [values...]}; unparsed is
    any line that didn't match (e.g. if 05's wording changes) so nothing
    is ever silently dropped from the email.
    """
    groups, unparsed = {}, []
    for line in lines:
        m = regex.search(line)
        if not m:
            unparsed.append(line)
            continue
        groups.setdefault(m.group(key_group), []).append(m.group(value_group))
    return groups, unparsed


def _inline(items, label):
    """' - label: a, b, c (+N more)' for short lists, '' when there's nothing to show."""
    items = [str(i) for i in items if str(i).strip()]
    if not items:
        return ""
    shown = ", ".join(items[:MAX_ITEMS_SHOWN])
    more = len(items) - MAX_ITEMS_SHOWN
    return f" - {label}: {shown}" + (f" (+{more} more)" if more > 0 else "")


def _unparsed_lines(unparsed):
    out = [f"   - {line}" for line in unparsed[:MAX_ITEMS_SHOWN]]
    if len(unparsed) > MAX_ITEMS_SHOWN:
        out.append(f"   - ... and {len(unparsed) - MAX_ITEMS_SHOWN} more (see A6)")
    return out


def build_attention_blocks(n_review, review_rows_summary, n_missing_location,
                           missing_location_summary, n_duplicate, duplicate_summary,
                           kit_only_summary):
    """
    Returns a list of (heading, detail_lines) for everything a person
    needs to look at, most important first. Empty list = nothing to do.
    Repeated rows are grouped (by Article Code / location prefix) so a
    batch with 45 identical-code rows reads as one line, not 45.
    """
    blocks = []

    if kit_only_summary:
        blocks.append((
            f"Kit articles to assess ({len(kit_only_summary)}) - excluded from A6",
            ["   These Kits reached a valid TMS mission. Break each one into its component",
             "   assets and import those into TMS separately."]
            + [f"   - {line}" for line in kit_only_summary],
        ))

    if n_review:
        groups, unparsed = _group(review_rows_summary, _REVIEW_RE)
        detail = ["   Included in A6 and highlighted - look each Article Code up in TMS."]
        for code, serials in groups.items():
            n = len(serials)
            detail.append(f"   - {code}: {n} row(s)" + (_inline(serials, "serials") if n <= MAX_ITEMS_SHOWN else ""))
        detail += _unparsed_lines(unparsed)
        blocks.append((
            f"Article Code not found in TMS - manual lookup ({n_review} row(s), {len(groups) or len(unparsed)} code(s))",
            detail,
        ))

    if n_missing_location:
        groups, unparsed = _group(missing_location_summary, _MISSING_LOC_RE, key_group=3, value_group=2)
        detail = ["   Location left blank in A6 - fill in by hand, or add the prefix to templates/location.xlsx."]
        for prefix, projects in groups.items():
            uniq = list(dict.fromkeys(projects))
            detail.append(f"   - Prefix '{prefix}': {len(projects)} row(s)" + _inline(uniq, "project codes"))
        detail += _unparsed_lines(unparsed)
        blocks.append((f"Missing Location match ({n_missing_location} row(s))", detail))

    if n_duplicate:
        groups, unparsed = _group(duplicate_summary, _DUPLICATE_RE)
        detail = ["   Same Article Code AND Serial appear more than once - TMS will block these",
                  "   on import. Check they aren't duplicate entries of the same item."]
        for code, serials in groups.items():
            detail.append(f"   - {code}: {len(serials)} row(s)" + _inline(list(dict.fromkeys(serials)), "serials"))
        detail += _unparsed_lines(unparsed)
        blocks.append((f"Duplicate Article Code + Serial ({n_duplicate} row(s))", detail))

    return blocks


def _status_line(kit_only_summary, n_review, n_missing_location, n_duplicate):
    parts = []
    if kit_only_summary:
        parts.append(f"{len(kit_only_summary)} Kit(s) to assess")
    if n_review:
        parts.append(f"{n_review} row(s) with an Article Code to look up")
    if n_missing_location:
        parts.append(f"{n_missing_location} row(s) missing a Location")
    if n_duplicate:
        parts.append(f"{n_duplicate} duplicate row(s)")
    if not parts:
        return "No review needed."
    return "REVIEW REQUIRED: " + "; ".join(parts) + "."


def build_breakdown_lines(summary, n_in_a6=None, n_clean=None):
    """
    "BATCH BREAKDOWN" block: total rows extracted from MSF Logistique
    (every row in A4) and how each was worked on. Every number appears
    here once; the rest of the email refers back to it rather than
    repeating it. n_in_a6 is only a cross-check against A4's blank-'Keep?'
    row count.
    """
    if not summary:
        return ["BATCH BREAKDOWN", "   Not available (A4 could not be read for this batch).", ""]

    imported = summary["imported"]
    imported_note = ""
    if n_clean is not None:
        imported_note = f" ({n_clean} clean, {imported - n_clean} flagged for review)"

    lines = [
        "BATCH BREAKDOWN",
        f"Rows extracted from MSF Logistique: {summary['total']}",
        f"   - Imported via A6: {imported}{imported_note}",
        f"   - Kit at a deployed mission - to assess, excluded from A6: {summary['kit_deployed']}",
        f"   - Kit not deployed - excluded, no action: {summary['kit_not_deployed']}",
        f"   - Not deployed (non-Kit) - excluded, no action: {summary['not_deployed']}",
    ]
    if summary["other"]:
        lines.append(f"   - Excluded for another reason: {summary['other']}")
    lines.append(f"   Excluded in total (listed in A5B): {summary['total'] - imported}")

    if n_in_a6 is not None and n_in_a6 != imported:
        lines.append(
            f"   NOTE: A6 has {n_in_a6} row(s) but A4 has {imported} row(s) with a blank "
            f"'Keep?' - these should match, please check this batch."
        )
    lines.append("")
    return lines


def _files_lines(a6_path, a4_path, a5b_path):
    """One 'output folder' line plus a note for anything expected but missing."""
    lines = ["FILES"]
    folder_src = a6_path or a4_path or a5b_path
    if folder_src:
        lines.append(f"Output folder: {os.path.dirname(folder_src)}")
    if not a4_path:
        lines.append("Note: A4 (full pre-filter extract) was not found for this batch.")
    if not a5b_path:
        lines.append("Note: A5B (excluded rows) was not found for this batch.")
    return lines


def build_completion_email(week_tag, date_tag, total_rows, n_clean, n_review, n_missing_location,
                           n_duplicate, review_rows_summary, missing_location_summary,
                           duplicate_summary, a6_path, a5b_path, a4_path, kit_only_summary,
                           extraction_summary=None):
    """
    The normal (A6) completion email - see the layout note above.

    Anything a person has to act on (a Kit to assess, Article Code
    'Review' rows, missing Location, duplicates) makes the subject read
    "review needed" and is listed first, grouped, under WHAT NEEDS YOUR
    ATTENTION. A Kit that reached a deployed mission counts on its own:
    it always needs assessing, independent of anything else in A6.

    For a batch with zero rows to import, see
    build_nothing_to_import_email() instead.
    """
    kit_only_summary = kit_only_summary or []
    needs_review = bool(n_review or n_missing_location or n_duplicate or kit_only_summary)

    if needs_review:
        subject = f"[MSF TMS Import] {week_tag} - Batch {date_tag} completed - review needed"
    else:
        subject = f"[MSF TMS Import] {week_tag} - Batch {date_tag} completed successfully"

    if extraction_summary:
        headline = (f"Batch {date_tag} (folder {week_tag}): {extraction_summary['total']} rows extracted "
                    f"from MSF Logistique, {total_rows} imported via A6.")
    else:
        headline = f"Batch {date_tag} (folder {week_tag}): {total_rows} item(s) processed into A6."

    body_lines = [headline,
                  _status_line(kit_only_summary, n_review, n_missing_location, n_duplicate),
                  ""]

    blocks = build_attention_blocks(n_review, review_rows_summary, n_missing_location,
                                    missing_location_summary, n_duplicate, duplicate_summary,
                                    kit_only_summary)
    if blocks:
        body_lines.append("WHAT NEEDS YOUR ATTENTION")
        for i, (heading, detail) in enumerate(blocks, start=1):
            body_lines.append(f"{i}. {heading}")
            body_lines.extend(detail)
            body_lines.append("")

    body_lines.extend(build_breakdown_lines(extraction_summary, n_in_a6=total_rows, n_clean=n_clean))
    body_lines.extend(_files_lines(a6_path, a4_path, a5b_path))
    return subject, body_lines


def build_nothing_to_import_email(week_tag, date_tag, total_rows, a5b_path, a4_path,
                                  kit_only_summary=None, extraction_summary=None):
    """
    Completion email for a batch where 04_validate.py found zero rows
    with a blank 'Keep?' - nothing new for TMS this run, so no A6
    exists. Same layout as build_completion_email(); still flags any
    Kit at a deployed mission, since that needs a person even when
    there is nothing to import.
    """
    kit_only_summary = kit_only_summary or []
    if kit_only_summary:
        subject = f"[MSF TMS Import] {week_tag} - Batch {date_tag} - nothing to import - review needed"
    else:
        subject = f"[MSF TMS Import] {week_tag} - Batch {date_tag} - nothing to import"

    headline = (f"Batch {date_tag} (folder {week_tag}): {total_rows} rows extracted from MSF Logistique, "
                f"none to import into TMS - every row already had a non-blank 'Keep?'. "
                f"No A6 file was produced.")
    status = (_status_line(kit_only_summary, 0, 0, 0) if kit_only_summary
              else "No action needed unless this is unexpected - check A4/A5B for what came in.")

    body_lines = [headline, status, ""]
    blocks = build_attention_blocks(0, [], 0, [], 0, [], kit_only_summary)
    if blocks:
        body_lines.append("WHAT NEEDS YOUR ATTENTION")
        for i, (heading, detail) in enumerate(blocks, start=1):
            body_lines.append(f"{i}. {heading}")
            body_lines.extend(detail)
            body_lines.append("")

    body_lines.extend(build_breakdown_lines(extraction_summary))
    body_lines.extend(_files_lines(None, a4_path, a5b_path))
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
    # EC2: smtp_password_param (AWS Parameter Store). Dev_local: smtp_password in config.conf.
    try:
        smtp_password = resolve_secret(cfg, "notifications", "smtp_password", "smtp_password_param")
    except Exception as e:
        print(f"  WARNING: could not load the SMTP password ({e}).")
        _print_email_fallback(subject, full_body)
        return
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
              f"and smtp_password (or smtp_password_param) in config.conf.")
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

    kit_only_summary = find_kit_only_rows(a5b_path)
    extraction_summary = summarise_extraction(a4_path)

    subject, body_lines = build_nothing_to_import_email(
        meta["week_tag"], meta["date_tag"], meta["total_rows"], a5b_path, a4_path,
        kit_only_summary=kit_only_summary, extraction_summary=extraction_summary,
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

    extraction_summary = summarise_extraction(a4_path)

    subject, body_lines = build_completion_email(
        meta["week_tag"], meta["date_tag"], meta["total_rows"], meta["n_clean"],
        meta["n_review"], meta["n_missing_location"], meta["n_duplicate"],
        meta["review_rows_summary"], meta["missing_location_summary"],
        meta["duplicate_summary"], meta["a6_path"], a5b_path, a4_path,
        kit_only_summary, extraction_summary,
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
