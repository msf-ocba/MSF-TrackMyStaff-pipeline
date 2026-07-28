"""
01_download_sftp.py
---------------------
Connects to the MSF Logistique SFTP server and downloads the latest
available extraction file for processing.

The downloaded file is saved to the pipeline's download folder, along
with a manifest file (downloaded_files.txt) recording the downloaded
file for use by the next stage of the pipeline.

The script verifies that the latest extraction is available before
downloading it. If no new extraction is ready, it exits cleanly so the
pipeline can be run again later without error.

Requires:
    paramiko

Usage:
    python 01_download_sftp.py
"""
import os
import re
import sys
import datetime
import paramiko

from common_config import load_config, get_paths, get_dated_subdir, compute_week_tag

# Exit code contract with run_pipeline.py (the master script):
#   0 = success, a file was downloaded, proceed with the rest of the pipeline
#   1 = real error (unparseable filename, SFTP/auth/connection failure, etc.)
#       - something needs a human's attention
#   2 = "not ready yet" - either there are no matching files on the server at
#       all, or the latest one isn't from the current week (this week's
#       extract hasn't landed yet). This is a NORMAL, EXPECTED outcome when
#       running on a schedule more often than the extract actually arrives -
#       run_pipeline.py treats this as "nothing to do this run" and stops
#       the rest of the pipeline cleanly, without treating it as a failure.
EXIT_OK = 0
EXIT_ERROR = 1
EXIT_NOT_READY = 2


# Matches a 6-digit YYMMDD date anywhere in the filename, e.g.
# "MSF_extract_240715_full.csv" -> "240715"
DATE_PATTERN = re.compile(r"(\d{6})")


def connect_sftp(cfg):
    host = cfg.get("sftp", "host")
    port = cfg.getint("sftp", "port", fallback=22)
    user = cfg.get("sftp", "user")
    password = cfg.get("sftp", "password", fallback="") or None
    key_path = cfg.get("sftp", "private_key_path", fallback="").strip() or None

    print(f"Connecting to {host}:{port} via SFTP ...")
    transport = paramiko.Transport((host, port))

    if key_path:
        pkey = paramiko.RSAKey.from_private_key_file(key_path, password=password)
        transport.connect(username=user, pkey=pkey)
    else:
        transport.connect(username=user, password=password)

    sftp = paramiko.SFTPClient.from_transport(transport)
    print("Connected.")
    return sftp, transport


def list_remote_files(sftp, cfg):
    remote_dir = cfg.get("sftp", "remote_dir", fallback="/")
    prefix = cfg.get("sftp", "file_prefix", fallback="")
    suffix = cfg.get("sftp", "file_suffix", fallback="")

    all_names = sftp.listdir(remote_dir)

    matched = [
        name for name in all_names
        if name.startswith(prefix) and name.endswith(suffix)
    ]
    matched.sort()  # filenames embed YYMMDD, so sorted = chronological
    return remote_dir, matched


def extract_date_from_filename(filename):
    """
    Pull a YYMMDD date out of the filename and return it as a
    datetime.date. Returns None if no date-like token is found.
    """
    match = DATE_PATTERN.search(filename)
    if not match:
        return None

    token = match.group(1)
    try:
        return datetime.datetime.strptime(token, "%y%m%d").date()
    except ValueError:
        return None


def pick_latest_file(filenames):
    """
    Given a chronologically-sorted list of filenames, return the last
    one (the latest), along with its parsed date. Raises ValueError
    if the latest filename has no parseable date.
    """
    latest_name = filenames[-1]
    latest_date = extract_date_from_filename(latest_name)
    if latest_date is None:
        raise ValueError(
            f"Could not extract a YYMMDD date from filename: {latest_name!r}. "
            "Refusing to guess — check file naming convention or DATE_PATTERN."
        )
    return latest_name, latest_date


def is_same_iso_week(date_a, date_b):
    """
    True if date_a and date_b fall in the same ISO year + ISO week.
    Using ISO week (Mon-Sun, isocalendar) avoids ambiguity around
    year boundaries and differing week-start conventions.
    """
    iso_a = date_a.isocalendar()  # (iso_year, iso_week, iso_weekday)
    iso_b = date_b.isocalendar()
    return (iso_a[0], iso_a[1]) == (iso_b[0], iso_b[1])


def download_one(sftp, remote_dir, filename, local_folder):
    remote_path = f"{remote_dir.rstrip('/')}/{filename}"
    local_path = os.path.join(local_folder, filename)
    print(f"  Downloading {remote_path} -> {local_path}")
    sftp.get(remote_path, local_path)
    return local_path


def main():
    cfg = load_config()
    paths = get_paths(cfg)

    today = datetime.date.today()

    # Download folder is named after the ISO week BEFORE the one this
    # script is RUN in (e.g. "Y26W25" if run during week 26) - see
    # common_config.compute_week_tag(). Matches the same convention
    # every later stage of the pipeline uses for its own folder - not
    # a per-run timestamp, so if this is re-run later in the same week
    # it lands in the same folder rather than scattering across
    # timestamped ones.
    week_tag = compute_week_tag()
    local_folder = get_dated_subdir(paths["download_root"], week_tag)

    sftp, transport = connect_sftp(cfg)
    try:
        remote_dir, remote_files = list_remote_files(sftp, cfg)
        if not remote_files:
            print("No matching files found on the server. Nothing to download.")
            sys.exit(EXIT_NOT_READY)

        print(f"Found {len(remote_files)} file(s) on server (up to ~4 weeks retained):")
        for name in remote_files:
            print(f"  - {name}")

        try:
            latest_name, latest_date = pick_latest_file(remote_files)
        except ValueError as e:
            print(f"ERROR: {e}")
            sys.exit(EXIT_ERROR)

        print(f"\nLatest file on server: {latest_name} (dated {latest_date.isoformat()})")
        print(f"Today's date:          {today.isoformat()}")

        if not is_same_iso_week(latest_date, today):
            iso_latest = latest_date.isocalendar()
            iso_today = today.isocalendar()
            print(
                "NOT READY: Latest file on the server is NOT from the current week.\n"
                f"  File's ISO week:  {iso_latest[0]}-W{iso_latest[1]:02d}\n"
                f"  Today's ISO week: {iso_today[0]}-W{iso_today[1]:02d}\n"
                "Refusing to download a stale file. Check whether this week's "
                "extract has been uploaded to the SFTP server yet. This is "
                "expected/normal if the cron job runs before the weekly extract "
                "has landed - it is not treated as a pipeline failure."
            )
            sys.exit(EXIT_NOT_READY)

        print("Latest file matches the current week. Proceeding with download.\n")

        downloaded = [download_one(sftp, remote_dir, latest_name, local_folder)]
    finally:
        sftp.close()
        transport.close()

    manifest_path = os.path.join(local_folder, "downloaded_files.txt")
    with open(manifest_path, "w") as mf:
        mf.write("\n".join(downloaded))

    print(f"\nDone. {len(downloaded)} file(s) saved to: {local_folder}")
    print(f"Manifest written to: {manifest_path}")
    print("\nNext step: run 02_read_files.py")


if __name__ == "__main__":
    main()
