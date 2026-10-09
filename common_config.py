"""
common_config.py
-----------------
Shared helper to load settings from config.conf.
Imported by every numbered script in this pipeline, keeps the
connection/path/rule settings in ONE place instead of duplicated
across files.
"""
import configparser
import os
from datetime import datetime, timedelta

def load_config(path="config.conf"):
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Config file not found at '{path}'. Edit config.conf and "
            f"fill in your SFTP credentials first."
        )
    cfg = configparser.ConfigParser()
    cfg.read(path)
    return cfg

# Secrets: AWS Parameter Store (EC2) with local config.conf fallback (dev_local)
DEFAULT_AWS_REGION = "us-east-1"


def get_secret(param_name, region=None):
    """
    Fetch and decrypt a SecureString from AWS Systems Manager Parameter
    Store.
    """
    try:
        import boto3
    except ImportError as exc:
        raise RuntimeError(
            f"A Parameter Store name ('{param_name}') is configured but boto3 "
            f"is not installed. Run: pip install boto3"
        ) from exc

    try:
        ssm = boto3.client("ssm", region_name=region or DEFAULT_AWS_REGION)
        response = ssm.get_parameter(Name=param_name, WithDecryption=True)
        return response["Parameter"]["Value"]
    except Exception as exc:  # botocore errors: no creds, AccessDenied, ParameterNotFound...
        raise RuntimeError(
            f"Could not read parameter '{param_name}' from AWS Parameter Store "
            f"({type(exc).__name__}: {exc}). Check the parameter name, the AWS "
            f"region, and that this machine's IAM role/credentials allow "
            f"ssm:GetParameter (and kms:Decrypt)."
        ) from exc


def get_aws_region(cfg):
    """Region from [aws] region in config.conf, default us-east-1."""
    return cfg.get("aws", "region", fallback=DEFAULT_AWS_REGION).strip() or DEFAULT_AWS_REGION


def resolve_secret(cfg, section, plain_key, param_key):
    """
    Return a secret using this precedence:
      1. [section] <param_key>  -> read from AWS Parameter Store   (EC2)
      2. [section] <plain_key>  -> plain value in config.conf      (laptop)
    Returns "" if neither is set. The secret is never printed or logged.
    """
    param_name = cfg.get(section, param_key, fallback="").strip()
    if param_name:
        return get_secret(param_name, region=get_aws_region(cfg))
    return cfg.get(section, plain_key, fallback="")


def get_paths(cfg):
    """
    Return (and create) the local working folders defined in
    config.conf. templates_dir is included alongside download_root /
    work_dir / output_dir so it's configured and auto-created the same
    way as the others - individual template files (see the
    [templates] section) default to living inside it, but can still be
    pointed elsewhere entirely via their own explicit path in
    config.conf if needed.
    """
    download_root = cfg.get("local", "download_root", fallback="./downloads")
    work_dir = cfg.get("local", "work_dir", fallback="./work")
    output_dir = cfg.get("local", "output_dir", fallback="./output")
    templates_dir = cfg.get("local", "templates_dir", fallback="./templates")

    for p in (download_root, work_dir, output_dir, templates_dir):
        os.makedirs(p, exist_ok=True)

    return {
        "download_root": download_root,
        "work_dir": work_dir,
        "output_dir": output_dir,
        "templates_dir": templates_dir,
    }


def compute_week_tag(dt=None):
    """
    Returns the ISO-week tag used for EVERY dated subfolder in this
    pipeline (download_root/<tag>/, work_dir/<tag>/, output_dir/<tag>/),
    e.g. "Y26W25".

    IMPORTANT: this is the ISO week BEFORE the one this script is
    actually run in, not the current week. MSF Logistique's weekly
    extract for a given week only lands on the SFTP server (and gets
    processed) during the FOLLOWING week - so if this pipeline is run
    during ISO week 26, the file being processed is last week's data
    (week 25), and every folder for that batch is tagged "Y26W25", not
    "Y26W26". Subtracting 7 days before taking isocalendar() also
    handles year boundaries correctly (e.g. week 1 of a new year
    correctly rolls back to week 52/53 of the previous year).

    This is the ONE place this calculation is made. Every script in
    the pipeline (01/02/04/05) imports this function instead of
    defining its own copy, so there is no risk of the scripts drifting
    out of sync with each other.
    """
    dt = dt or datetime.now()
    previous_week_dt = dt - timedelta(days=7)
    iso_year, iso_week, _ = previous_week_dt.isocalendar()
    return f"Y{iso_year % 100:02d}W{iso_week:02d}"


def get_dated_subdir(root, date_tag):
    """
    Return (and create) a subfolder under the given root, named with
    the pipeline's shared week tag (e.g. "Y26W26" for ISO week 26 of
    2026), NOT a per-day date. This is the same week_tag computed by
    compute_week_tag() in 01_download_sftp.py (and duplicated in
    02_read_files.py / 04_validate.py / 05_export.py), so download_root,
    work_dir, and output_dir all end up with matching subfolder names
    for a given batch - e.g. downloads/Y26W26/, work/Y26W26/,
    output/Y26W26/ - rather than scattering across per-run or per-day
    folders. Safe to call multiple times for the same week; it reuses
    the same folder instead of overwriting or duplicating it.
    """
    dated_path = os.path.join(root, date_tag)
    os.makedirs(dated_path, exist_ok=True)
    return dated_path
