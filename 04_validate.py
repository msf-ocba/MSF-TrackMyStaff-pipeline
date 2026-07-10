"""
04_validate.py
----------------
The script takes this batch's A4 workbook,
extracts only the rows where 'Keep?' is BLANK (i.e. rows that haven't
already been disqualified by the Family/customer/Mission checks in
03_clean_transform.py), and checks each row's FCL_ART_CODE against the
known article codes in the TMS article export.

Reproduces spec section 4, step 10 ("Check that all article codes are
available in TMS"), simplified to its core check for now:

  - FCL_ART_CODE found in the TMS article list -> "Checked"
  - FCL_ART_CODE NOT found (the spec's "N/A" case) -> "Review"

The TMS article list itself lives in a separate, easily-replaceable
file: templates/TMS_UniDataArticles.xlsx (re-export this from TMS
whenever you need fresher data - the column this script reads is
"Article Code", same as the TMS_UniDataArticles tab inside the B
template). No code changes needed to refresh it.

ARTICLE COMPOSE SUBSTITUTION (new):
Some MSF Logistique "Kit" articles are exported WITHOUT the leading
'K' that normally identifies a Kit family (see 03_clean_transform.py's
Family/NOT ASSET check). Because they don't have that leading 'K',
they pass the Family check as normal assets, but their raw
FCL_ART_CODE does not match how TMS lists them - so without a fix
they'd always come back "Review" even when TMS does recognize them.

templates/ArticleCompose.xlsx documents these: column "Article" is the
raw code as it appears in the extraction, column "ArticleCompose" is
the code TMS actually uses for the same item. Before the TMS check, we
look up each row's FCL_ART_CODE in the "Article" column; on a match we
replace FCL_ART_CODE with the "ArticleCompose" value so the TMS check
runs against the code TMS actually recognizes.

Output:
  - A5_EXTRACTION_TMS_OCBA_YYMMDD.xlsx, in the SAME dated work folder
    as A4, containing only the blank-Keep? rows from A4 plus
    "Article Code Check" ("Checked" / "Review"), with conditional
    formatting keyed off that column but applied to the ENTIRE ROW
    (green for Checked, red for Review), so a flagged row is obvious
    at a glance across all its fields, not just in that one column.
  - A5B_EXTRACTION_TMS_OCBA_YYMMDD.xlsx, written to the configured
    output_dir's <week_tag>/ subfolder (config.conf [local] output_dir,
    e.g. output/Y26W25/ - tagged with the ISO week BEFORE the one this
    script is RUN in, not the batch's own date tag - see
    resolve_output_dir() / common_config.compute_week_tag()). This is
    the same folder 05_export.py later writes A6 into, and from which
    06_notify.py emails A5B alongside A6. Contains the rows EXCLUDED
    from A5 (i.e. every A4 row whose Keep? was NOT blank). Keeps the
    same Keep?-based conditional formatting used on A4, unchanged.

Usage:
    python 04_validate.py [path_to_A4_workbook.xlsx] [path_to_tms_article_list.xlsx] [path_to_article_compose.xlsx]

If the A4 path is omitted, the script looks in the dated work
subfolders (same convention as 03_clean_transform.py) for the most
recent A4_*.xlsx. If the TMS article list / article compose paths are
omitted, the paths in config.conf ([templates] tms_article_list_path /
article_compose_path) are used.
"""
import os
import sys
import glob
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill

from common_config import load_config, get_paths, get_dated_subdir, compute_week_tag

TMS_ARTICLE_LIST_DEFAULT = "./templates/TMS_UniDataArticles.xlsx"
TMS_ARTICLE_CODE_COLUMN = "Article Code"

ARTICLE_COMPOSE_DEFAULT = "./templates/ArticleCompose.xlsx"
ARTICLE_COMPOSE_SOURCE_COL = "Article"
ARTICLE_COMPOSE_TARGET_COL = "ArticleCompose"

CHECKED = "Checked"
REVIEW = "Review"

# Mirrors 03_clean_transform.py's Keep? reason strings, needed here only
# to reproduce the same conditional-formatting rules on A5B (the
# excluded-rows workbook). Kept in sync manually, same as the rest of
# this pipeline's per-script duplication (see find_latest_a4 vs
# find_latest_a3).
NOT_ASSET = "NOT ASSET"
NOT_DEPLOYED = "NOT DEPLOYED"


def find_latest_a4(work_dir):
    """
    Same convention as 03_clean_transform.py's find_latest_a3(): prefer
    the most recently modified A4 found INSIDE a dated subfolder under
    work_dir. A flat work_dir/A4_....xlsx (not inside a dated folder)
    is only used as a last-resort fallback, and is flagged as
    potentially stale.
    """
    subfolders = [f for f in glob.glob(os.path.join(work_dir, "*")) if os.path.isdir(f)]
    subfolders.sort(key=os.path.getmtime, reverse=True)
    for folder in subfolders:
        candidates = sorted(glob.glob(os.path.join(folder, "A4_EXTRACTION_TMS_OCBA_*.xlsx")))
        if candidates:
            chosen = max(candidates, key=os.path.getmtime)
            direct = glob.glob(os.path.join(work_dir, "A4_EXTRACTION_TMS_OCBA_*.xlsx"))
            if direct:
                print(f"  NOTE: ignoring {len(direct)} A4 file(s) sitting directly in "
                      f"'{work_dir}' (not inside a dated subfolder) - these look stale "
                      f"and are not used. Using the dated subfolder copy instead: {chosen}")
            return chosen

    direct = sorted(glob.glob(os.path.join(work_dir, "A4_EXTRACTION_TMS_OCBA_*.xlsx")))
    if direct:
        chosen = max(direct, key=os.path.getmtime)
        print(f"  WARNING: no dated subfolder under '{work_dir}' contains an A4 file. "
              f"Falling back to a flat file found directly in '{work_dir}': {chosen}. "
              f"This is likely a leftover from before dated folders were introduced - "
              f"re-run 02/03 to regenerate it properly.")
        return chosen

    raise FileNotFoundError(
        f"No A4_EXTRACTION_TMS_OCBA_*.xlsx file found under '{work_dir}' "
        f"(checked dated subfolders and the flat folder). Run 03_clean_transform.py first."
    )


def resolve_template_path(path):
    """
    Look for a template file at the exact given path first. If not
    found, fall back to a case-insensitive filename match in the same
    folder - real exports/templates have been seen with inconsistent
    capitalization (e.g. "TMS_UNiDataArticles.xlsx"), and re-exports
    may vary in casing again later. Avoids a confusing
    FileNotFoundError over what is really just a capitalization
    mismatch. Used for both the TMS article list and ArticleCompose.
    """
    if os.path.exists(path):
        return path

    folder = os.path.dirname(path) or "."
    target_name = os.path.basename(path).lower()
    if os.path.isdir(folder):
        for name in os.listdir(folder):
            if name.lower() == target_name:
                resolved = os.path.join(folder, name)
                print(f"  NOTE: found '{name}' (case-insensitive match for "
                      f"'{os.path.basename(path)}') - using that.")
                return resolved

    return path  # nothing found; return original so the caller's error message is accurate


def load_tms_article_codes(path):
    path = resolve_template_path(path)
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"TMS article list not found at '{path}'. Export the current article "
            f"list from TMS and save it there (or pass a path as the 2nd argument), "
            f"keeping the '{TMS_ARTICLE_CODE_COLUMN}' column name."
        )
    df = pd.read_excel(path, dtype=str)
    if TMS_ARTICLE_CODE_COLUMN not in df.columns:
        raise KeyError(
            f"'{path}' has no '{TMS_ARTICLE_CODE_COLUMN}' column. Found columns: "
            f"{list(df.columns)}. Check the export still uses that header name."
        )
    codes = set(df[TMS_ARTICLE_CODE_COLUMN].dropna().astype(str).str.strip())
    return codes


def load_article_compose_map(path):
    """
    Load the raw-code -> TMS-composed-code lookup for "hidden" Kit
    articles (see module docstring). Returns a dict, e.g.
        {"EEMDCONA1200": "EEMDCONE12-", ...}
    Required, same as the TMS article list - if this file is missing
    or malformed we stop rather than silently skip the substitution,
    since that would just make those rows come back "Review" for the
    wrong reason.
    """
    path = resolve_template_path(path)
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Article compose template not found at '{path}'. This file maps raw "
            f"extraction codes for 'hidden' Kit articles (no leading 'K') to the "
            f"code TMS actually uses for them. Save it there (or pass a path as "
            f"the 3rd argument), keeping the '{ARTICLE_COMPOSE_SOURCE_COL}' / "
            f"'{ARTICLE_COMPOSE_TARGET_COL}' column names."
        )
    df = pd.read_excel(path, dtype=str)
    missing = [c for c in (ARTICLE_COMPOSE_SOURCE_COL, ARTICLE_COMPOSE_TARGET_COL) if c not in df.columns]
    if missing:
        raise KeyError(
            f"'{path}' is missing expected column(s): {missing}. Found columns: "
            f"{list(df.columns)}."
        )
    df = df.dropna(subset=[ARTICLE_COMPOSE_SOURCE_COL, ARTICLE_COMPOSE_TARGET_COL])
    mapping = {
        str(src).strip(): str(dst).strip()
        for src, dst in zip(df[ARTICLE_COMPOSE_SOURCE_COL], df[ARTICLE_COMPOSE_TARGET_COL])
    }
    return mapping


def apply_article_compose(blank_df, compose_map):
    """
    For rows whose FCL_ART_CODE matches a key in compose_map, replace
    it with the composed/TMS-recognized code before the TMS check.

    Returns the modified dataframe and the number of rows substituted.
    """
    raw_codes = blank_df["FCL_ART_CODE"].astype(str).str.strip()
    composed = raw_codes.map(compose_map)
    matched_mask = composed.notna()

    n_matched = int(matched_mask.sum())
    if n_matched:
        blank_df.loc[matched_mask, "FCL_ART_CODE"] = composed[matched_mask]

    return blank_df, n_matched


def add_article_check_conditional_formatting(ws, n_rows, n_cols, check_col_idx):
    """
    Color the ENTIRE ROW based on 'Article Code Check': green for
    Checked, red for Review - not just the check-column cell itself,
    so a flagged row is obvious at a glance across all its fields.
    Formula-based with an absolute column / relative row reference
    (same pattern as add_keep_conditional_formatting's Keep?-based row
    highlighting on A5B), so it stays correct if rows are re-sorted or
    edited later.
    """
    check_col_letter = get_column_letter(check_col_idx)
    last_col_letter = get_column_letter(n_cols)
    data_range = f"A2:{last_col_letter}{n_rows}"
    ref_cell = f"${check_col_letter}2"

    fill_checked = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    fill_review = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=[f'{ref_cell}="{CHECKED}"'], fill=fill_checked, stopIfTrue=True),
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=[f'{ref_cell}="{REVIEW}"'], fill=fill_review, stopIfTrue=True),
    )


def add_keep_conditional_formatting(ws, n_rows, n_cols, keep_col_idx):
    """
    Same rules as 03_clean_transform.py's add_keep_conditional_formatting,
    reproduced here so A5B (the excluded-rows workbook) keeps the exact
    same Keep?-based row coloring that A4 has. See that script for the
    rule rationale.
    """
    keep_col_letter = get_column_letter(keep_col_idx)
    last_col_letter = get_column_letter(n_cols)
    data_range = f"A2:{last_col_letter}{n_rows}"
    ref_cell = f"${keep_col_letter}2"

    fill_both = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_not_asset = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
    fill_not_deployed = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    fill_pending = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    rules = [
        FormulaRule(
            formula=[f'AND(ISNUMBER(SEARCH("{NOT_ASSET}",{ref_cell})),'
                     f'ISNUMBER(SEARCH("{NOT_DEPLOYED}",{ref_cell})))'],
            fill=fill_both, stopIfTrue=True,
        ),
        FormulaRule(
            formula=[f'ISNUMBER(SEARCH("{NOT_ASSET}",{ref_cell}))'],
            fill=fill_not_asset, stopIfTrue=True,
        ),
        FormulaRule(
            formula=[f'ISNUMBER(SEARCH("{NOT_DEPLOYED}",{ref_cell}))'],
            fill=fill_not_deployed, stopIfTrue=True,
        ),
        FormulaRule(
            formula=[f'{ref_cell}=""'],
            fill=fill_pending, stopIfTrue=True,
        ),
    ]
    for rule in rules:
        ws.conditional_formatting.add(data_range, rule)


def _write_table_workbook(df, out_path, sheet_title):
    """
    Shared workbook-writing mechanics (table + column widths + text
    formatting for lot/serial numbers) used by both A5 and A5B. Does
    NOT apply conditional formatting - callers add whichever rules fit
    their own sheet after this returns.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]

    headers = list(df.columns)
    ws.append(headers)
    for row in df.itertuples(index=False, name=None):
        ws.append(row)

    n_rows = len(df) + 1
    n_cols = len(headers)
    last_col_letter = get_column_letter(n_cols)
    table_ref = f"A1:{last_col_letter}{n_rows}"

    table = Table(displayName=f"Table_{sheet_title[:20].replace(' ', '_')}", ref=table_ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)

    if "PCL_NO_SERIE_LOT" in headers:
        col_idx = headers.index("PCL_NO_SERIE_LOT") + 1
        col_letter = get_column_letter(col_idx)
        for cell in ws[col_letter][1:]:
            cell.number_format = "@"

    for i, col_name in enumerate(headers, start=1):
        width = max(12, min(40, len(str(col_name)) + 2))
        ws.column_dimensions[get_column_letter(i)].width = width

    try:
        wb.save(out_path)
    except PermissionError:
        raise PermissionError(
            f"\nCould not save '{out_path}' - permission denied.\n"
            f"This almost always means a file with that name is currently open "
            f"in Excel - close it and re-run."
        )
    return wb, ws, headers, n_rows, n_cols


def save_a5_xlsx(df, work_dir, date_tag):
    a5_path = os.path.join(work_dir, f"A5_EXTRACTION_TMS_OCBA_{date_tag}.xlsx")
    wb, ws, headers, n_rows, n_cols = _write_table_workbook(
        df, a5_path, f"A5_EXTRACTION_TMS_OCBA_{date_tag}"
    )

    check_col_idx = headers.index("Article Code Check") + 1
    add_article_check_conditional_formatting(ws, n_rows, n_cols, check_col_idx)

    wb.save(a5_path)  # re-save with conditional formatting added
    return a5_path


def save_a5b_xlsx(df, output_dir, date_tag):
    a5b_path = os.path.join(output_dir, f"A5B_EXTRACTION_TMS_OCBA_{date_tag}.xlsx")
    wb, ws, headers, n_rows, n_cols = _write_table_workbook(
        df, a5b_path, f"A5B_EXTRACTION_TMS_OCBA_{date_tag}"
    )

    keep_col_idx = headers.index("Keep?") + 1
    add_keep_conditional_formatting(ws, n_rows, n_cols, keep_col_idx)

    wb.save(a5b_path)  # re-save with conditional formatting added
    return a5b_path


def resolve_output_dir(paths):
    """
    Destination for A5B (and later A6, written by 05_export.py) - the
    folder these get emailed from. Uses the configured output_dir from
    config.conf ([local] output_dir), with a subfolder named after the
    shared week_tag convention (e.g. "Y26W25" - the ISO week BEFORE the
    one this script is run in, per common_config.compute_week_tag()),
    NOT the batch's own date tag, so batches processed in the same
    week share a folder with 05_export.py's A6 output.
    """
    week_tag = compute_week_tag()
    output_dir = get_dated_subdir(paths["output_dir"], week_tag)
    return output_dir, week_tag


def main():
    cfg = load_config()
    paths = get_paths(cfg)

    a4_arg = sys.argv[1] if len(sys.argv) > 1 else None
    tms_arg = sys.argv[2] if len(sys.argv) > 2 else None
    compose_arg = sys.argv[3] if len(sys.argv) > 3 else None

    a4_path = a4_arg or find_latest_a4(paths["work_dir"])
    print(f"Reading A4 workbook: {a4_path}")

    tms_list_path = tms_arg or cfg.get(
        "templates", "tms_article_list_path", fallback=TMS_ARTICLE_LIST_DEFAULT
    )
    article_compose_path = compose_arg or cfg.get(
        "templates", "article_compose_path", fallback=ARTICLE_COMPOSE_DEFAULT
    )

    work_dir_for_output = os.path.dirname(a4_path)
    base = os.path.basename(a4_path)
    date_tag = base.replace("A4_EXTRACTION_TMS_OCBA_", "").replace(".xlsx", "")

    try:
        df = pd.read_excel(a4_path, dtype=str, keep_default_na=True)
    except PermissionError:
        raise PermissionError(
            f"\nCould not open '{a4_path}' - permission denied.\n"
            f"This almost always means the file is currently open in Excel "
            f"(Windows locks it while open) - close it and try again. If it's "
            f"not open, it may also be mid-sync via OneDrive/SharePoint; wait a "
            f"moment and retry."
        )

    if "Keep?" not in df.columns:
        raise KeyError("A4 workbook has no 'Keep?' column - check it was produced by 03_clean_transform.py.")
    if "FCL_ART_CODE" not in df.columns:
        raise KeyError("A4 workbook has no 'FCL_ART_CODE' column.")

    keep_blank_mask = df["Keep?"].isna() | (df["Keep?"].astype(str).str.strip() == "")
    blank_df = df[keep_blank_mask].copy()
    excluded_df = df[~keep_blank_mask].copy()
    print(f"A4 rows: {len(df)} total | {len(blank_df)} with blank Keep? carried forward to A5 | "
          f"{len(excluded_df)} excluded (non-blank Keep?) carried to A5B")

    # --- Article compose substitution (hidden Kits without leading 'K') ---
    compose_map = load_article_compose_map(article_compose_path)
    print(f"Loaded {len(compose_map)} article-compose mapping(s) from: {article_compose_path}")
    blank_df, n_composed = apply_article_compose(blank_df, compose_map)
    if n_composed:
        print(f"  Substituted {n_composed} row(s)' FCL_ART_CODE with its ArticleCompose "
              f"value before the TMS check.")

    # --- TMS article code check ---
    tms_codes = load_tms_article_codes(tms_list_path)
    print(f"Loaded {len(tms_codes)} known article code(s) from: {tms_list_path}")

    article_codes = blank_df["FCL_ART_CODE"].astype(str).str.strip()
    blank_df["Article Code Check"] = article_codes.apply(
        lambda code: CHECKED if code in tms_codes else REVIEW
    )

    n_checked = (blank_df["Article Code Check"] == CHECKED).sum()
    n_review = (blank_df["Article Code Check"] == REVIEW).sum()
    print(f"Article Code Check: {n_checked} Checked | {n_review} Review (not found in TMS list)")

    a5_path = save_a5_xlsx(blank_df, work_dir_for_output, date_tag)
    print(f"\nA5 workbook saved to: {a5_path}")

    output_dir, week_tag = resolve_output_dir(paths)
    a5b_path = save_a5b_xlsx(excluded_df, output_dir, date_tag)
    print(f"A5B workbook (excluded rows) saved to: {a5b_path}  (output folder: {week_tag})")

    print("\nNext step: pending further guidance on processing A5 / assembling A6.")


if __name__ == "__main__":
    main()
