import os

import shutil

import datetime

import pandas as pd

from pathlib import Path

import unicodedata

import re

from collections import defaultdict

# ============================

# CONFIG – EDIT THESE IF NEEDED

# ============================

INPUT_PATH = r"C:\Users\jgonzale\OneDrive - Cuscal\Desktop\move_name\List_Jul18.xlsx"  # <-- .csv, .xlsx, or .xls

DRY_RUN = False  # True = simulate; False = actually move

# Column names (case-insensitive). We'll try to find these by fuzzy matching.

COL_SURNAME_CANDIDATES = {"surname", "last name", "last_name", "family name", "family_name", "lastname"}

COL_GIVEN_CANDIDATES   = {"given names", "given name", "first names", "first name", "first_name", "first", "firstname"}

COL_PREF_CANDIDATES    = {"preferred name", "preferred", "preferred_name", "preferred given name", "nickname"}

# Source folders to search (top-level only)

FOLDERS_TO_SEARCH = [


     r"C:\Folder to search path 1",

     r"C:\Folder to search path 2",

     r"C:\Folder to search path 3",

     r"C:\Folder to search path 4",

]

# Destination folder

DESTINATION_FOLDER = r"C:\Destination Folder"

# ============================

# NORMALIZATION & HELPERS

# ============================

def strip_accents(s):

    if not isinstance(s, str):

        return ""

    nfkd = unicodedata.normalize("NFKD", s)

    return "".join(c for c in nfkd if not unicodedata.combining(c))

_PUNCT_RE = re.compile(r"[,\.\-’'`()]")  # punctuation to ignore for extra-flex matching

def canonical_spaces(s):

    return " ".join(s.split()).strip()

def norm_key_variants(s):

    """Return multiple normalized keys for robust matching."""

    if not isinstance(s, str) or not s.strip():

        return set()

    base = canonical_spaces(s).lower()

    no_acc = canonical_spaces(strip_accents(s)).lower()

    no_punct = canonical_spaces(_PUNCT_RE.sub("", base))

    no_punct_no_acc = canonical_spaces(_PUNCT_RE.sub("", no_acc))

    return {base, no_acc, no_punct, no_punct_no_acc}

def ensure_unique_destination(dest_dir: Path, folder_name: str) -> Path:

    candidate = dest_dir / folder_name

    if not candidate.exists():

        return candidate

    counter = 1

    while True:

        candidate2 = dest_dir / f"{folder_name}_{counter}"

        if not candidate2.exists():

            return candidate2

        counter += 1

def collect_top_level_dirs(parent: Path):

    try:

        with os.scandir(parent) as it:

            for entry in it:

                if entry.is_dir(follow_symlinks=False):

                    yield Path(entry.path)

    except (FileNotFoundError, PermissionError):

        return

def find_column(df: pd.DataFrame, candidates: set) -> str | None:

    """Find a column whose lowercase name matches any candidate (spaces/underscores ignored)."""

    lower_map = {c.lower(): c for c in df.columns}

    for cand in candidates:

        if cand in lower_map:

            return lower_map[cand]

    normalized = {c.lower().replace(" ", "").replace("_", ""): c for c in df.columns}

    for cand in candidates:

        key = cand.replace(" ", "").replace("_", "")

        if key in normalized:

            return normalized[key]

    return None

def safe_val(v) -> str:

    return "" if pd.isna(v) else str(v).strip()

def generate_name_candidates(surname: str, given: str, preferred: str) -> set:

    """Generate all possible name variants for matching (standalone + combos)."""

    cands = set()

    s = safe_val(surname)

    g = safe_val(given)

    p = safe_val(preferred)

    # Standalone

    if p: cands.add(p)

    if g: cands.add(g)

    if s: cands.add(s)

    # Given + Surname combos

    if g and s:

        cands.add(f"{g} {s}")      # Given Surname

        cands.add(f"{s} {g}")      # Surname Given

        cands.add(f"{s}, {g}")     # Surname, Given

    # Preferred + Surname combos

    if p and s:

        cands.add(f"{p} {s}")      # Preferred Surname

        cands.add(f"{s} {p}")      # Surname Preferred

        cands.add(f"{s}, {p}")     # Surname, Preferred

    return {canonical_spaces(x) for x in cands if x}

def read_table(path: str) -> pd.DataFrame:

    ext = Path(path).suffix.lower()

    if ext == ".csv":

        return pd.read_csv(path)

    elif ext in (".xlsx", ".xls"):

        return pd.read_excel(path)

    else:

        raise ValueError("Unsupported input file type. Use .csv, .xlsx, or .xls")

# ============================

# MAIN

# ============================

def main():

    ts_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    dest_dir = Path(DESTINATION_FOLDER)

    dest_dir.mkdir(parents=True, exist_ok=True)

    # Read input

    try:

        df = read_table(INPUT_PATH)

    except Exception as e:

        print(f"❌ Could not read input at {INPUT_PATH}: {e}")

        return

    if df.empty:

        print("ℹ️ Input file is empty.")

        return

    # Resolve columns (capture the EXACT header names to reuse in "Not Found" sheet)

    cs = find_column(df, COL_SURNAME_CANDIDATES)

    cg = find_column(df, COL_GIVEN_CANDIDATES)

    cp = find_column(df, COL_PREF_CANDIDATES)

    if not any([cs, cg, cp]):

        print("❌ Could not find surname/given/preferred columns.")

        return

    # Build candidate keys and also track per-input-row for "Not Found"

    # We'll record which input rows got at least one match.

    row_matched = [False] * len(df)

    # Precompute all candidate keys per row for efficient matching

    row_keys: list[set[str]] = []

    for idx, row in df.iterrows():

        s = row[cs] if cs else ""

        g = row[cg] if cg else ""

        p = row[cp] if cp else ""

        keys = set()

        for cand in generate_name_candidates(s, g, p):

            keys |= norm_key_variants(cand)

        row_keys.append({k for k in keys if k})

    # Index source folders across all search roots

    index = {}

    total_indexed = 0

    for root in FOLDERS_TO_SEARCH:

        root_path = Path(root)

        if not root_path.exists():

            print(f"⚠️ Source path does not exist (skipping): {root}")

            continue

        for subdir in collect_top_level_dirs(root_path):

            for k in norm_key_variants(subdir.name):

                if not k:

                    continue

                index.setdefault(k, []).append(subdir)

            total_indexed += 1

    print(f"📁 Indexed {total_indexed} top-level folders from sources.")

    # Build folder -> all matching keys (so we move/log each folder ONCE)

    folder_to_keys = defaultdict(set)

    # Also remember which keys matched so we can set row_matched flags

    for idx, keys in enumerate(row_keys):

        matched_this_row = False

        for k in keys:

            for f in index.get(k, []):

                folder_to_keys[f].add(k)

                matched_this_row = True

        row_matched[idx] = matched_this_row

    # Build Matches log (one row per unique folder)

    log_rows = []

    moved_count = planned_count = error_count = 0

    for src_folder, keys in sorted(folder_to_keys.items(), key=lambda x: str(x[0]).lower()):

        unique_dest = ensure_unique_destination(dest_dir, src_folder.name)

        if DRY_RUN:

            action, result = "DRY-RUN (move planned)", "Simulated"

            planned_count += 1

        else:

            try:

                shutil.move(str(src_folder), str(unique_dest))

                action, result = "MOVE", "Success"

                moved_count += 1

            except Exception as e:

                action, result = "MOVE", f"ERROR: {e}"

                error_count += 1

        log_rows.append({

            "Matched Variants": " / ".join(sorted(keys)),

            "Folder Name": src_folder.name,

            "Source Folder": str(src_folder),

            "Destination Folder": str(unique_dest),

            "Action": action,

            "Result": result,

            "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        })

        print(f"✅ {action}: {src_folder} -> {unique_dest} ({result})")

    # Build "Not Found" sheet using the ORIGINAL columns as they appear in the input

    # Only include the columns that actually exist in the file (and in the same header names).

    not_found_cols = [c for c in [cs, cg, cp] if c]  # keep original header names/order

    not_found_records = []

    for i, matched in enumerate(row_matched):

        if not matched:

            rec = {}

            for col in not_found_cols:

                rec[col] = df.iloc[i][col]

            not_found_records.append(rec)

    # Save Excel with 2 sheets: Matches + Not Found

    matches_df = pd.DataFrame(log_rows)

    not_found_df = pd.DataFrame(not_found_records)

    out_xlsx = Path(DESTINATION_FOLDER) / f"move_name_folders_log_{ts_str}.xlsx"

    try:

        with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:

            # If there were no matches, still create an empty sheet with headers

            if not matches_df.empty:

                matches_df.to_excel(writer, sheet_name="Matches", index=False)

            else:

                pd.DataFrame(columns=[

                    "Matched Variants", "Folder Name", "Source Folder",

                    "Destination Folder", "Action", "Result", "Timestamp"

                ]).to_excel(writer, sheet_name="Matches", index=False)

            # "Not Found" shows rows as they appeared in the original file (selected columns only)

            if not not_found_df.empty:

                not_found_df.to_excel(writer, sheet_name="Not Found", index=False)

            else:

                # still create the sheet with appropriate headers if none missing

                pd.DataFrame(columns=not_found_cols).to_excel(writer, sheet_name="Not Found", index=False)

        print(f"\n📄 Log saved: {out_xlsx}")

    except Exception as e:

        print(f"❌ Failed to save Excel log: {e}")

    # Summary

    print(f"\n📊 Summary: planned={planned_count}, moved={moved_count}, errors={error_count}, "

          f"matched_folders={len(log_rows)}, not_found_rows={len(not_found_records)}")

if __name__ == "__main__":

    main()

 
