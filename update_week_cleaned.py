import argparse
import pandas as pd
from pathlib import Path
from typing import Optional, List

###############################
# CONFIG: set these weekly
###############################
WEEK_NUMBER: int = 7
WEEK_TOPIC: str = "Provider: Compensation, Education, Burn-out and Satisfaction"
# Optional explicit date to include in Week label like "9/11"; if None, we try to parse from header
WEEK_DATE: Optional[str] = None

# Filenames
RAW_CANDIDATES: List[Path] = [
    Path("American HealthCare Class Participation.xlsx"),
    Path("American HealthCare Class.xlsx"),
]
CLEANED_FILE = Path("American_Healthcare_Class_Cleaned.xlsx")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Append or update a week's participation+attendance from raw to cleaned Excel."
    )
    parser.add_argument(
        "--week", type=int, help="Week number to append/update (e.g., 7)"
    )
    parser.add_argument(
        "--topic", type=str, help="Topic/title for the week to store in cleaned file"
    )
    parser.add_argument(
        "--date", type=str, default=None, help="Optional date like MM/DD (e.g., 10/16). If omitted, will try parsing from header."
    )
    parser.add_argument(
        "--raw",
        type=str,
        action="append",
        help="Path to raw participation Excel. Can be passed multiple times to specify a search order.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Do not write the cleaned file; print a short summary instead.",
    )
    return parser.parse_args()


def find_existing(paths: List[Path]) -> Optional[Path]:
    for p in paths:
        if p.exists():
            return p
    return None


def detect_header_row(raw_path: Path) -> int:
    """Return the row index to use as header (0-based)."""
    tmp = pd.read_excel(raw_path, header=None)
    for i in range(min(20, len(tmp))):
        row_has_week = tmp.iloc[i].astype(str).str.contains("Week", case=False, na=False)
        if row_has_week.any():
            return i
    return 0


def count_hashes(value) -> int:
    if pd.isna(value):
        return 0
    return str(value).count("#")


def attendance_status(value) -> str:
    """Map raw cell symbols to an attendance label.

    Supported symbols/markers:
    - Present: contains '*', '✓', '✔', or '#'; or equals 'P'/'p' or contains 'PRESENT'
    - Excused: contains '$'; or equals 'E'/'e' or contains 'EXCUSED'
    - Absent: contains '%'; or equals 'A'/'a' or 'X'/'x' or contains 'ABSENT'

    Precedence: Excused > Present > Absent. Empty/NaN -> Absent.
    """
    if pd.isna(value):
        return "Absent"
    s = str(value)
    s_upper = s.strip().upper()

    # Excused first
    if "$" in s or s_upper == "E" or "EXCUSED" in s_upper:
        return "Excused"

    # Present if explicit symbols or participation marks present
    if any(sym in s for sym in ("*", "✓", "✔")) or "#" in s or s_upper == "P" or "PRESENT" in s_upper:
        return "Present"

    # Absent symbols/markers
    if "%" in s or s_upper in {"A", "X"} or "ABSENT" in s_upper:
        return "Absent"

    # Default to Absent if unknown marker
    return "Absent"


def extract_date_from_header(header: str) -> Optional[str]:
    if not isinstance(header, str):
        return None
    if "(" in header and ")" in header and header.find("(") < header.find(")"):
        return header.split("(")[-1].split(")")[0].strip()
    return None


def main(dry_run: bool = False):
    raw_path = find_existing(RAW_CANDIDATES)
    if not raw_path:
        raise FileNotFoundError(
            f"Could not find raw Excel file in: {[str(p) for p in RAW_CANDIDATES]}"
        )

    header_idx = detect_header_row(raw_path)
    df = pd.read_excel(raw_path, header=header_idx)

    # Standardize first column as Student
    first_col = df.columns[0]
    df = df.rename(columns={first_col: "Student"})
    df = df.dropna(subset=["Student"]).copy()
    df["Student"] = df["Student"].astype(str).str.strip()

    # Find the Week N column (robustly match the numeric after 'Week')
    import re
    def _is_target_week(col) -> bool:
        s = str(col)
        s_norm = " ".join(s.split()).lower()
        if "week" not in s_norm:
            return False
        m = re.search(r"week\s*(\d+)", s_norm, flags=re.I)
        if m:
            try:
                return int(m.group(1)) == WEEK_NUMBER
            except ValueError:
                return False
        # fallback exact substring
        return f"week {WEEK_NUMBER}" in s_norm

    week_col = None
    for c in df.columns:
        if _is_target_week(c):
            week_col = c
            break
    if week_col is None:
        candidate_weeks = [str(c) for c in df.columns if "week" in str(c).lower()]
        raise ValueError(
            f"Could not find a column for Week {WEEK_NUMBER}. Found week-like columns: {candidate_weeks}"
        )

    # Build Week label and Topic
    header_str = str(week_col)
    date_str = WEEK_DATE or extract_date_from_header(header_str)
    if date_str:
        week_label = f"Week {WEEK_NUMBER} ({date_str})"
    else:
        week_label = f"Week {WEEK_NUMBER}"

    # Build new rows
    new_records = []
    for _, row in df.iterrows():
        val = row.get(week_col, None)
        new_records.append({
            "Student": row["Student"],
            "Week": week_label,
            "Date": date_str,
            "Topic": WEEK_TOPIC,
            "Participation": int(count_hashes(val)),
            "Attendance": attendance_status(val),
        })

    new_df = pd.DataFrame(new_records)

    if dry_run:
        print(f"[DRY-RUN] Would append {len(new_df)} rows for Week {WEEK_NUMBER} (label '{week_label}')")
        print("Sample:\n", new_df.head().to_string(index=False))
        return

    # If existing cleaned file present, append & dedupe; else create new
    if CLEANED_FILE.exists():
        cleaned_df = pd.read_excel(CLEANED_FILE)
        # Keep a stable column order: existing columns first
        # Ensure required columns exist
        for col in ["Student", "Week", "Topic", "Participation", "Attendance"]:
            if col not in new_df.columns:
                new_df[col] = pd.NA

        # Align columns (union)
        all_cols = list(dict.fromkeys(list(cleaned_df.columns) + list(new_df.columns)))
        cleaned_df = cleaned_df.reindex(columns=all_cols)
        new_df = new_df.reindex(columns=all_cols)

        updated = pd.concat([cleaned_df, new_df], ignore_index=True)
        # Deduplicate by Student+Week, keep last (so re-runs update)
        if "Student" in updated.columns and "Week" in updated.columns:
            updated.sort_values(by=["Student", "Week"], inplace=True)
            updated = updated.drop_duplicates(subset=["Student", "Week"], keep="last").reset_index(drop=True)
        updated.to_excel(CLEANED_FILE, index=False)
        print(f"✅ Appended Week {WEEK_NUMBER} to cleaned file: {CLEANED_FILE}")
    else:
        # No existing cleaned file: create with new week only
        new_df.to_excel(CLEANED_FILE, index=False)
        print(f"✅ Created cleaned file with Week {WEEK_NUMBER}: {CLEANED_FILE}")


if __name__ == "__main__":
    args = parse_args()

    # Override config from CLI if provided
    if args.week is not None:
        WEEK_NUMBER = args.week
    if args.topic is not None:
        WEEK_TOPIC = args.topic
    # Allow explicit empty string to clear date
    if args.date is not None:
        WEEK_DATE = args.date or None
    if args.raw:
        RAW_CANDIDATES = [Path(p) for p in args.raw]

    main(dry_run=bool(args.dry_run))
