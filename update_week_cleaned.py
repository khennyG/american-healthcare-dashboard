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
    if pd.isna(value):
        return "Absent"
    s = str(value)
    if "$" in s:
        return "Excused"
    if "%" in s:
        return "Absent"
    if "*" in s:
        return "Present"
    return "Present" if "#" in s else "Absent"


def extract_date_from_header(header: str) -> Optional[str]:
    if not isinstance(header, str):
        return None
    if "(" in header and ")" in header and header.find("(") < header.find(")"):
        return header.split("(")[-1].split(")")[0].strip()
    return None


def main():
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

    # Find the Week N column
    week_col = None
    for c in df.columns:
        if "week" in str(c).lower() and str(WEEK_NUMBER) in str(c):
            week_col = c
            break
    if week_col is None:
        raise ValueError(f"Could not find a column for Week {WEEK_NUMBER} in raw file headers: {list(df.columns)}")

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
    main()
