import pandas as pd
from pathlib import Path

# -----------------------------
# CONFIG
# -----------------------------
input_file = Path("American HealthCare Class.xlsx")
output_file = Path("American_Healthcare_Class_Cleaned.xlsx")
backup_file = Path("American_Healthcare_Class_Backup.xlsx")

# -----------------------------
# STEP 1 – Load and Backup
# -----------------------------
df_raw = pd.read_excel(input_file, header=None)
df_raw.to_excel(backup_file, index=False)  # backup original sheet

# Find the header row (the one that contains "Week 1")
header_row_idx = None
for i in range(len(df_raw)):
    row = df_raw.iloc[i].astype(str).str.contains("Week", case=False, na=False)
    if row.any():
        header_row_idx = i
        break

# Read again with correct header
df = pd.read_excel(input_file, header=header_row_idx)
df = df.rename(columns={df.columns[0]: "Student"})
df = df.dropna(subset=["Student"])

# -----------------------------
# STEP 2 – Standardize Columns
# -----------------------------
week_names = {
    0: "Week 1 (9/4)",
    1: "Week 2 (9/11)",
    2: "Week 3 (9/18)",
    3: "Week 4 (9/25)",
    4: "Week 5 (10/2)",
    5: "Week 6 (10/9)"
}

week_cols = [c for c in df.columns if "week" in str(c).lower()]
for i, c in enumerate(week_cols):
    if i < len(week_names):
        df.rename(columns={c: week_names[i]}, inplace=True)

# -----------------------------
# STEP 3 – Extract Participation and Attendance
# -----------------------------
def count_hashes(value):
    if pd.isna(value):
        return 0
    return str(value).count("#")

def attendance_status(value):
    if pd.isna(value):
        return "Absent"
    s = str(value)
    s_upper = s.strip().upper()
    if "$" in s or s_upper == "E" or "EXCUSED" in s_upper:
        return "Excused"
    if any(sym in s for sym in ("*", "✓", "✔")) or "#" in s or s_upper == "P" or "PRESENT" in s_upper:
        return "Present"
    if "%" in s or s_upper in {"A", "X"} or "ABSENT" in s_upper:
        return "Absent"
    return "Absent"

# Week details
week_info = {
    "Week 1 (9/4)": "U.S. Healthcare Ecosystem Overview and History",
    "Week 2 (9/11)": "Patient: Experience, Responsibility, and Paying for Healthcare",
    "Week 3 (9/18)": "Patient: Care Delivery – Where, Why, Impact and Challenges",
    "Week 4 (9/25)": "Payer: History, Overview, ACA and Services",
    "Week 5 (10/2)": "Payer: Medicare and Medicaid",
    "Week 6 (10/9)": "Provider: History, Types and Challenges"
}

# -----------------------------
# STEP 4 – Reshape to Long Format
# -----------------------------
records = []
for _, row in df.iterrows():
    student = str(row["Student"]).strip()
    for week, topic in week_info.items():
        if week in df.columns:
            val = str(row[week])
            participation = count_hashes(val)
            attendance = attendance_status(val)
            records.append({
                "Student": student,
                "Week": week,
                "Date": week.split("(")[-1].strip(")"),
                "Topic": topic,
                "Participation": participation,
                "Attendance": attendance
            })

clean_df = pd.DataFrame(records)

# -----------------------------
# STEP 5 – Save Cleaned Data
# -----------------------------
clean_df.to_excel(output_file, index=False)
print(f"✅ Clean file created successfully at: {output_file.resolve()}")
