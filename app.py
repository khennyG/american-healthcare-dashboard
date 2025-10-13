import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

# -------------------------------
# CONFIG
# -------------------------------
st.set_page_config(page_title="American Healthcare Class Dashboard", layout="wide")

# Week info mapping
week_info = {
    "Week 1 (9/4)": "U.S. Healthcare Ecosystem Overview and History",
    "Week 2 (9/11)": "Patient: Experience, Responsibility, and Paying for Healthcare",
    "Week 3 (9/18)": "Patient: Care Delivery – Where, Why, Impact and Challenges",
    "Week 4 (9/25)": "Payer: History, Overview, ACA and Services",
    "Week 5 (10/2)": "Payer: Medicare and Medicaid",
    "Week 6 (10/9)": "Provider: History, Types and Challenges"
}

# -------------------------------
# LOAD DATA
# -------------------------------
@st.cache_data
def load_data():
    """Load the participation Excel file and normalize headers.

    - Searches for the file in project root and in this file's folder
    - Detects the correct header row (some sheets have headers on row 2)
    - Normalizes column names (strip + collapse whitespace)
    - Ensures the first column is named "Student"
    - Validates that at least one Week column exists
    """
    here = Path(__file__).resolve().parent
    # Search for likely Excel files in project root and sub-app dir
    candidates = []
    for base in [here, here / "american_healthcare_dashboard"]:
        candidates.extend(sorted(base.glob("*.xlsx")))
    # Prefer files that look like the class participation sheet
    def score(p: Path) -> int:
        name = p.name.lower()
        s = 0
        if "participation" in name:
            s += 3
        if "healthcare" in name or "health care" in name:
            s += 2
        if "class" in name:
            s += 1
        return s
    candidates = sorted(set(candidates), key=lambda p: (score(p), -p.stat().st_mtime), reverse=True)
    excel_path = candidates[0] if candidates else None
    if excel_path is None:
        st.error(
            "Couldn't find 'American HealthCare Class Participation.xlsx'. "
            "Place it in the project root or alongside the app file."
        )
        st.stop()

    # Try a few header rows to find "Week" columns
    chosen = None
    for hdr in range(0, 4):
        try:
            tmp = pd.read_excel(excel_path, header=hdr)
        except Exception:
            continue
        tmp.columns = [" ".join(str(c).split()) for c in tmp.columns]
        if any("Week" in str(c) for c in tmp.columns):
            chosen = tmp
            break
    df = chosen if chosen is not None else pd.read_excel(excel_path)

    # Normalize column names (trim + collapse spaces)
    df.columns = [" ".join(str(c).split()) for c in df.columns]

    # Rename first column to "Student" and clean values
    first_col = df.columns[0]
    df = df.rename(columns={first_col: "Student"})
    df = df.dropna(subset=["Student"])  # drop blank rows
    df["Student"] = df["Student"].astype(str).str.strip()

    # Validate presence of Week columns
    week_cols = [c for c in df.columns if "Week" in str(c)]
    if not week_cols:
        st.error(
            "No 'Week' columns were detected in the Excel file. "
            "Please ensure the header row is set correctly and the columns contain 'Week'."
        )
        st.stop()

    return df

df = load_data()

# -------------------------------
# CLEAN & EXTRACT PARTICIPATION / ATTENDANCE
# -------------------------------
# Detect participation (# count)
def count_hashes(value):
    if pd.isna(value):
        return 0
    return value.count("#")

# Detect attendance
def attendance_status(value):
    if pd.isna(value):
        return "Absent"
    if "$" in str(value):
        return "Excused"
    if "%" in str(value):
        return "Absent"
    if "*" in str(value):
        return "Present"
    return "Present" if "#" in str(value) else "Absent"

# Extract relevant columns (Weeks only)
week_cols = [c for c in df.columns if "Week" in str(c)]
participation_data = []

for _, row in df.iterrows():
    student = row["Student"]
    for week in week_cols:
        value = str(row[week])
        participation = count_hashes(value)
        attendance = attendance_status(value)
        topic = week_info.get(week.strip(), "N/A")
        participation_data.append({
            "Student": student,
            "Week": week,
            "Topic": topic,
            "Participation": participation,
            "Attendance": attendance
        })

df_long = pd.DataFrame(participation_data)

# -------------------------------
# DASHBOARD
# -------------------------------
st.title("🏥 American Healthcare Class – Participation & Attendance Dashboard")
st.markdown("### Midterm Overview (Weeks 1–6)")

# Sidebar Filters
st.sidebar.header("Filter Options")
view_mode = st.sidebar.radio("View Mode", ["Overview", "By Student", "By Week"])

# -------------------------------
# OVERVIEW MODE
# -------------------------------
if view_mode == "Overview":
    st.subheader("📊 Class Participation Overview")

    # Total participation leaderboard
    total_participation = (
        df_long.groupby("Student")["Participation"].sum().reset_index().sort_values(by="Participation", ascending=False)
    )
    fig1 = px.bar(
        total_participation,
        x="Participation",
        y="Student",
        orientation="h",
        title="Total Participation Leaderboard",
        color="Participation",
        color_continuous_scale="Tealgrn"
    )
    st.plotly_chart(fig1, use_container_width=True)

    # Average participation per week
    avg_week = df_long.groupby("Week")["Participation"].mean().reset_index()
    fig2 = px.bar(
        avg_week,
        x="Week",
        y="Participation",
        title="Average Participation per Week",
        color="Participation",
        color_continuous_scale="Blues"
    )
    st.plotly_chart(fig2, use_container_width=True)

    # Attendance Summary
    st.subheader("🧾 Attendance Summary")
    attendance_counts = df_long.groupby(["Attendance"]).size().reset_index(name="Count")
    fig3 = px.pie(attendance_counts, names="Attendance", values="Count", title="Overall Attendance Breakdown")
    st.plotly_chart(fig3, use_container_width=True)

# -------------------------------
# BY STUDENT MODE
# -------------------------------
elif view_mode == "By Student":
    students = sorted(df_long["Student"].unique())
    student_choice = st.sidebar.selectbox("Select a student", students)
    student_df = df_long[df_long["Student"] == student_choice]

    st.subheader(f"👩‍🎓 Participation Trend – {student_choice}")
    fig4 = px.line(
        student_df,
        x="Week",
        y="Participation",
        markers=True,
        title=f"Weekly Participation Trend for {student_choice}",
        text="Participation"
    )
    st.plotly_chart(fig4, use_container_width=True)

    # Attendance pie for this student
    att = student_df["Attendance"].value_counts().reset_index()
    att.columns = ["Status", "Count"]
    fig5 = px.pie(att, names="Status", values="Count", title=f"Attendance Breakdown for {student_choice}")
    st.plotly_chart(fig5, use_container_width=True)

    total = student_df["Participation"].sum()
    avg = student_df["Participation"].mean()
    weeks_spoken = (student_df["Participation"] > 0).sum()
    st.markdown(f"**Total Participation:** {total}  |  **Weekly Average:** {avg:.2f}  |  **Spoke in:** {weeks_spoken} / 6 weeks")

# -------------------------------
# BY WEEK MODE
# -------------------------------
else:
    weeks = sorted(df_long["Week"].unique())
    week_choice = st.sidebar.selectbox("Select a week", weeks)
    week_df = df_long[df_long["Week"] == week_choice]
    st.subheader(f"📅 {week_choice} – {week_info.get(week_choice, '')}")

    fig6 = px.bar(
        week_df,
        x="Participation",
        y="Student",
        orientation="h",
        color="Participation",
        color_continuous_scale="Viridis",
        title=f"Participation by Student – {week_choice}"
    )
    st.plotly_chart(fig6, use_container_width=True)

    att_week = week_df["Attendance"].value_counts().reset_index()
    att_week.columns = ["Status", "Count"]
    fig7 = px.pie(att_week, names="Status", values="Count", title=f"Attendance Breakdown – {week_choice}")
    st.plotly_chart(fig7, use_container_width=True)
