import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import re
from io import BytesIO

# -------------------------------
# CONFIG
# -------------------------------
st.set_page_config(page_title="American Healthcare Dashboard", layout="wide")

# -------------------------------
# THEME: Custom fonts and colors
# -------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap');

    :root {
        --neu-red: #CC0000;
        --neu-red-dark: #990000;
        --neu-black: #111111;
        --neu-gray-50: #FAFAFA;
        --neu-gray-200: #E5E7EB;
    }

    html, body, [class*="css"], .stApp {
        font-family: 'Poppins', sans-serif;
        background-color: #FFFFFF;
        color: var(--neu-black);
    }

    h1, h2, h3, h4, h5 {
        color: var(--neu-red) !important;
        font-weight: 600 !important;
    }

    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #FFF5F5 !important; /* light red tint */
        border-right: 1px solid var(--neu-gray-200);
    }

    /* Radio/Select labels */
    .stRadio label, .stSelectbox label, .stSlider label {
        font-weight: 500;
        color: var(--neu-black);
    }

    /* Buttons */
    .stButton>button {
        background-color: var(--neu-red);
        color: #FFFFFF;
        border: 0;
        border-radius: 6px;
        padding: 0.5rem 0.9rem;
    }
    .stButton>button:hover {
        background-color: var(--neu-red-dark);
        color: #FFFFFF;
    }

    /* Plotly axes titles */
    .plotly .xtitle, .plotly .ytitle {
        font-family: 'Poppins', sans-serif !important;
        fill: var(--neu-red) !important;
    }
    /* Plotly main chart title */
    .plotly .gtitle {
        fill: var(--neu-red) !important;
    }

    /* Notes block / blockquote */
    blockquote {
        border-left: 4px solid var(--neu-red);
        background-color: #FFF5F5;
        padding: 10px 15px;
        border-radius: 6px;
        margin: 0 0 1rem 0;
    }
    /* Metric label styling for By Student section */
    .metric-label {
        color: var(--neu-red);
        font-weight: 600;
    }

    /* Top-left brand badge */
    .brand-badge {
        position: fixed;
        top: 8px;
        left: 16px;
        display: flex;
        align-items: center;
        gap: 10px;
        z-index: 9999;
        pointer-events: none; /* don't block clicks */
    }
    .brand-logo {
        width: 28px;
        height: 28px;
        border-radius: 50%;
        background: var(--neu-red);
        color: #FFFFFF;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 14px;
        letter-spacing: 0.5px;
    }
    .brand-text {
        font-weight: 600;
        color: var(--neu-black);
        font-size: 14px;
    }

    /* Sidebar brand */
    .sidebar-brand {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 8px;
        padding-bottom: 12px;
        border-bottom: 1px solid var(--neu-gray-200);
        margin-bottom: 8px;
    }
    [data-testid="stSidebar"] .sidebar-brand .brand-logo {
        width: 24px;
        height: 24px;
        font-size: 12px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Plotly defaults: font + transparent backgrounds for seamless theme
pio.templates["ahcd"] = go.layout.Template(
    layout=go.Layout(
        font=dict(family="Poppins, sans-serif", color="#111111"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        title=dict(font=dict(color="#CC0000", size=18)),
        colorway=["#CC0000", "#990000", "#FF4D4D", "#7F1D1D", "#9CA3AF"],
    )
)
pio.templates.default = "ahcd+plotly_white"

# NU brand badges: page top-left and sidebar
st.markdown(
    """
    <div class="brand-badge">
        <div class="brand-logo">NU</div>
        <div class="brand-text">Northeastern University</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.sidebar.markdown(
    """
    <div class="sidebar-brand">
        <div class="brand-logo">NU</div>
        <div class="brand-text">Northeastern University</div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.title("HINF 5105 American Healthcare – Participation & Attendance Dashboard")
st.markdown("This dashboard provides a weekly breakdown of student participation and attendance.")

# -------------------------------
# LOAD DATA
# -------------------------------
@st.cache_data(ttl=30, show_spinner=False)
def load_data():
    """Load cleaned participation data from Excel file."""
    df = pd.read_excel("American_Healthcare_Class_Cleaned.xlsx")
    df["Participation"] = df["Participation"].fillna(0).astype(int)
    return df

df = load_data()

# -------------------------------
# SCORING
# -------------------------------
def calculate_score(student_df: pd.DataFrame, weeks_total: int | None = None) -> float:
    """Calculate a fair participation score out of 100.

    Rules:
    - Full credit for a week if student speaks at least once in that week.
    - Base score is proportion of weeks spoken times 100.
    - Bonus for extra speeches beyond first per week: 2 points each, capped at 10.
    - Final score capped at 100.
    """
    if student_df.empty:
        return 0.0
    if weeks_total is None:
        weeks_total = int(student_df["Week"].nunique()) or 0
    if weeks_total == 0:
        return 0.0

    spoke_weeks = int((student_df["Participation"] > 0).sum())
    total_speaks = int(student_df["Participation"].sum())

    base_score = (spoke_weeks / weeks_total) * 100.0
    bonus = max(0, total_speaks - spoke_weeks) * 2  # 2 pts per extra speak
    bonus = min(bonus, 10)  # cap bonus at 10

    return round(min(base_score + bonus, 100.0), 1)

# -------------------------------
# PDF GENERATION (Per-student report)
# -------------------------------
def generate_student_pdf(
    student_name: str,
    fig_line: go.Figure,
    fig_pie: go.Figure,
    metrics: dict,
    attendance_detail: dict,
) -> bytes:
    """Build a PDF report for a student with charts and summary.

    Requires 'kaleido' for Plotly image export and 'reportlab' for PDF.
    """
    # Lazy imports so the app can run even if not installed locally
    try:
        from reportlab.lib.pagesizes import letter  # type: ignore
        from reportlab.pdfgen import canvas  # type: ignore
        from reportlab.lib.utils import ImageReader  # type: ignore
        from reportlab.lib.colors import black  # type: ignore
    except Exception as e:
        # Provide a friendly error as PDF bytes placeholder
        return f"Missing PDF dependencies: {e}. Please install 'reportlab'.".encode()

    # Export figures to PNG bytes (needs kaleido)
    try:
        line_png = pio.to_image(fig_line, format="png", scale=2)
        pie_png = pio.to_image(fig_pie, format="png", scale=2)
    except Exception as e:
        return f"Chart export failed: {e}. Please ensure 'kaleido' is installed.".encode()

    buf = BytesIO()
    page_w, page_h = letter  # 612 x 792 pt
    margin = 40
    y = page_h - margin

    c = canvas.Canvas(buf, pagesize=letter)

    # Header
    c.setFillColorRGB(0.8, 0, 0)  # NEU red
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin, y, f"Participation Report — {student_name}")
    y -= 24
    c.setStrokeColorRGB(0.8, 0, 0)
    c.line(margin, y, page_w - margin, y)
    y -= 20

    # Metrics
    c.setFillColor(black)
    c.setFont("Helvetica", 11)
    metrics_lines = [
        f"Total Participation: {metrics.get('total', 0)}",
        f"Average Weekly Participation: {metrics.get('avg', 0):.2f}",
        f"Weeks Spoken: {metrics.get('spoke_weeks', 0)} / {metrics.get('weeks_total', 0)}",
        f"Participation Score: {metrics.get('score', 0):.1f}/100",
    ]
    for line in metrics_lines:
        c.drawString(margin, y, line)
        y -= 16

    y -= 6
    # Attendance summary
    c.setFillColorRGB(0.8, 0, 0)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margin, y, "Attendance Summary:")
    y -= 18
    c.setFillColor(black)
    c.setFont("Helvetica", 11)
    present = attendance_detail.get("present_count", 0)
    excused = attendance_detail.get("excused_count", 0)
    absent = attendance_detail.get("absent_count", 0)
    excused_suffix = attendance_detail.get("excused_suffix", "")
    absent_suffix = attendance_detail.get("absent_suffix", "")
    c.drawString(margin, y, f"Present: {present}")
    y -= 16
    c.drawString(margin, y, f"Excused: {excused}{excused_suffix}")
    y -= 16
    c.drawString(margin, y, f"Absent: {absent}{absent_suffix}")
    y -= 22

    # Line chart
    try:
        img = ImageReader(BytesIO(line_png))
        img_w = page_w - 2 * margin
        img_h = img_w * 0.4  # aspect heuristic
        if y - img_h < margin:
            c.showPage()
            y = page_h - margin
        c.drawImage(img, margin, y - img_h, width=img_w, height=img_h, preserveAspectRatio=True, mask='auto')
        y -= img_h + 18
    except Exception:
        pass

    # Pie chart
    try:
        img = ImageReader(BytesIO(pie_png))
        img_w = (page_w - 3 * margin) / 2
        img_h = img_w  # square
        if y - img_h < margin:
            c.showPage()
            y = page_h - margin
        c.drawImage(img, margin, y - img_h, width=img_w, height=img_h, preserveAspectRatio=True, mask='auto')
        y -= img_h + 12
    except Exception:
        pass

    c.showPage()
    c.save()
    return buf.getvalue()

# Sidebar filters
st.sidebar.header("🔍 Filters")
# Manual refresh of cached data
if st.sidebar.button("Refresh"):
    st.cache_data.clear()
    # Use new API when available, fallback for older versions
    if hasattr(st, "rerun"):
        st.rerun()
    else:
        st.experimental_rerun()
view_mode = st.sidebar.radio("Select View", ["Overview", "By Student", "By Week"])

# -------------------------------
# OVERVIEW MODE
# -------------------------------
if view_mode == "Overview":
    st.subheader("Class Overview")
    # Leaderboard toggle and computation
    leaderboard_mode = st.radio(
        "Leaderboard Metric",
        ["Total Participation", "Participation Score"],
        horizontal=True,
        key="leaderboard_mode",
    )

    # Aggregates
    total = (
        df.groupby("Student")["Participation"]
        .sum()
        .reset_index()
        .sort_values(by="Participation", ascending=False)
    )
    weeks_total = int(df["Week"].nunique())
    by_student_week = df.groupby(["Student", "Week"])['Participation'].sum().reset_index()

    # Scores per student
    scores = []
    for student, sdf in by_student_week.groupby("Student"):
        scores.append({
            "Student": student,
            "Score": calculate_score(sdf, weeks_total=weeks_total)
        })
    score_df = pd.DataFrame(scores).sort_values(by="Score", ascending=False)

    # Select dataset and styling based on toggle
    if leaderboard_mode == "Total Participation":
        data = total
        x_col = "Participation"
        title = "Total Participation Leaderboard"
        color_col = "Participation"
        color_scale = "Reds"
    else:
        data = score_df
        x_col = "Score"
        title = "Participation Score Leaderboard"
        color_col = "Score"
        color_scale = "Reds"

    n_students = len(data)
    fig = px.bar(
        data,
        y="Student",
        x=x_col,
        orientation="h",
        title=title,
        color=color_col,
        color_continuous_scale=color_scale,
        height=max(420, 28 * n_students),
    )
    category_array = data["Student"].tolist()[::-1]
    fig.update_yaxes(
        categoryorder="array",
        categoryarray=category_array,
        tickmode="array",
        tickvals=data["Student"].tolist(),
        ticktext=data["Student"].tolist(),
        ticklabelstep=1,
        automargin=True,
        tickfont=dict(size=10),
    )
    if leaderboard_mode == "Participation Score":
        fig.update_xaxes(range=[0, 100])
    fig.update_layout(margin=dict(l=180, r=20, t=60, b=20))
    st.plotly_chart(fig, use_container_width=True)

    # Average participation per week - sort numerically by week number
    avg = df.groupby("Week")["Participation"].mean().reset_index()
    def natural_sort_key(week_str):
        """Extract week number for proper numeric sorting"""
        match = re.search(r'Week\s*(\d+)', str(week_str))
        return int(match.group(1)) if match else 0
    avg["_sort"] = avg["Week"].apply(natural_sort_key)
    avg = avg.sort_values("_sort").drop("_sort", axis=1)
    
    fig2 = px.bar(
        avg,
        x="Week",
        y="Participation",
        title="Average Participation Per Week",
        color="Participation",
        color_continuous_scale="Reds"
    )
    fig2.update_xaxes(categoryorder='array', categoryarray=avg["Week"].tolist())
    st.plotly_chart(fig2, use_container_width=True)

    # Attendance summary
    attendance = df["Attendance"].value_counts().reset_index()
    attendance.columns = ["Status", "Count"]
    fig3 = px.pie(
        attendance,
        names="Status",
        values="Count",
        title="Overall Attendance Breakdown",
        color_discrete_sequence=px.colors.qualitative.Safe
    )
    st.plotly_chart(fig3, use_container_width=True)
# -------------------------------
# BY STUDENT MODE
# -------------------------------
elif view_mode == "By Student":
    students = sorted(df["Student"].unique())
    student_choice = st.sidebar.selectbox("Select a student", students)
    student_df = df[df["Student"] == student_choice]

    st.subheader(f"Participation Trend – {student_choice}")

    # Line chart
    fig4 = px.line(
        student_df,
        x="Week",
        y="Participation",
        markers=True,
        text="Participation",
        title=f"Participation Over Time – {student_choice}",
    )
    st.plotly_chart(fig4, use_container_width=True)

    # Attendance pie chart
    att = student_df["Attendance"].value_counts().reset_index()
    att.columns = ["Status", "Count"]
    fig5 = px.pie(
        att,
        names="Status",
        values="Count",
        title=f"Attendance Breakdown – {student_choice}",
        color_discrete_sequence=px.colors.qualitative.Pastel
    )
    st.plotly_chart(fig5, use_container_width=True)

    total = student_df["Participation"].sum()
    avg = student_df["Participation"].mean()
    spoke_weeks = (student_df["Participation"] > 0).sum()
    weeks_total = int(df["Week"].nunique())
    score = calculate_score(student_df, weeks_total=weeks_total)

    # Attendance breakdown summary
    attendance_counts = student_df["Attendance"].value_counts().to_dict()
    present_count = attendance_counts.get("Present", 0)
    excused_count = attendance_counts.get("Excused", 0)
    absent_count = attendance_counts.get("Absent", 0)

    # Build inline list of absent weeks with dates e.g., wk 2 (9/11), wk 5 (10/2)
    def _format_absent_label(label: str) -> str:
        if not isinstance(label, str):
            return ""
        # Match patterns like 'Week 2 (9/11)' with optional spaces and date
        m = re.search(r"Week\s*(\d+)\s*(?:\(([^)]+)\))?", label, flags=re.IGNORECASE)
        if not m:
            return label.strip()
        wk = m.group(1)
        date = (m.group(2) or "").strip()
        return f"wk {wk} ({date})".strip() if date else f"wk {wk}"

    absent_weeks_raw = (
        student_df.loc[student_df["Attendance"] == "Absent", "Week"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )
    absent_weeks_fmt = [_format_absent_label(w) for w in absent_weeks_raw]
    absent_weeks_fmt = [w for w in absent_weeks_fmt if w]
    absent_suffix = f" - {', '.join(absent_weeks_fmt)}" if absent_weeks_fmt else ""

    # Excused weeks in same format
    excused_weeks_raw = (
        student_df.loc[student_df["Attendance"] == "Excused", "Week"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )
    excused_weeks_fmt = [_format_absent_label(w) for w in excused_weeks_raw]
    excused_weeks_fmt = [w for w in excused_weeks_fmt if w]
    excused_suffix = f" - {', '.join(excused_weeks_fmt)}" if excused_weeks_fmt else ""

    st.markdown(
        f"""
        <span class=\"metric-label\">Total Participation:</span> {total}<br>
        <span class=\"metric-label\">Average Weekly Participation:</span> {avg:.2f}<br>
        <span class=\"metric-label\">Weeks Spoken:</span> {spoke_weeks} / {weeks_total}<br>
        <span class=\"metric-label\">Participation Score:</span> {score}/100
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        "<div class=\"metric-label\">Attendance Summary:</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        - ✅ Present: {present_count}  
        - 💛 Excused: {excused_count}{excused_suffix}  
        - ❌ Absent: {absent_count}{absent_suffix}
        """
    )

    # Download PDF report
    metrics_dict = {
        "total": int(total),
        "avg": float(avg) if pd.notna(avg) else 0.0,
        "spoke_weeks": int(spoke_weeks),
        "weeks_total": int(weeks_total),
        "score": float(score),
    }
    attendance_detail = {
        "present_count": int(present_count),
        "excused_count": int(excused_count),
        "absent_count": int(absent_count),
        "excused_suffix": excused_suffix,
        "absent_suffix": absent_suffix,
    }
    pdf_bytes = generate_student_pdf(student_choice, fig4, fig5, metrics_dict, attendance_detail)
    st.download_button(
        label="Download PDF report",
        data=pdf_bytes,
        file_name=f"{student_choice.replace(' ', '_')}_participation_report.pdf",
        mime="application/pdf",
    )

# -------------------------------
# BY WEEK MODE
# -------------------------------
else:
    # Sort weeks numerically by week number, not alphabetically
    def natural_sort_key(week_str):
        """Extract week number for proper numeric sorting"""
        match = re.search(r'Week\s*(\d+)', str(week_str))
        return int(match.group(1)) if match else 0
    
    weeks = sorted(df["Week"].unique(), key=natural_sort_key)
    week_choice = st.sidebar.selectbox("Select a week", weeks)
    week_df = df[df["Week"] == week_choice]
    topic = week_df["Topic"].iloc[0] if not week_df.empty else ""

    st.subheader(f"{week_choice}: {topic}")

    # Control who shows in the chart
    show_only_participants = st.sidebar.checkbox(
        "Show only students with participation", value=True
    )

    # Aggregate in case of multiple rows per student-week
    week_agg = (
        week_df.groupby("Student")["Participation"].sum().reset_index()
    )
    # Optionally include all students (with zero participation) to make absence obvious
    if not show_only_participants:
        all_students = pd.DataFrame({"Student": sorted(df["Student"].unique())})
        week_agg = (
            all_students
            .merge(week_agg, on="Student", how="left")
            .fillna({"Participation": 0})
        )
    # Sort descending by Participation
    week_agg = week_agg.sort_values(by="Participation", ascending=False)

    # Build figure with dynamic height and full label visibility
    n_students = len(week_agg)
    fig6 = px.bar(
        week_agg,
        y="Student",
        x="Participation",
        orientation="h",
        color="Participation",
        color_continuous_scale="Reds",
        title=f"Participation by Student – {week_choice}",
        height=max(420, 28 * n_students),
    )
    category_array = week_agg["Student"].tolist()[::-1]
    fig6.update_yaxes(
        categoryorder="array",
        categoryarray=category_array,
        tickmode="array",
        tickvals=week_agg["Student"].tolist(),
        ticktext=week_agg["Student"].tolist(),
        ticklabelstep=1,
        automargin=True,
        tickfont=dict(size=10),
    )
    # X-axis padding
    xmax = float(week_agg["Participation"].max() or 0)
    fig6.update_xaxes(range=[0, xmax * 1.1 if xmax > 0 else 1])
    fig6.update_layout(margin=dict(l=180, r=20, t=60, b=20))
    st.plotly_chart(fig6, use_container_width=True)

    att = week_df["Attendance"].value_counts().reset_index()
    att.columns = ["Status", "Count"]
    fig7 = px.pie(
        att,
        names="Status",
        values="Count",
        title=f"Attendance Breakdown – {week_choice}",
        color_discrete_sequence=px.colors.qualitative.Set2
    )
    st.plotly_chart(fig7, use_container_width=True)
