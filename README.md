# American Healthcare Dashboard

A Streamlit dashboard for class participation and attendance analytics with a Northeastern-themed design.

## Features
- Auto-detects Excel file with cleaned data (American_Healthcare_Class_Cleaned.xlsx)
- Overview leaderboards (Total Participation and Participation Score)
- By Student deep dive with score, attendance breakdown, absent/excused week/date tags
- By Week view with option to show only participants or everyone
- Theming: Northeastern red/white, custom fonts, Plotly defaults
- Auto-refresh: cache TTL + manual Refresh Data button
- Export: Download per-student PDF including charts and attendance summary

## Project Structure
- `dashboard.py` – main Streamlit app
- `american_healthcare_dashboard/` – earlier version of the app (optional)
- `.venv/` – local virtual environment (ignored by git)

## Local Setup
1. Create and activate a virtual environment (optional if you have one already):
   - macOS/Linux (zsh):
     ```
     python3 -m venv .venv
     source .venv/bin/activate
     ```
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Run the app:
   ```
   streamlit run dashboard.py
   ```

## Deployment (Streamlit Community Cloud)
1. Push this repo to GitHub.
2. In Streamlit Cloud, create a new app pointing to `dashboard.py`.
3. Add any required secrets to `.streamlit/secrets.toml` (not required for public data).
4. Set Python version to 3.11+ and make sure requirements are installed.

## Requirements
See `requirements.txt`.

Includes:
- streamlit, pandas, plotly, openpyxl
- kaleido (Plotly static export)
- reportlab (PDF generation)

## Notes
- Data file expected: `American_Healthcare_Class_Cleaned.xlsx` in the repo root.
- If your Excel file name changes, update `dashboard.py` or add auto-discovery logic.
