import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook

MEAL_LABELS = ["B","L","P"]

def _safe_str(x):
    try:
        return str(x) if not pd.isna(x) else ""
    except Exception:
        return str(x)

def parse_sheet(df):
    facility = None
    class_name = None

    for r in range(0, 30):
        for c in range(min(df.shape[1], 100)):
            v = df.iat[r, c]
            if isinstance(v, str):
                if r == 7 and c == 3:
                    facility = v.strip()
                if "Class" in v:
                    m = re.search(r"Class\s*([A-Za-z0-9]+)", v, re.IGNORECASE)
                    if m:
                        class_name = m.group(1)

    day_row = label_row = None
    for r in range(8, min(60, df.shape[0])):
        row_vals = df.iloc[r].tolist()
        if any(isinstance(x, str) and x.strip().endswith(".") for x in row_vals):
            day_row = r
            label_row = r + 1
            break
    if day_row is None or label_row is None:
        return None

    day_starts = []
    for c in range(df.shape[1]):
        day = df.iat[day_row, c]
        if isinstance(day, str) and day.strip().endswith("."):
            day_starts.append((c, day.strip().rstrip(".")))
    if not day_starts:
        return None

    blocks = {}
    day_starts_sorted = sorted(day_starts)
    for i, (c, dayname) in enumerate(day_starts_sorted):
        end = day_starts_sorted[i+1][0] if i+1 < len(day_starts_sorted) else df.shape[1]
        labels = []
        for cc in range(c, end):
            val = df.iat[label_row, cc]
            if isinstance(val, str) and val.strip() in (["At"] + MEAL_LABELS):
                labels.append((cc, val.strip()))
        labels = sorted(labels, key=lambda x: x[0])
        if labels:
            blocks[dayname] = labels
    if not blocks:
        return None

    start_row = None
    for r in range(label_row + 1, min(label_row + 200, df.shape[0])):
        v0 = df.iat[r, 0]
        v1 = df.iat[r, 1]
        if (isinstance(v0, (int, float)) and not pd.isna(v0)) and isinstance(v1, str) and v1.strip():
            start_row = r
            break
    if start_row is None:
        return None

    end_row = start_row
    for r in range(start_row, df.shape[0]):
        row = df.iloc[r, 0:10]
        if row.isna().sum() >= 8:
            end_row = r - 1
            break
    if end_row < start_row:
        end_row = min(start_row + 80, df.shape[0] - 1)

    students_df = df.iloc[start_row:end_row+1, :].copy()

    rows = []
    for day, cols in blocks.items():
        att_cols = [col for col, label in cols if label == "At"]
        att_marked = 0
        if att_cols:
            col_vals = students_df.iloc[:, att_cols[0]]
            att_marked = int(col_vals.apply(lambda x: isinstance(x, str) and x.strip() != "").sum())
        for col_idx, label in cols:
            if label in MEAL_LABELS:
                col_vals = students_df.iloc[:, col_idx]
                marked = int(col_vals.apply(lambda x: isinstance(x, str) and x.strip() != "").sum())
                rows.append({"Day": day, "Meal": label, "MarkedCount": marked, "AttendanceMarked": att_marked})

    counts = pd.DataFrame(rows)

    week_dates = []
    date_label_row = day_row - 1
    for c in range(df.shape[1]):
        if _safe_str(df.iat[date_label_row, c]).strip() == "Date":
            for rr in range(date_label_row + 1, date_label_row + 3):
                val = df.iat[rr, c]
                try:
                    dt = pd.to_datetime(val)
                    if not pd.isna(dt):
                        week_dates.append(dt)
                        break
                except Exception:
                    continue
    week_start = min(week_dates).date() if week_dates else None
    week_end   = max(week_dates).date() if week_dates else None

    return {"Campus": facility or "Unknown Campus",
            "Class": class_name or "Unknown Class",
            "WeekStart": week_start,
            "WeekEnd": week_end,
            "Counts": counts,
            "StudentRows": (start_row, end_row)}

def build_report(file_bytes):
    xls = pd.ExcelFile(file_bytes)
    all_records = []
    for s in xls.sheet_names:
        df = pd.read_excel(file_bytes, sheet_name=s, header=None)
        parsed = parse_sheet(df)
        if not parsed or parsed["Counts"].empty:
            continue
        campus = parsed["Campus"]
        cls = parsed["Class"]
        wk = f"{parsed['WeekStart']} to {parsed['WeekEnd']}" if parsed["WeekStart"] else "Unknown Week"
        counts = parsed["Counts"].copy()
        counts["Issue"] = (counts["MarkedCount"] == 0) & (counts["AttendanceMarked"] > 0)
        for _, row in counts.iterrows():
            all_records.append({
                "Campus": campus,
                "Class": cls,
                "Week": wk,
                "Day": row["Day"],
                "Meal": row["Meal"],
                "MarkedCount": int(row["MarkedCount"]),
                "AttendanceMarked": int(row["AttendanceMarked"]),
                "Issue_MissingRecording": bool(row["Issue"]),
            })
    summary_df = pd.DataFrame(all_records)
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        if summary_df.empty:
            pd.DataFrame(columns=["Campus","Month","MissingMealRecordings","Pct_Days_All3MealsRecorded"]).to_excel(writer, sheet_name="Dashboard_Month", index=False)
            pd.DataFrame(columns=["Campus","Month","Week","MissingMealRecordings","Pct_Days_All3MealsRecorded"]).to_excel(writer, sheet_name="Dashboard_Month_Week", index=False)
            summary_df.to_excel(writer, sheet_name="Summary_All", index=False)
        else:
            def parse_week_start(s):
                try:
                    if isinstance(s, str) and "to" in s:
                        start = s.split("to")[0].strip()
                        return pd.to_datetime(start).date()
                except Exception:
                    return pd.NaT
                return pd.NaT
            summary_df["WeekStart"] = summary_df["Week"].apply(parse_week_start)
            summary_df["Month"] = pd.to_datetime(summary_df["WeekStart"]).dt.to_period("M").astype(str)
            miss = summary_df[summary_df["Issue_MissingRecording"] == True]
            campus_month_missing = (miss.groupby(["Campus","Month"]).size().reset_index(name="MissingMealRecordings"))
            meals3 = summary_df[summary_df["Meal"].isin(MEAL_LABELS)].copy()
            daily_ok = (meals3.groupby(["Campus","Class","Week","WeekStart","Month","Day"])["MarkedCount"]
                              .apply(lambda s: int((s > 0).sum() == 3))
                              .reset_index(name="All3MealsRecorded"))
            campus_month_kpi = (daily_ok.groupby(["Campus","Month"])["All3MealsRecorded"].mean().reset_index(name="Pct_Days_All3MealsRecorded"))
            dash_month = pd.merge(campus_month_missing, campus_month_kpi, on=["Campus","Month"], how="outer") \
                            .fillna({"MissingMealRecordings":0,"Pct_Days_All3MealsRecorded":0})
            campus_week_missing = (miss.groupby(["Campus","Month","Week"]).size().reset_index(name="MissingMealRecordings"))
            campus_week_kpi = (daily_ok.groupby(["Campus","Month","Week"])["All3MealsRecorded"].mean().reset_index(name="Pct_Days_All3MealsRecorded"))
            dash_month_week = pd.merge(campus_week_missing, campus_week_kpi, on=["Campus","Month","Week"], how="outer") \
                                 .fillna({"MissingMealRecordings":0,"Pct_Days_All3MealsRecorded":0})
            # Write
            dash_month.sort_values(["Month","Campus"]).to_excel(writer, sheet_name="Dashboard_Month", index=False)
            dash_month_week.sort_values(["Month","Week","Campus"]).to_excel(writer, sheet_name="Dashboard_Month_Week", index=False)
            summary_df.to_excel(writer, sheet_name="Summary_All", index=False)
    out.seek(0)
    return out

st.set_page_config(page_title="Meal Count Monitoring (B/L/P)", layout="wide")
st.title("Meal Count Monitoring (Breakfast, Lunch, PM Snack)")

uploaded = st.file_uploader("Upload the Weekly Forms workbook (.xlsx)", type=["xlsx"])
if uploaded:
    if st.button("Generate Monitoring Report"):
        report_bytes = build_report(uploaded)
        st.success("Report generated.")
        st.download_button("Download Excel Report", report_bytes, file_name="MealCount_Full_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
