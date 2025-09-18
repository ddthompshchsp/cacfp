
from textwrap import dedent

debug_app_path = "/mnt/data/meal_count_streamlit_debug.py"

code = dedent(r"""
import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime
import matplotlib.pyplot as plt

st.set_page_config(page_title="Meal Count Debugger (B/L/P)", layout="wide")
st.title("🛠️ Meal Count Debugger — Breakfast, Lunch, PM Snack")

st.markdown("""
If your last run produced an *empty spreadsheet*, use this debugger.
It will auto-detect flexible headers and show you exactly what it found **per sheet**.

**What this app does:**
- Tries multiple header spellings and positions (not hard-coded).
- Accepts day names with or without periods (`Mon` or `Mon.`).
- Accepts meal labels like `B`, `Br`, `Breakfast`, `L`, `Lunch`, `P`, `PM Snack`, `Snack PM`, etc.
- Counts a **mark** if the cell is **not empty** (number, X, check, etc.), not just strings.
- Lets you treat **attendance-less** days as present if any meal is marked (optional).
""")

# --- Normalizers & matching helpers ---
def norm(s):
    s = str(s).strip().lower()
    s = re.sub(r'[^a-z0-9 ]+', '', s)
    return s

def is_day_token(s):
    if not isinstance(s, str): return False
    t = norm(s)
    for day in ["mon","tue","wed","thu","fri","monday","tuesday","wednesday","thursday","friday"]:
        if t.startswith(day):
            return True
    return False

def meal_label_from_token(s):
    t = norm(s)
    # attendance
    if t.startswith("at") or t.startswith("att") or "attend" in t:
        return "At"
    # breakfast
    if t in ["b","br","bf","breakfast","am snack","amsnack"] or t.startswith("break") or t.startswith("brk"):
        return "B"
    # lunch
    if t in ["l","ln","lunch"] or t.startswith("lun"):
        return "L"
    # pm snack
    if t in ["p","pm","pmsnack","snack pm","snackpm","pm snack","pmmeal","snack"] or "pm" in t and "snack" in t:
        return "P"
    return None

def nonempty_mark(x):
    if pd.isna(x): return False
    if isinstance(x, str):
        return len(x.strip()) > 0
    # numbers/booleans/counts count as marked
    return True

uploaded = st.file_uploader("Upload your .xlsx workbook", type=["xlsx"])
assume_attendance_if_any_meal = st.checkbox("If attendance is blank, treat day as present when ANY meal is marked", value=True)

def parse_sheet(df):
    # 1) Find day row (first row that has >=2 day tokens)
    day_row = None
    for r in range(min(80, df.shape[0])):
        tokens = [is_day_token(v) for v in df.iloc[r].tolist()]
        if sum(tokens) >= 2:
            day_row = r
            break
    if day_row is None:
        return None, "No day row found"

    # 2) Find label row = next non-empty row
    label_row = None
    for r in range(day_row+1, min(day_row+6, df.shape[0])):
        row_vals = df.iloc[r].tolist()
        if any((isinstance(x, str) and x.strip()) or (not pd.isna(x)) for x in row_vals):
            label_row = r
            break
    if label_row is None:
        return None, "No label row found after day row"

    # 3) Build day blocks by locating day tokens in day_row
    day_starts = []
    for c in range(df.shape[1]):
        v = df.iat[day_row, c]
        if is_day_token(v):
            # Canonicalize label to Mon/Tue/.. short form
            t = norm(v)
            short = v.strip().split()[0].rstrip('.')
            day_starts.append((c, short))
    if not day_starts:
        return None, "Could not locate day start columns"
    day_starts.sort()

    # 4) Within each day-block, map label cells on label_row to At/B/L/P
    blocks = {}
    for i, (c, day) in enumerate(day_starts):
        end = day_starts[i+1][0] if i+1 < len(day_starts) else df.shape[1]
        labels = []
        for cc in range(c, end):
            val = df.iat[label_row, cc]
            if isinstance(val, str) or not pd.isna(val):
                ml = meal_label_from_token(val)
                if ml in ["At","B","L","P"]:
                    labels.append((cc, ml))
        if labels:
            blocks[day] = sorted(labels, key=lambda x: x[0])

    if not blocks:
        return None, "Found day row but no meal/attendance labels underneath"

    # 5) Find start of student rows: look for first row below label_row with a non-empty name col (col1) and any mark in the block
    start_row = None
    for r in range(label_row+1, min(label_row+200, df.shape[0])):
        name_cell = df.iat[r, 1] if df.shape[1] > 1 else None
        if isinstance(name_cell, str) and name_cell.strip():
            start_row = r
            break
    if start_row is None:
        return None, "No student rows detected"

    # 6) Find end of student rows: first mostly empty row
    end_row = start_row
    for r in range(start_row, df.shape[0]):
        row = df.iloc[r, 0:10]
        if row.isna().sum() >= 8:
            end_row = r - 1
            break
    if end_row < start_row:
        end_row = min(start_row + 60, df.shape[0]-1)

    students = df.iloc[start_row:end_row+1, :].copy()

    # 7) Count marks
    rows = []
    for day, cols in blocks.items():
        # Attendance
        att_cols = [col for col, lab in cols if lab == "At"]
        att_present = 0
        if att_cols:
            att_col = students.iloc[:, att_cols[0]]
            att_present = int(att_col.apply(nonempty_mark).sum())

        # If requested: infer attendance from any meal marks
        if att_present == 0 and assume_attendance_if_any_meal:
            meal_cols = [col for col, lab in cols if lab in ["B","L","P"]]
            if meal_cols:
                any_meal = students.iloc[:, meal_cols].applymap(nonempty_mark).any(axis=1).sum()
                att_present = int(any_meal)

        for col_idx, lab in cols:
            if lab in ["B","L","P"]:
                col_vals = students.iloc[:, col_idx]
                marks = int(col_vals.apply(nonempty_mark).sum())
                rows.append({"Day": day, "Meal": lab, "MarkedCount": marks, "AttendanceMarked": att_present})
    result = pd.DataFrame(rows)

    # Try to scrape week start by scanning above the day row for a date-like value
    week_start = None
    for r in range(max(0, day_row-5), day_row+1):
        for c in range(df.shape[1]):
            v = df.iat[r, c]
            try:
                dt = pd.to_datetime(v, errors="coerce")
                if pd.notna(dt):
                    if week_start is None or dt < week_start:
                        week_start = dt
            except Exception:
                pass

    info = {
        "day_row": day_row,
        "label_row": label_row,
        "start_row": start_row,
        "end_row": end_row,
        "blocks": blocks,
        "week_start": str(week_start.date()) if isinstance(week_start, pd.Timestamp) else None
    }
    return (result, info), None

def build_all(file_like):
    xls = pd.ExcelFile(file_like)
    debug = []
    all_records = []
    for s in xls.sheet_names:
        df = pd.read_excel(file_like, sheet_name=s, header=None)
        parsed, err = parse_sheet(df)
        if err:
            debug.append({"Sheet": s, "Status": "ERROR", "Detail": err})
            continue
        (counts, info) = parsed
        debug.append({"Sheet": s, "Status": "OK", "Detail": info})
        for _, row in counts.iterrows():
            all_records.append({
                "Sheet": s,
                "Day": row["Day"],
                "Meal": row["Meal"],
                "MarkedCount": int(row["MarkedCount"]),
                "AttendanceMarked": int(row["AttendanceMarked"])
            })
    return pd.DataFrame(all_records), pd.DataFrame(debug)

def make_excel(summary, debug_df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        debug_df.to_excel(writer, sheet_name="Debug_Log", index=False)
        summary.to_excel(writer, sheet_name="Extracted_Counts", index=False)
    out.seek(0)
    return out

uploaded = st.file_uploader("Upload the workbook to debug", type=["xlsx"])
if uploaded:
    with st.spinner("Scanning sheets and detecting headers..."):
        summary, debug_df = build_all(uploaded)

    st.subheader("🔎 Debug Log (per sheet)")
    st.dataframe(debug_df, use_container_width=True, height=300)

    st.subheader("✅ Extracted Counts (raw)")
    st.dataframe(summary.head(1000), use_container_width=True, height=400)

    st.download_button("Download Debug Excel", make_excel(summary, debug_df), file_name="MealCount_Debug_Output.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Upload your Excel to diagnose parsing issues.")
""")

with open(debug_app_path, "w", encoding="utf-8") as f:
    f.write(code)

debug_app_path

