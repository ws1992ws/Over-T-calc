
import streamlit as st
import pandas as pd
from datetime import date, timedelta
from io import BytesIO
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# APP SETUP
# ============================================================
st.set_page_config(
    page_title="Overtime Calculator",
    page_icon="🧮",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_RULES = {
    "weekday_factor": 3.5,
    "weekend_factor": 7.0,
    "salary_divisor": 155.0,
    "transport_divisor": 30.0,
}

SALARY_COLUMNS = ["Effective Date", "Basic Salary (SAR)"]
TRANSPORT_COLUMNS = ["Effective Date", "Transportation Allowance (SAR)"]
EXCLUSION_COLUMNS = ["Type", "Start Date", "End Date"]

st.markdown("""
<style>
.block-container {
    padding-top: 1.4rem;
    padding-bottom: 2rem;
    max-width: 1500px;
}
[data-testid="stMetric"] {
    border: 1px solid rgba(49,51,63,.12);
    padding: 0.8rem;
    border-radius: 0.7rem;
}
div[data-testid="stDataEditor"] {
    border-radius: 0.6rem;
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# SESSION STATE
# ============================================================
def empty_exclusions():
    return pd.DataFrame({
        "Type": pd.Series(dtype="object"),
        "Start Date": pd.Series(dtype="object"),
        "End Date": pd.Series(dtype="object"),
    })

def initialize_state():
    st.session_state.setdefault("period_start", date(date.today().year, 1, 1))
    st.session_state.setdefault("period_end", date.today())

    st.session_state.setdefault(
        "salary_data",
        pd.DataFrame({
            "Effective Date": [date(date.today().year, 1, 1)],
            "Basic Salary (SAR)": [0.0],
        })
    )

    st.session_state.setdefault(
        "transport_data",
        pd.DataFrame({
            "Effective Date": [date(date.today().year, 1, 1)],
            "Transportation Allowance (SAR)": [0.0],
        })
    )

    st.session_state.setdefault("exclusion_data", empty_exclusions())

    for key, value in DEFAULT_RULES.items():
        st.session_state.setdefault(key, value)

    st.session_state.setdefault("editor_generation", 0)
    st.session_state.setdefault("calc_result", None)
    st.session_state.setdefault("load_message", "")

initialize_state()

# ============================================================
# DATA HELPERS
# ============================================================
def ensure_columns(df, columns):
    if df is None:
        return pd.DataFrame(columns=columns)
    df = df.copy()
    for c in columns:
        if c not in df.columns:
            df[c] = None
    return df[columns]

def parse_date_value(value):
    if value is None or pd.isna(value):
        return None
    if isinstance(value, date):
        return value
    return pd.to_datetime(value).date()

def normalize_history(df, date_col, amount_col):
    df = ensure_columns(df, [date_col, amount_col])
    clean = df.dropna(subset=[date_col, amount_col]).copy()

    if clean.empty:
        return []

    clean[date_col] = clean[date_col].apply(parse_date_value)
    clean[amount_col] = pd.to_numeric(clean[amount_col], errors="coerce")
    clean = clean.dropna(subset=[date_col, amount_col])
    clean = clean.sort_values(date_col)
    clean = clean.drop_duplicates(subset=[date_col], keep="last")

    return list(zip(clean[date_col], clean[amount_col]))

def normalize_exclusions(df):
    df = ensure_columns(df, EXCLUSION_COLUMNS)
    clean = df.dropna(subset=EXCLUSION_COLUMNS).copy()

    if clean.empty:
        return []

    clean["Start Date"] = clean["Start Date"].apply(parse_date_value)
    clean["End Date"] = clean["End Date"].apply(parse_date_value)

    rows = []
    for _, row in clean.iterrows():
        start = row["Start Date"]
        end = row["End Date"]

        if start is None or end is None:
            continue

        if end < start:
            raise ValueError(
                f"{row['Type']} exclusion has an End Date before its Start Date."
            )

        rows.append((str(row["Type"]), start, end))

    return rows

def effective_value(history, target_date):
    eligible = [item for item in history if item[0] <= target_date]
    if not eligible:
        return None
    return max(eligible, key=lambda x: x[0])[1]

def exclusion_reason(target_date, exclusion_history):
    reasons = [
        typ for typ, start, end in exclusion_history
        if start <= target_date <= end
    ]
    if not reasons:
        return None
    return " / ".join(dict.fromkeys(reasons))

# ============================================================
# CALCULATION ENGINE
# ============================================================
def calculate_overtime(
    start_date,
    end_date,
    salary_df,
    transport_df,
    exclusion_df,
    rules,
):
    if end_date < start_date:
        raise ValueError("End date cannot be before Start date.")

    if rules["salary_divisor"] == 0:
        raise ValueError("Salary divisor cannot be zero.")

    if rules["transport_divisor"] == 0:
        raise ValueError("Transportation divisor cannot be zero.")

    salary_history = normalize_history(
        salary_df, "Effective Date", "Basic Salary (SAR)"
    )
    transport_history = normalize_history(
        transport_df, "Effective Date", "Transportation Allowance (SAR)"
    )
    exclusion_history = normalize_exclusions(exclusion_df)

    if not salary_history:
        raise ValueError("Enter at least one Basic Salary.")

    if not transport_history:
        raise ValueError("Enter at least one Transportation Allowance.")

    rows = []
    current = start_date

    while current <= end_date:
        reason = exclusion_reason(current, exclusion_history)

        if reason:
            rows.append({
                "Date": current,
                "Day": current.strftime("%A"),
                "Status": "Excluded",
                "Exclusion Reason": reason,
                "Basic Salary (SAR)": None,
                "Transportation Allowance (SAR)": None,
                "OT Type": "Excluded",
                "Base OT (SAR)": 0.0,
                "Transportation OT (SAR)": 0.0,
                "Daily Total (SAR)": 0.0,
            })
        else:
            basic = effective_value(salary_history, current)
            transportation = effective_value(transport_history, current)

            if basic is None:
                raise ValueError(
                    f"No Basic Salary is effective on {current.isoformat()}. "
                    "Add a salary with an effective date on or before this date."
                )

            if transportation is None:
                raise ValueError(
                    f"No Transportation Allowance is effective on {current.isoformat()}. "
                    "Add an allowance with an effective date on or before this date."
                )

            is_weekend = current.weekday() in (4, 5)  # Friday / Saturday

            if is_weekend:
                ot_type = "Weekend"
                base_ot = (
                    float(basic)
                    * rules["weekend_factor"]
                    / rules["salary_divisor"]
                )
            else:
                ot_type = "Weekday"
                base_ot = (
                    float(basic)
                    * rules["weekday_factor"]
                    / rules["salary_divisor"]
                )

            transport_ot = (
                float(transportation)
                / rules["transport_divisor"]
            )

            total = base_ot + transport_ot

            rows.append({
                "Date": current,
                "Day": current.strftime("%A"),
                "Status": "Included",
                "Exclusion Reason": "",
                "Basic Salary (SAR)": float(basic),
                "Transportation Allowance (SAR)": float(transportation),
                "OT Type": ot_type,
                "Base OT (SAR)": base_ot,
                "Transportation OT (SAR)": transport_ot,
                "Daily Total (SAR)": total,
            })

        current += timedelta(days=1)

    daily_df = pd.DataFrame(rows)
    included = daily_df[daily_df["Status"] == "Included"].copy()

    monthly_records = []
    if not included.empty:
        included["Month"] = pd.to_datetime(included["Date"]).dt.to_period("M").astype(str)

        for month, group in included.groupby("Month", sort=True):
            monthly_records.append({
                "Month": month,
                "Weekday Days": int((group["OT Type"] == "Weekday").sum()),
                "Weekend Days": int((group["OT Type"] == "Weekend").sum()),
                "Weekday OT (SAR)": group.loc[
                    group["OT Type"] == "Weekday", "Base OT (SAR)"
                ].sum(),
                "Weekend OT (SAR)": group.loc[
                    group["OT Type"] == "Weekend", "Base OT (SAR)"
                ].sum(),
                "Transportation OT (SAR)": group["Transportation OT (SAR)"].sum(),
                "Monthly Total (SAR)": group["Daily Total (SAR)"].sum(),
            })

    monthly_df = pd.DataFrame(
        monthly_records,
        columns=[
            "Month",
            "Weekday Days",
            "Weekend Days",
            "Weekday OT (SAR)",
            "Weekend OT (SAR)",
            "Transportation OT (SAR)",
            "Monthly Total (SAR)",
        ],
    )

    summary = {
        "total": float(included["Daily Total (SAR)"].sum()) if not included.empty else 0.0,
        "weekday_days": int((included["OT Type"] == "Weekday").sum()) if not included.empty else 0,
        "weekend_days": int((included["OT Type"] == "Weekend").sum()) if not included.empty else 0,
        "excluded_days": int((daily_df["Status"] == "Excluded").sum()),
        "weekday_ot": float(
            included.loc[included["OT Type"] == "Weekday", "Base OT (SAR)"].sum()
        ) if not included.empty else 0.0,
        "weekend_ot": float(
            included.loc[included["OT Type"] == "Weekend", "Base OT (SAR)"].sum()
        ) if not included.empty else 0.0,
        "transport_ot": float(
            included["Transportation OT (SAR)"].sum()
        ) if not included.empty else 0.0,
    }

    return daily_df, monthly_df, summary

# ============================================================
# PROJECT SAVE / LOAD
# ============================================================
def dataframe_to_salary_records(df):
    df = ensure_columns(df, SALARY_COLUMNS)
    records = []

    for _, row in df.iterrows():
        d = parse_date_value(row["Effective Date"])
        amount = row["Basic Salary (SAR)"]

        if d is None or pd.isna(amount):
            continue

        records.append({
            "effective_date": d.isoformat(),
            "amount": float(amount),
        })

    return records

def dataframe_to_transport_records(df):
    df = ensure_columns(df, TRANSPORT_COLUMNS)
    records = []

    for _, row in df.iterrows():
        d = parse_date_value(row["Effective Date"])
        amount = row["Transportation Allowance (SAR)"]

        if d is None or pd.isna(amount):
            continue

        records.append({
            "effective_date": d.isoformat(),
            "amount": float(amount),
        })

    return records

def dataframe_to_exclusion_records(df):
    df = ensure_columns(df, EXCLUSION_COLUMNS)
    records = []

    for _, row in df.iterrows():
        typ = row["Type"]
        start = parse_date_value(row["Start Date"])
        end = parse_date_value(row["End Date"])

        if pd.isna(typ) or start is None or end is None:
            continue

        records.append({
            "type": str(typ),
            "start": start.isoformat(),
            "end": end.isoformat(),
        })

    return records

def build_project_json(
    period_start,
    period_end,
    salary_df,
    transport_df,
    exclusion_df,
    rules,
):
    project = {
        "app": "Overtime Calculator",
        "version": 4,
        "period": {
            "start": period_start.isoformat(),
            "end": period_end.isoformat(),
        },
        "rules": {
            "weekday_factor": float(rules["weekday_factor"]),
            "weekend_factor": float(rules["weekend_factor"]),
            "salary_divisor": float(rules["salary_divisor"]),
            "transport_divisor": float(rules["transport_divisor"]),
        },
        "salary_history": dataframe_to_salary_records(salary_df),
        "transport_history": dataframe_to_transport_records(transport_df),
        "exclusions": dataframe_to_exclusion_records(exclusion_df),
    }

    return json.dumps(project, indent=2).encode("utf-8")

def validate_project_dict(project):
    if not isinstance(project, dict):
        raise ValueError("This is not a valid project file.")

    if "period" not in project:
        raise ValueError("Project file is missing the overtime period.")

    period = project["period"]

    if "start" not in period or "end" not in period:
        raise ValueError("Project file is missing Start or End date.")

    start = date.fromisoformat(period["start"])
    end = date.fromisoformat(period["end"])

    if end < start:
        raise ValueError("Saved End date is before saved Start date.")

    return start, end

def apply_project(project):
    start, end = validate_project_dict(project)

    salary_records = project.get("salary_history", [])
    transport_records = project.get("transport_history", [])
    exclusion_records = project.get("exclusions", [])

    salary_df = pd.DataFrame(
        [
            {
                "Effective Date": date.fromisoformat(x["effective_date"]),
                "Basic Salary (SAR)": float(x["amount"]),
            }
            for x in salary_records
        ],
        columns=SALARY_COLUMNS,
    )

    transport_df = pd.DataFrame(
        [
            {
                "Effective Date": date.fromisoformat(x["effective_date"]),
                "Transportation Allowance (SAR)": float(x["amount"]),
            }
            for x in transport_records
        ],
        columns=TRANSPORT_COLUMNS,
    )

    exclusion_df = pd.DataFrame(
        [
            {
                "Type": x["type"],
                "Start Date": date.fromisoformat(x["start"]),
                "End Date": date.fromisoformat(x["end"]),
            }
            for x in exclusion_records
        ],
        columns=EXCLUSION_COLUMNS,
    )

    rules = DEFAULT_RULES.copy()
    saved_rules = project.get("rules", {})

    for key in rules:
        if key in saved_rules:
            rules[key] = float(saved_rules[key])

    # Replace ALL project state in one controlled step.
    st.session_state.period_start = start
    st.session_state.period_end = end
    st.session_state.salary_data = salary_df
    st.session_state.transport_data = transport_df
    st.session_state.exclusion_data = exclusion_df

    for key, value in rules.items():
        st.session_state[key] = value

    # Force fresh widgets after project load.
    st.session_state.editor_generation += 1
    st.session_state.calc_result = None
    st.session_state.load_message = (
        f"Project loaded: {start.isoformat()} to {end.isoformat()}"
    )

# ============================================================
# EXCEL EXPORT
# ============================================================
def make_excel_report(
    daily_df,
    monthly_df,
    summary,
    period_start,
    period_end,
    salary_df,
    transport_df,
    exclusion_df,
    rules,
):
    wb = Workbook()

    dark_fill = PatternFill("solid", fgColor="1F4E78")
    light_fill = PatternFill("solid", fgColor="D9EAF7")
    white_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)
    thin = Side(style="thin", color="D0D7DE")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Summary
    ws = wb.active
    ws.title = "Summary"

    ws["A1"] = "Overtime Calculation Report"
    ws["A1"].font = Font(size=18, bold=True)
    ws.merge_cells("A1:D1")

    ws["A3"] = "Period"
    ws["A3"].fill = dark_fill
    ws["A3"].font = white_font
    ws["B3"] = f"{period_start.isoformat()} to {period_end.isoformat()}"

    rows = [
        ("Total Overtime (SAR)", summary["total"]),
        ("Weekday OT Days", summary["weekday_days"]),
        ("Weekend OT Days", summary["weekend_days"]),
        ("Excluded Days", summary["excluded_days"]),
        ("Weekday OT (SAR)", summary["weekday_ot"]),
        ("Weekend OT (SAR)", summary["weekend_ot"]),
        ("Transportation OT (SAR)", summary["transport_ot"]),
    ]

    r = 5
    for label, value in rows:
        ws.cell(r, 1, label)
        ws.cell(r, 1).font = bold_font
        ws.cell(r, 1).fill = light_fill
        ws.cell(r, 2, value)
        ws.cell(r, 1).border = border
        ws.cell(r, 2).border = border
        if "SAR" in label:
            ws.cell(r, 2).number_format = '#,##0.00'
        r += 1

    r += 1
    ws.cell(r, 1, "Calculation Rules")
    ws.cell(r, 1).fill = dark_fill
    ws.cell(r, 1).font = white_font
    r += 1

    rule_rows = [
        ("Weekday definition", "Sunday–Thursday"),
        ("Weekend definition", "Friday–Saturday"),
        (
            "Weekday formula",
            f"Basic Salary × {rules['weekday_factor']} ÷ {rules['salary_divisor']}",
        ),
        (
            "Weekend formula",
            f"Basic Salary × {rules['weekend_factor']} ÷ {rules['salary_divisor']}",
        ),
        (
            "Transportation formula",
            f"Transportation Allowance ÷ {rules['transport_divisor']}",
        ),
        ("Public holidays", "Not automatically excluded"),
    ]

    for label, value in rule_rows:
        ws.cell(r, 1, label)
        ws.cell(r, 2, value)
        ws.cell(r, 1).border = border
        ws.cell(r, 2).border = border
        r += 1

    # Generic dataframe writer
    def add_dataframe_sheet(name, df):
        sheet = wb.create_sheet(name)

        for c, header in enumerate(df.columns, 1):
            cell = sheet.cell(1, c, header)
            cell.fill = dark_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center")
            cell.border = border

        for ri, row in enumerate(df.itertuples(index=False), 2):
            for ci, value in enumerate(row, 1):
                cell = sheet.cell(ri, ci, value)
                cell.border = border

                header = df.columns[ci - 1]

                if header == "Date" or "Date" in header:
                    if hasattr(value, "isoformat"):
                        cell.number_format = "yyyy-mm-dd"

                if "SAR" in header:
                    cell.number_format = '#,##0.00'

        sheet.freeze_panes = "A2"

        for col in range(1, sheet.max_column + 1):
            max_len = 12

            for rr in range(1, sheet.max_row + 1):
                value = sheet.cell(rr, col).value
                if value is not None:
                    max_len = max(max_len, min(35, len(str(value)) + 2))

            sheet.column_dimensions[get_column_letter(col)].width = max_len

    add_dataframe_sheet("Monthly Summary", monthly_df)
    add_dataframe_sheet("Daily Detail", daily_df)

    input_rows = []

    for d, amount in normalize_history(
        salary_df, "Effective Date", "Basic Salary (SAR)"
    ):
        input_rows.append({
            "Category": "Basic Salary",
            "Type": "",
            "Start / Effective Date": d,
            "End Date": None,
            "Amount (SAR)": float(amount),
        })

    for d, amount in normalize_history(
        transport_df, "Effective Date", "Transportation Allowance (SAR)"
    ):
        input_rows.append({
            "Category": "Transportation",
            "Type": "",
            "Start / Effective Date": d,
            "End Date": None,
            "Amount (SAR)": float(amount),
        })

    for typ, start, end in normalize_exclusions(exclusion_df):
        input_rows.append({
            "Category": "Exclusion",
            "Type": typ,
            "Start / Effective Date": start,
            "End Date": end,
            "Amount (SAR)": None,
        })

    inputs_df = pd.DataFrame(
        input_rows,
        columns=[
            "Category",
            "Type",
            "Start / Effective Date",
            "End Date",
            "Amount (SAR)",
        ],
    )

    add_dataframe_sheet("Inputs", inputs_df)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ============================================================
# SIDEBAR - CONTROLLED PROJECT LOAD
# ============================================================
with st.sidebar:
    st.header("Project")

    uploaded_file = st.file_uploader(
        "Choose saved project",
        type=["json"],
        key="project_uploader",
        help="Choose a project file first. Nothing changes until you click Load Project.",
    )

    load_clicked = st.button(
        "Load Project",
        type="primary",
        use_container_width=True,
        disabled=(uploaded_file is None),
    )

    if load_clicked and uploaded_file is not None:
        try:
            uploaded_file.seek(0)
            project = json.load(uploaded_file)
            apply_project(project)
            st.rerun()
        except Exception as exc:
            st.error(f"Could not load project: {exc}")

    if st.session_state.load_message:
        st.success(st.session_state.load_message)

    if st.button("New / Clear Project", use_container_width=True):
        st.session_state.period_start = date(date.today().year, 1, 1)
        st.session_state.period_end = date.today()
        st.session_state.salary_data = pd.DataFrame(
            {
                "Effective Date": [date(date.today().year, 1, 1)],
                "Basic Salary (SAR)": [0.0],
            }
        )
        st.session_state.transport_data = pd.DataFrame(
            {
                "Effective Date": [date(date.today().year, 1, 1)],
                "Transportation Allowance (SAR)": [0.0],
            }
        )
        st.session_state.exclusion_data = empty_exclusions()

        for k, v in DEFAULT_RULES.items():
            st.session_state[k] = v

        st.session_state.editor_generation += 1
        st.session_state.calc_result = None
        st.session_state.load_message = ""
        st.rerun()

    st.divider()
    st.subheader("Calculation Rules")

    weekday_factor = st.number_input(
        "Weekday factor",
        min_value=0.0,
        value=float(st.session_state.weekday_factor),
        step=0.1,
        key=f"weekday_factor_widget_{st.session_state.editor_generation}",
    )

    weekend_factor = st.number_input(
        "Weekend factor",
        min_value=0.0,
        value=float(st.session_state.weekend_factor),
        step=0.1,
        key=f"weekend_factor_widget_{st.session_state.editor_generation}",
    )

    salary_divisor = st.number_input(
        "Salary divisor",
        min_value=0.01,
        value=float(st.session_state.salary_divisor),
        step=1.0,
        key=f"salary_divisor_widget_{st.session_state.editor_generation}",
    )

    transport_divisor = st.number_input(
        "Transportation divisor",
        min_value=0.01,
        value=float(st.session_state.transport_divisor),
        step=1.0,
        key=f"transport_divisor_widget_{st.session_state.editor_generation}",
    )

    st.session_state.weekday_factor = weekday_factor
    st.session_state.weekend_factor = weekend_factor
    st.session_state.salary_divisor = salary_divisor
    st.session_state.transport_divisor = transport_divisor

    st.caption(
        "Weekdays: Sunday–Thursday\n\n"
        "Weekends: Friday–Saturday"
    )

# ============================================================
# MAIN UI
# ============================================================
st.title("🧮 Overtime Calculator")
st.caption(
    "Daily overtime calculation with salary/allowance history, "
    "Vacation/Intedab exclusions, reusable project files, and Excel reports."
)

input_tab, result_tab, help_tab = st.tabs(
    ["Inputs", "Results", "How It Works"]
)

generation = st.session_state.editor_generation

with input_tab:
    st.subheader("1. Overtime Period")

    p1, p2 = st.columns(2)

    period_start = p1.date_input(
        "Start date",
        value=st.session_state.period_start,
        key=f"period_start_widget_{generation}",
    )

    period_end = p2.date_input(
        "End date",
        value=st.session_state.period_end,
        key=f"period_end_widget_{generation}",
    )

    st.session_state.period_start = period_start
    st.session_state.period_end = period_end

    st.caption(
        "Every date in this range is counted as overtime unless it falls inside "
        "a Vacation or Intedab exclusion."
    )

    st.subheader("2. Basic Salary History")
    st.caption(
        "Add each salary change with the exact date it became effective. "
        "The app uses the correct salary separately for each day."
    )

    salary_df = st.data_editor(
        ensure_columns(st.session_state.salary_data, SALARY_COLUMNS),
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        key=f"salary_editor_{generation}",
        column_config={
            "Effective Date": st.column_config.DateColumn(
                "Effective Date",
                format="YYYY-MM-DD",
                required=True,
            ),
            "Basic Salary (SAR)": st.column_config.NumberColumn(
                "Basic Salary (SAR)",
                min_value=0.0,
                step=100.0,
                format="%.2f",
                required=True,
            ),
        },
    )

    st.session_state.salary_data = ensure_columns(
        salary_df, SALARY_COLUMNS
    )

    st.subheader("3. Transportation Allowance History")
    st.caption(
        "Add each transportation allowance change with its exact effective date."
    )

    transport_df = st.data_editor(
        ensure_columns(st.session_state.transport_data, TRANSPORT_COLUMNS),
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        key=f"transport_editor_{generation}",
        column_config={
            "Effective Date": st.column_config.DateColumn(
                "Effective Date",
                format="YYYY-MM-DD",
                required=True,
            ),
            "Transportation Allowance (SAR)": st.column_config.NumberColumn(
                "Transportation Allowance (SAR)",
                min_value=0.0,
                step=50.0,
                format="%.2f",
                required=True,
            ),
        },
    )

    st.session_state.transport_data = ensure_columns(
        transport_df, TRANSPORT_COLUMNS
    )

    st.subheader("4. Excluded Periods")
    st.caption(
        "Add Vacation or Intedab ranges. Public holidays are not excluded automatically."
    )

    exclusion_df = st.data_editor(
        ensure_columns(st.session_state.exclusion_data, EXCLUSION_COLUMNS),
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        key=f"exclusion_editor_{generation}",
        column_config={
            "Type": st.column_config.SelectboxColumn(
                "Type",
                options=["Vacation", "Intedab"],
                required=True,
            ),
            "Start Date": st.column_config.DateColumn(
                "Start Date",
                format="YYYY-MM-DD",
                required=True,
            ),
            "End Date": st.column_config.DateColumn(
                "End Date",
                format="YYYY-MM-DD",
                required=True,
            ),
        },
    )

    st.session_state.exclusion_data = ensure_columns(
        exclusion_df, EXCLUSION_COLUMNS
    )

    current_rules = {
        "weekday_factor": float(st.session_state.weekday_factor),
        "weekend_factor": float(st.session_state.weekend_factor),
        "salary_divisor": float(st.session_state.salary_divisor),
        "transport_divisor": float(st.session_state.transport_divisor),
    }

    st.divider()

    action1, action2 = st.columns([1, 1])

    project_file_bytes = build_project_json(
        st.session_state.period_start,
        st.session_state.period_end,
        st.session_state.salary_data,
        st.session_state.transport_data,
        st.session_state.exclusion_data,
        current_rules,
    )

    action1.download_button(
        "💾 Save Project",
        data=project_file_bytes,
        file_name="overtime_project.json",
        mime="application/json",
        use_container_width=True,
    )

    calculate_clicked = action2.button(
        "Calculate Overtime",
        type="primary",
        use_container_width=True,
    )

    if calculate_clicked:
        try:
            daily_df, monthly_df, summary = calculate_overtime(
                st.session_state.period_start,
                st.session_state.period_end,
                st.session_state.salary_data,
                st.session_state.transport_data,
                st.session_state.exclusion_data,
                current_rules,
            )

            st.session_state.calc_result = {
                "daily_df": daily_df,
                "monthly_df": monthly_df,
                "summary": summary,
                "period_start": st.session_state.period_start,
                "period_end": st.session_state.period_end,
                "salary_df": st.session_state.salary_data.copy(),
                "transport_df": st.session_state.transport_data.copy(),
                "exclusion_df": st.session_state.exclusion_data.copy(),
                "rules": current_rules.copy(),
            }

            st.success(
                "Calculation completed. Open the Results tab."
            )
        except Exception as exc:
            st.error(str(exc))

with result_tab:
    calculation = st.session_state.calc_result

    if calculation is None:
        st.info(
            "No calculation yet. Enter your information in the Inputs tab, "
            "then click Calculate Overtime."
        )
    else:
        summary = calculation["summary"]
        daily_df = calculation["daily_df"]
        monthly_df = calculation["monthly_df"]

        st.subheader("Calculation Summary")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Overtime", f"{summary['total']:,.2f} SAR")
        m2.metric("Weekday OT Days", summary["weekday_days"])
        m3.metric("Weekend OT Days", summary["weekend_days"])
        m4.metric("Excluded Days", summary["excluded_days"])

        m5, m6, m7 = st.columns(3)
        m5.metric("Weekday OT", f"{summary['weekday_ot']:,.2f} SAR")
        m6.metric("Weekend OT", f"{summary['weekend_ot']:,.2f} SAR")
        m7.metric("Transportation OT", f"{summary['transport_ot']:,.2f} SAR")

        st.subheader("Monthly Summary")
        st.dataframe(
            monthly_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Weekday OT (SAR)": st.column_config.NumberColumn(format="%.2f"),
                "Weekend OT (SAR)": st.column_config.NumberColumn(format="%.2f"),
                "Transportation OT (SAR)": st.column_config.NumberColumn(format="%.2f"),
                "Monthly Total (SAR)": st.column_config.NumberColumn(format="%.2f"),
            },
        )

        st.subheader("Daily Audit Trail")
        st.dataframe(
            daily_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Basic Salary (SAR)": st.column_config.NumberColumn(format="%.2f"),
                "Transportation Allowance (SAR)": st.column_config.NumberColumn(format="%.2f"),
                "Base OT (SAR)": st.column_config.NumberColumn(format="%.2f"),
                "Transportation OT (SAR)": st.column_config.NumberColumn(format="%.2f"),
                "Daily Total (SAR)": st.column_config.NumberColumn(format="%.2f"),
            },
        )

        st.subheader("Download Results")

        excel_data = make_excel_report(
            calculation["daily_df"],
            calculation["monthly_df"],
            calculation["summary"],
            calculation["period_start"],
            calculation["period_end"],
            calculation["salary_df"],
            calculation["transport_df"],
            calculation["exclusion_df"],
            calculation["rules"],
        )

        d1, d2, d3 = st.columns(3)

        d1.download_button(
            "Download Excel Report",
            data=excel_data,
            file_name="Overtime_Calculation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        d2.download_button(
            "Download Daily CSV",
            data=daily_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="overtime_daily_detail.csv",
            mime="text/csv",
            use_container_width=True,
        )

        d3.download_button(
            "Download Monthly CSV",
            data=monthly_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="overtime_monthly_summary.csv",
            mime="text/csv",
            use_container_width=True,
        )

with help_tab:
    st.subheader("Calculation Rules")

    st.markdown(
        f"""
**Weekday overtime — Sunday through Thursday**

`Basic Salary × {st.session_state.weekday_factor} ÷ {st.session_state.salary_divisor}`

**Weekend overtime — Friday and Saturday**

`Basic Salary × {st.session_state.weekend_factor} ÷ {st.session_state.salary_divisor}`

**Transportation overtime**

`Transportation Allowance ÷ {st.session_state.transport_divisor}`

### Effective dates
The app checks each overtime date individually and uses the latest Basic Salary
and Transportation Allowance whose effective date is on or before that day.

### Excluded periods
Vacation and Intedab dates are completely excluded from overtime.

### Public holidays
Public holidays are not automatically removed. A public holiday is still classified
by its normal calendar day unless you manually add it as an exclusion.
"""
    )

    st.subheader("Saving and Loading")
    st.markdown(
        """
1. Enter or update your information.
2. Click **Save Project** in the Inputs tab.
3. Keep the downloaded `.json` file.
4. Later, choose that file in the left sidebar.
5. Click **Load Project**.
6. The app replaces all current inputs with the saved project in one operation.

Uploading a file by itself does **not** change anything. This is intentional to prevent
accidental resets and repeated Streamlit reruns.
"""
    )

st.caption(
    "The app does not intentionally store your salary or overtime information in a database. "
    "Project data is only retained in the current Streamlit session unless you save it as a JSON file."
)
