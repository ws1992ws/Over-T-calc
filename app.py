
import streamlit as st
import pandas as pd
from datetime import date, timedelta

st.set_page_config(
    page_title="Overtime Calculator",
    page_icon="🧮",
    layout="wide"
)

# -----------------------------
# Default calculation constants
# -----------------------------
DEFAULT_WEEKDAY_FACTOR = 3.5
DEFAULT_WEEKEND_FACTOR = 7.0
DEFAULT_DIVISOR = 155.0
DEFAULT_TRANSPORT_DIVISOR = 30.0

st.title("🧮 Overtime Calculator")
st.caption(
    "Daily overtime calculator with salary/allowance history, Friday–Saturday weekends, "
    "Vacation/Intedab exclusions, and detailed results."
)

with st.expander("Calculation Rules", expanded=False):
    st.markdown("""
**Default rules**

- Weekday: **Sunday–Thursday**
- Weekend: **Friday–Saturday**
- Weekday OT = `Basic Salary × 3.5 ÷ 155`
- Weekend OT = `Basic Salary × 7 ÷ 155`
- Transportation OT = `Transportation Allowance ÷ 30`
- Salary and transportation are selected separately for **each day** based on their effective dates.
- Vacation and Intedab periods are excluded.
- Public holidays are **not** automatically excluded.
    """)

    c1, c2, c3, c4 = st.columns(4)
    weekday_factor = c1.number_input("Weekday factor", value=DEFAULT_WEEKDAY_FACTOR, step=0.1)
    weekend_factor = c2.number_input("Weekend factor", value=DEFAULT_WEEKEND_FACTOR, step=0.1)
    salary_divisor = c3.number_input("Salary divisor", value=DEFAULT_DIVISOR, step=1.0)
    transport_divisor = c4.number_input("Transportation divisor", value=DEFAULT_TRANSPORT_DIVISOR, step=1.0)

st.divider()

# -----------------------------
# Calculation period
# -----------------------------
st.subheader("1. Overtime Period")

pc1, pc2 = st.columns(2)
start_date = pc1.date_input("Start date", value=date(date.today().year, 1, 1))
end_date = pc2.date_input("End date", value=date.today())

st.caption(
    "Every date in this period is treated as an overtime day unless it falls inside an excluded Vacation or Intedab period."
)

# -----------------------------
# Salary history
# -----------------------------
st.subheader("2. Basic Salary History")
st.caption(
    "Add every salary change with the date from which it became effective. "
    "The calculator applies the correct salary to each overtime date."
)

default_salary = pd.DataFrame(
    {
        "Effective Date": [date(start_date.year, 1, 1)],
        "Basic Salary (SAR)": [0.0],
    }
)

salary_df = st.data_editor(
    default_salary,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Effective Date": st.column_config.DateColumn(
            "Effective Date",
            format="YYYY-MM-DD",
        ),
        "Basic Salary (SAR)": st.column_config.NumberColumn(
            "Basic Salary (SAR)",
            min_value=0.0,
            step=100.0,
            format="%.2f",
        ),
    },
    key="salary_editor",
)

# -----------------------------
# Transport history
# -----------------------------
st.subheader("3. Transportation Allowance History")
st.caption(
    "Add every transportation allowance change with its effective date. "
    "The applicable allowance is calculated separately for each day."
)

default_transport = pd.DataFrame(
    {
        "Effective Date": [date(start_date.year, 1, 1)],
        "Transportation Allowance (SAR)": [0.0],
    }
)

transport_df = st.data_editor(
    default_transport,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Effective Date": st.column_config.DateColumn(
            "Effective Date",
            format="YYYY-MM-DD",
        ),
        "Transportation Allowance (SAR)": st.column_config.NumberColumn(
            "Transportation Allowance (SAR)",
            min_value=0.0,
            step=50.0,
            format="%.2f",
        ),
    },
    key="transport_editor",
)

# -----------------------------
# Exclusions
# -----------------------------
st.subheader("4. Excluded Periods")
st.caption(
    "Add Vacation or Intedab periods to remove them from overtime. Public holidays stay included unless you add them here manually."
)

default_exclusions = pd.DataFrame(
    {
        "Type": pd.Series(dtype="str"),
        "Start Date": pd.Series(dtype="datetime64[ns]"),
        "End Date": pd.Series(dtype="datetime64[ns]"),
    }
)

exclusion_df = st.data_editor(
    default_exclusions,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Type": st.column_config.SelectboxColumn(
            "Type",
            options=["Vacation", "Intedab"],
        ),
        "Start Date": st.column_config.DateColumn(
            "Start Date",
            format="YYYY-MM-DD",
        ),
        "End Date": st.column_config.DateColumn(
            "End Date",
            format="YYYY-MM-DD",
        ),
    },
    key="exclusion_editor",
)

# -----------------------------
# Helpers
# -----------------------------
def normalize_history(df, date_col, amount_col):
    clean = df.copy()
    clean = clean.dropna(subset=[date_col, amount_col])
    if clean.empty:
        return []
    clean[date_col] = pd.to_datetime(clean[date_col]).dt.date
    clean[amount_col] = pd.to_numeric(clean[amount_col], errors="coerce")
    clean = clean.dropna(subset=[amount_col])
    clean = clean.sort_values(date_col)
    # If same effective date appears more than once, keep the last entered row.
    clean = clean.drop_duplicates(subset=[date_col], keep="last")
    return list(zip(clean[date_col], clean[amount_col]))

def effective_value(history, day):
    applicable = [x for x in history if x[0] <= day]
    if not applicable:
        return None
    return max(applicable, key=lambda x: x[0])[1]

def normalize_exclusions(df):
    if df.empty:
        return []
    clean = df.dropna(subset=["Type", "Start Date", "End Date"]).copy()
    if clean.empty:
        return []
    clean["Start Date"] = pd.to_datetime(clean["Start Date"]).dt.date
    clean["End Date"] = pd.to_datetime(clean["End Date"]).dt.date
    rows = []
    for _, r in clean.iterrows():
        if r["End Date"] < r["Start Date"]:
            raise ValueError(
                f"{r['Type']} exclusion has an end date before the start date."
            )
        rows.append((r["Type"], r["Start Date"], r["End Date"]))
    return rows

def exclusion_reason(day, exclusions):
    for typ, start, end in exclusions:
        if start <= day <= end:
            return typ
    return None

# -----------------------------
# Calculate
# -----------------------------
st.divider()

if st.button("Calculate Overtime", type="primary", use_container_width=True):
    try:
        if end_date < start_date:
            raise ValueError("End date cannot be before start date.")

        salary_history = normalize_history(
            salary_df, "Effective Date", "Basic Salary (SAR)"
        )
        transport_history = normalize_history(
            transport_df,
            "Effective Date",
            "Transportation Allowance (SAR)",
        )
        exclusions = normalize_exclusions(exclusion_df)

        if not salary_history:
            raise ValueError("Enter at least one Basic Salary.")
        if not transport_history:
            raise ValueError("Enter at least one Transportation Allowance.")

        rows = []
        d = start_date

        while d <= end_date:
            reason = exclusion_reason(d, exclusions)

            if reason:
                rows.append(
                    {
                        "Date": d,
                        "Day": d.strftime("%A"),
                        "Status": f"Excluded - {reason}",
                        "Basic Salary (SAR)": None,
                        "Transportation Allowance (SAR)": None,
                        "OT Type": "Excluded",
                        "Base OT (SAR)": 0.0,
                        "Transportation OT (SAR)": 0.0,
                        "Daily Total (SAR)": 0.0,
                    }
                )
            else:
                basic = effective_value(salary_history, d)
                transport = effective_value(transport_history, d)

                if basic is None:
                    raise ValueError(
                        f"No Basic Salary is effective on {d.isoformat()}. "
                        "Add a salary with an effective date on or before this date."
                    )
                if transport is None:
                    raise ValueError(
                        f"No Transportation Allowance is effective on {d.isoformat()}. "
                        "Add an allowance with an effective date on or before this date."
                    )

                is_weekend = d.weekday() in (4, 5)  # Friday, Saturday

                if is_weekend:
                    ot_type = "Weekend"
                    base_ot = basic * weekend_factor / salary_divisor
                else:
                    ot_type = "Weekday"
                    base_ot = basic * weekday_factor / salary_divisor

                transport_ot = transport / transport_divisor
                total = base_ot + transport_ot

                rows.append(
                    {
                        "Date": d,
                        "Day": d.strftime("%A"),
                        "Status": "Included",
                        "Basic Salary (SAR)": basic,
                        "Transportation Allowance (SAR)": transport,
                        "OT Type": ot_type,
                        "Base OT (SAR)": base_ot,
                        "Transportation OT (SAR)": transport_ot,
                        "Daily Total (SAR)": total,
                    }
                )

            d += timedelta(days=1)

        result_df = pd.DataFrame(rows)
        included_df = result_df[result_df["Status"] == "Included"].copy()

        total_ot = included_df["Daily Total (SAR)"].sum()
        weekday_days = (included_df["OT Type"] == "Weekday").sum()
        weekend_days = (included_df["OT Type"] == "Weekend").sum()
        excluded_days = len(result_df) - len(included_df)

        st.success("Calculation completed.")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Overtime", f"{total_ot:,.2f} SAR")
        m2.metric("Weekday OT Days", int(weekday_days))
        m3.metric("Weekend OT Days", int(weekend_days))
        m4.metric("Excluded Days", int(excluded_days))

        st.subheader("Monthly Summary")

        if not included_df.empty:
            monthly = included_df.copy()
            monthly["Month"] = pd.to_datetime(monthly["Date"]).dt.to_period("M").astype(str)

            monthly_summary = monthly.groupby("Month").apply(
                lambda g: pd.Series(
                    {
                        "Weekday Days": int((g["OT Type"] == "Weekday").sum()),
                        "Weekend Days": int((g["OT Type"] == "Weekend").sum()),
                        "Weekday OT (SAR)": g.loc[g["OT Type"] == "Weekday", "Base OT (SAR)"].sum(),
                        "Weekend OT (SAR)": g.loc[g["OT Type"] == "Weekend", "Base OT (SAR)"].sum(),
                        "Transportation OT (SAR)": g["Transportation OT (SAR)"].sum(),
                        "Monthly Total (SAR)": g["Daily Total (SAR)"].sum(),
                    }
                ),
                include_groups=False,
            ).reset_index()

            st.dataframe(
                monthly_summary,
                use_container_width=True,
                hide_index=True,
            )
        else:
            monthly_summary = pd.DataFrame()
            st.info("No included overtime days remain after exclusions.")

        st.subheader("Daily Calculation Detail")
        display_df = result_df.copy()

        money_cols = [
            "Basic Salary (SAR)",
            "Transportation Allowance (SAR)",
            "Base OT (SAR)",
            "Transportation OT (SAR)",
            "Daily Total (SAR)",
        ]

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                c: st.column_config.NumberColumn(c, format="%.2f")
                for c in money_cols
            },
        )

        csv_bytes = result_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Download Daily Detail CSV",
            data=csv_bytes,
            file_name="overtime_daily_detail.csv",
            mime="text/csv",
            use_container_width=True,
        )

        if not monthly_summary.empty:
            monthly_csv = monthly_summary.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "Download Monthly Summary CSV",
                data=monthly_csv,
                file_name="overtime_monthly_summary.csv",
                mime="text/csv",
                use_container_width=True,
            )

    except Exception as exc:
        st.error(str(exc))

st.divider()
st.caption(
    "This application does not intentionally save your entered salary or overtime data. "
    "Data exists only during the active Streamlit session unless you download the result files."
)
