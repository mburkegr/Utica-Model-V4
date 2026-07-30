import os
from datetime import date
from io import BytesIO

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import base64

from model import (
    run_deal_model,
    load_type_curve_library,
    load_price_file,
    prepare_slot_inputs,
    prepare_deal_settings,
    prepare_global_assumptions,
    calc_slot_metrics,
    build_slot_ngl_factors,
    run_single_slot_economics,
)

PRICE_FILE_PATH = "price_file_library.xlsx"

st.set_page_config(page_title="Utica Deal Model", layout="wide")


# ----------------------------- 
# Helpers
# -----------------------------
def pretty_column_name(col):
    name_map = {
        "date": "Date",
        "slot_id": "Slot ID",
        "tc_name": "Type Curve",
        "index_oil_price": "Index Oil Price",
        "index_gas_price": "Index Gas Price",
        "slot_net_oil_production": "Net Oil Production",
        "slot_net_gas_production": "Net Gas Production",
        "slot_net_ngl_production": "Net NGL Production",
        "slot_oil_revenue": "Net Oil Revenue",
        "slot_gas_revenue": "Net Gas Revenue",
        "slot_ngl_revenue": "Net NGL Revenue",
        "slot_total_revenue": "Total Revenue",
        "slot_loe": "Total LOE",
        "slot_tax": "Total Tax",
        "slot_operating_profit": "Operating Profit",
        "slot_capex": "Capex",
        "slot_asset_purchase": "Acquisition",
        "dale_promote": "WI Reversion Eligible",
        "pre_promote_working_interest": "Pre-Reversion WI",
        "promote_wi_transferred": "WI Transferred",
        "post_promote_working_interest": "Post-Reversion WI",
        "effective_working_interest": "Effective WI",
        "promote_running_multiple": "Reversion Multiple",
        "promote_hurdle_reached": "Reversion Hurdle Reached",
        "promote_active": "Reversion Active",
        "promote_hurdle_date": "Reversion Hurdle Date",
        "promote_effective_date": "Reversion Effective Date",
        "slot_total_cash_flow": "Total Cash Flow",
        "cum_total_cf": "Cumulative Total Cash Flow",
        "economic_limit_reached": "Economic Limit Reached",
        "well_shut_in": "Well Shut In",
        "pre_shut_in_operating_cf": "Pre-Shut-In Operating CF",
    }
    return name_map.get(col, col.replace("_", " ").title())


def is_effectively_zero(x, tol=1e-9):
    return pd.notnull(x) and abs(float(x)) < tol


def format_accounting_number(
    x,
    decimals=1,
    prefix="",
    suffix="",
    zero_as_dash=True,
    null_as_blank=True,
):
    if pd.isnull(x):
        return "" if null_as_blank else "-"

    x = float(x)

    if zero_as_dash and is_effectively_zero(x):
        return "-"

    abs_text = f"{abs(x):,.{decimals}f}"
    text = f"{prefix}{abs_text}{suffix}"

    return f"({text})" if x < 0 else text


def format_accounting_percent(
    x,
    decimals=0,
    zero_as_dash=True,
    null_as_blank=True,
):
    if pd.isnull(x):
        return "" if null_as_blank else "-"

    x = float(x)

    if zero_as_dash and is_effectively_zero(x):
        return "-"

    abs_text = f"{abs(x):.{decimals}%}"
    return f"({abs_text})" if x < 0 else abs_text


def format_accounting_production(
    x,
    large_decimals=0,
    small_decimals=2,
    threshold=10,
    zero_as_dash=True,
    null_as_blank=True,
):
    if pd.isnull(x):
        return "" if null_as_blank else "-"

    x = float(x)

    if zero_as_dash and is_effectively_zero(x):
        return "-"

    decimals = large_decimals if abs(x) >= threshold else small_decimals
    abs_text = f"{abs(x):,.{decimals}f}"
    return f"({abs_text})" if x < 0 else abs_text


def format_display_df(df):
    display_df = df.copy()

    percent_cols = {
        "working_interest",
        "pre_promote_working_interest",
        "promote_wi_transferred",
        "post_promote_working_interest",
        "effective_working_interest",
    }
    multiple_cols = {"promote_running_multiple"}

    for col in display_df.columns:
        if pd.api.types.is_datetime64_any_dtype(display_df[col]):
            display_df[col] = display_df[col].dt.strftime("%Y-%m-%d")
        elif pd.api.types.is_bool_dtype(display_df[col]):
            display_df[col] = display_df[col].map(
                lambda x: "Yes" if bool(x) else "No"
            )
        elif col in percent_cols:
            display_df[col] = display_df[col].map(
                lambda x: format_accounting_percent(
                    x,
                    decimals=2,
                    zero_as_dash=False,
                )
            )
        elif col in multiple_cols:
            display_df[col] = display_df[col].map(
                lambda x: format_accounting_number(
                    x,
                    decimals=2,
                    suffix="x",
                    zero_as_dash=False,
                )
            )
        elif pd.api.types.is_numeric_dtype(display_df[col]):
            display_df[col] = display_df[col].map(
                lambda x: format_accounting_number(x, decimals=1)
            )

    display_df.columns = [pretty_column_name(col) for col in display_df.columns]
    return display_df


def format_thousands_short(
    x,
    decimals=1,
    prefix="$",
    suffix="k",
    zero_as_dash=True,
    null_as_blank=True,
):
    if pd.isnull(x):
        return "" if null_as_blank else "-"

    x = float(x)

    if zero_as_dash and is_effectively_zero(x):
        return "-"

    x_thousands = x / 1000.0
    abs_text = f"{abs(x_thousands):,.{decimals}f}"
    text = f"{prefix}{abs_text}{suffix}"

    return f"({text})" if x < 0 else text


QUARTERLY_HEADER_COLOR = "#4E80B1"
BUTTON_DARK = "#2E4D6A"
MONTHLY_BTN = "#C0D4E4"
YEAR_FILL = "#CADEEE"


def inject_app_css():
    st.markdown(
        f"""
        <style>
        div[data-testid="stButton"] button[kind="primary"] {{
            background-color: {BUTTON_DARK} !important;
            color: white !important;
            border: 1px solid {BUTTON_DARK} !important;
            font-weight: 700 !important;
            border-radius: 10px !important;
        }}

        div[data-testid="stButton"] button[kind="primary"]:hover,
        div[data-testid="stButton"] button[kind="primary"]:focus,
        div[data-testid="stButton"] button[kind="primary"]:active {{
            background-color: {BUTTON_DARK} !important;
            color: white !important;
            border: 1px solid {BUTTON_DARK} !important;
            box-shadow: none !important;
            filter: brightness(1.05) !important;
        }}

        div[data-testid="stButton"] button[kind="secondary"] {{
            background-color: {MONTHLY_BTN} !important;
            color: #1f2d3d !important;
            border: 1px solid {MONTHLY_BTN} !important;
            font-weight: 700 !important;
            border-radius: 10px !important;
        }}

        div[data-testid="stButton"] button[kind="secondary"]:hover,
        div[data-testid="stButton"] button[kind="secondary"]:focus,
        div[data-testid="stButton"] button[kind="secondary"]:active {{
            background-color: {MONTHLY_BTN} !important;
            color: #1f2d3d !important;
            border: 1px solid {MONTHLY_BTN} !important;
            box-shadow: none !important;
            filter: brightness(1.05) !important;
        }}

        div[data-testid="stDownloadButton"] button {{
            background-color: {MONTHLY_BTN} !important;
            color: #1f2d3d !important;
            border: 1px solid {MONTHLY_BTN} !important;
            font-weight: 700 !important;
            border-radius: 10px !important;
        }}

        div[data-testid="stDownloadButton"] button:hover {{
            background-color: {MONTHLY_BTN} !important;
            filter: brightness(1.03) !important;
        }}

        div[data-testid="stFormSubmitButton"] button {{
            background-color: #2E4D6A !important;
            color: white !important;
            border: 1px solid #2E4D6A !important;
            font-weight: 700 !important;
            border-radius: 10px !important;
        }}
        
        div[data-testid="stFormSubmitButton"] button:hover {{
            filter: brightness(1.05) !important;
        }}

        div[data-testid="stExpander"] summary {{
            background-color: {QUARTERLY_HEADER_COLOR} !important;
            color: white !important;
            border-radius: 10px !important;
            font-weight: 700 !important;
            padding: 8px 12px !important;
        }}

        div[data-testid="stDataEditor"] [role="columnheader"],
        div[data-testid="stDataEditor"] thead th {{
            background-color: {QUARTERLY_HEADER_COLOR} !important;
            color: white !important;
            font-weight: 700 !important;
        }}

        div[data-testid="stDataEditor"] [role="columnheader"] *,
        div[data-testid="stDataEditor"] thead th * {{
            color: white !important;
            fill: white !important;
            font-weight: 700 !important;
        }}

        div[data-testid="stDataEditor"] [role="gridcell"] {{
            border-color: #e6e6e6 !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


inject_app_css()
st.title("Utica Deal Model")


@st.cache_data(show_spinner=False)
def load_tc_names(file_mtime):
    tc_metadata = pd.read_excel("type_curve_library.xlsx", sheet_name="tc_metadata")
    tc_metadata["tc_name"] = tc_metadata["tc_name"].astype(str).str.strip()
    return tc_metadata["tc_name"].dropna().unique().tolist()


def next_month_start():
    today = date.today()
    if today.month == 12:
        return date(today.year + 1, 1, 1)
    return date(today.year, today.month + 1, 1)


def build_slot_template(num_slots):
    rows = []
    for i in range(1, num_slots + 1):
        rows.append(
            {
                "include_slot": True,
                "dale_promote": False,
                "carry_enabled": False,
                "carry_wi_reversion_pct": 0.0,
                "slot_id": i,
                "tc_name": "Choose TC",
                "gross_wells": 1.0,
                "net_acres": 25.0,
                "unit_acres": 200.0,
                "use_calc_unit_acres": False,
                "pct_unitized": 1.0,
                "drilling_spud_month": next_month_start(),
                "flowback_delay": 4,
                "net_revenue_interest": 0.80,
                "lateral_length": 10000,
                "dc_costs": 750.0,
                "tc_risk": 1.00,
                "bid_per_acre": 8000.0,
                "oil_diff": -10.00,
                "gas_diff": -3.00,
                "ngl_diff": 0.00,
                "oil_opex_bbl": 1.78,
                "gas_opex_mcf": 0.25,
                "ngl_opex": 2.50,
                "fixed_loe": 3534.0,
                "ngl_yield": 3.0,
            }
        )
    return pd.DataFrame(rows)


def resize_slot_df(existing_df, target_slots):
    existing_df = existing_df.copy().reset_index(drop=True)
    current_slots = len(existing_df)

    if current_slots == target_slots:
        existing_df["slot_id"] = range(1, target_slots + 1)
        return existing_df

    if current_slots < target_slots:
        new_rows = build_slot_template(target_slots).iloc[current_slots:].copy()
        existing_df = pd.concat([existing_df, new_rows], ignore_index=True)
        existing_df["slot_id"] = range(1, target_slots + 1)
        return existing_df

    trimmed_df = existing_df.iloc[:target_slots].copy().reset_index(drop=True)
    trimmed_df["slot_id"] = range(1, target_slots + 1)
    return trimmed_df


def to_excel_bytes(deal_df, slot_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        deal_df.to_excel(writer, index=False, sheet_name="Deal Audit")
        slot_df.to_excel(writer, index=False, sheet_name="Slot Audit")
    output.seek(0)
    return output.getvalue()

def build_deal_log_csv(opportunity_name, overview_df, model_slot_df):
    # This will be the very first section in the exported CSV.
    deal_name_df = pd.DataFrame(
        [{"Deal Name": str(opportunity_name).strip()}]
    )

    return (
        deal_name_df.to_csv(index=False)
        + "\nDEAL RUN OVERVIEW\n"
        + overview_df.to_csv(index=False)
        + "\nTC INPUTS USED FOR MODEL RUN\n"
        + model_slot_df.to_csv(index=False)
    )

def apply_calc_unit_acres(df):
    df = df.copy()
    mask = df["use_calc_unit_acres"].fillna(False)

    df.loc[mask, "unit_acres"] = (
        df.loc[mask, "lateral_length"] / 50.0 * df.loc[mask, "gross_wells"]
    )

    return df


def build_sensitivity_range(
    base_value,
    step,
    steps_each_way=3,
    min_value=None,
):
    """Build a centered sensitivity range while respecting an optional floor.

    If the lower values hit the floor, duplicate values are removed and the
    upper end is extended so the table keeps the requested number of points.
    The base value remains in the range.
    """
    target_count = (steps_each_way * 2) + 1
    raw_values = [
        float(base_value) + float(step) * i
        for i in range(-steps_each_way, steps_each_way + 1)
    ]

    if min_value is not None:
        raw_values = [max(float(min_value), value) for value in raw_values]

    values = []
    for value in raw_values:
        rounded = round(float(value), 10)
        if rounded not in values:
            values.append(rounded)

    while len(values) < target_count:
        next_value = max(values) + float(step)
        values.append(round(next_value, 10))

    return sorted(values)

def build_percentage_sensitivity_range(
    base_value,
    pct_step=0.05,
    steps_each_way=4,
    min_value=0.0,
):
    """Build a sensitivity range using percentage changes from the base.

    Example: a 5% step with four steps each way produces 80% through
    120% of the base value. Values are rounded to two decimals.
    """
    base_value = float(base_value)
    values = [
        max(
            float(min_value),
            base_value * (1.0 + float(pct_step) * i),
        )
        for i in range(-steps_each_way, steps_each_way + 1)
    ]
    return [round(value, 2) for value in values]


def weighted_avg_by_net_acres(slot_df, value_col, weight_col="net_acres"):
    """Return a net-acre weighted average for a slot-level input."""
    values = pd.to_numeric(slot_df[value_col], errors="coerce")
    weights = pd.to_numeric(slot_df[weight_col], errors="coerce").fillna(0.0)

    valid = values.notna() & weights.gt(0)

    if not valid.any():
        return float(values.mean()) if values.notna().any() else 0.0

    return float((values[valid] * weights[valid]).sum() / weights[valid].sum())


def weighted_avg_spud_month_by_net_acres(
    slot_df,
    date_col="drilling_spud_month",
    weight_col="net_acres",
):
    """Return a net-acre-weighted representative spud month.

    The sensitivity shifts every slot by the same month delta, so relative
    timing between slots is preserved. This weighted month is used only for
    the displayed center date and axis labels.
    """
    dates = pd.to_datetime(slot_df[date_col], errors="coerce")
    weights = pd.to_numeric(slot_df[weight_col], errors="coerce").fillna(0.0)

    valid = dates.notna() & weights.gt(0)
    if not valid.any():
        valid_dates = dates.dropna()
        if valid_dates.empty:
            return pd.Timestamp(next_month_start())
        return valid_dates.iloc[0].to_period("M").to_timestamp()

    month_numbers = (
        dates[valid].dt.year.astype(float) * 12.0
        + dates[valid].dt.month.astype(float)
        - 1.0
    )
    weighted_month_number = int(
        round((month_numbers * weights[valid]).sum() / weights[valid].sum())
    )

    year, month_index = divmod(weighted_month_number, 12)
    return pd.Timestamp(year=year, month=month_index + 1, day=1)


@st.cache_data(show_spinner=False)
def run_two_way_sensitivity(
    slot_df,
    deal_inputs,
    x_values,
    x_variable,
    y_values,
    y_variable,
):
    """Run a generic two-variable IRR and MOIC sensitivity table."""
    irr_table = pd.DataFrame(index=y_values, columns=x_values, dtype=float)
    moic_table = pd.DataFrame(index=y_values, columns=x_values, dtype=float)

    base_tc_risk = weighted_avg_by_net_acres(slot_df, "tc_risk")
    base_ngl_yield = weighted_avg_by_net_acres(slot_df, "ngl_yield")
    base_spud_month = weighted_avg_spud_month_by_net_acres(slot_df)

    def apply_value(sens_slot_df, sens_deal_inputs, variable, value):
        if variable == "spud_date":
            target_spud_month = pd.Timestamp(value).to_period("M").to_timestamp()
            month_delta = (
                (target_spud_month.year - base_spud_month.year) * 12
                + target_spud_month.month
                - base_spud_month.month
            )
            sens_slot_df["drilling_spud_month"] = (
                pd.to_datetime(
                    sens_slot_df["drilling_spud_month"],
                    errors="coerce",
                )
                + pd.DateOffset(months=int(month_delta))
            )
            return

        value = float(value)

        if variable == "bid":
            sens_deal_inputs["use_bid_override"] = True
            sens_deal_inputs["bid_override"] = max(1.0, value)

        elif variable == "dc":
            sens_deal_inputs["use_dc_override"] = True
            sens_deal_inputs["dc_override"] = value

        elif variable == "oil":
            sens_deal_inputs["oil_price"] = value

        elif variable == "gas":
            sens_deal_inputs["gas_price"] = value

        elif variable == "tc_risk":
            tc_risk_delta = value - float(base_tc_risk)
            sens_slot_df["tc_risk"] = (
                pd.to_numeric(
                    sens_slot_df["tc_risk"],
                    errors="coerce",
                ).fillna(0.0)
                + tc_risk_delta
            ).clip(lower=0.0)

        elif variable == "ngl_yield":
            ngl_delta = value - float(base_ngl_yield)
            sens_slot_df["ngl_yield"] = (
                pd.to_numeric(
                    sens_slot_df["ngl_yield"],
                    errors="coerce",
                ).fillna(0.0)
                + ngl_delta
            ).clip(lower=0.0)

        else:
            raise ValueError(f"Unsupported sensitivity variable: {variable}")

    for y_value in y_values:
        for x_value in x_values:
            sens_deal_inputs = deal_inputs.copy()
            sens_slot_df = slot_df.copy()

            apply_value(
                sens_slot_df,
                sens_deal_inputs,
                x_variable,
                x_value,
            )
            apply_value(
                sens_slot_df,
                sens_deal_inputs,
                y_variable,
                y_value,
            )

            try:
                _, _, _, _, irr, moic = run_deal_model(
                    sens_slot_df,
                    sens_deal_inputs,
                )
                irr_table.loc[y_value, x_value] = irr
                moic_table.loc[y_value, x_value] = moic
            except Exception:
                irr_table.loc[y_value, x_value] = None
                moic_table.loc[y_value, x_value] = None

    return irr_table, moic_table

@st.cache_data(show_spinner=False)
def run_individual_slot_returns(slot_df, deal_inputs):
    slot_returns = {}

    total_net_acres = float(slot_df["net_acres"].sum())

    for _, slot_row in slot_df.iterrows():
        slot_id = int(slot_row["slot_id"])

        one_slot_df = pd.DataFrame([slot_row]).copy()
        one_slot_df = one_slot_df.drop(columns=["include_slot"], errors="ignore")

        slot_deal_inputs = deal_inputs.copy()

        # If using a deal-level acquisition override, allocate it across slots
        # based on each slot's share of total net acres.
        if slot_deal_inputs.get("use_acquisition_override", False):
            total_override = float(slot_deal_inputs.get("acquisition_cost_override", 0.0))
            slot_net_acres = float(slot_row.get("net_acres", 0.0))

            allocated_override = (
                total_override * slot_net_acres / total_net_acres
                if total_net_acres > 0
                else 0.0
            )

            slot_deal_inputs["acquisition_cost_override"] = allocated_override

        try:
            _, _, _, _, slot_irr, slot_moic = run_deal_model(
                one_slot_df,
                slot_deal_inputs,
            )

            slot_returns[slot_id] = {
                "irr": slot_irr,
                "moic": slot_moic,
            }

        except Exception:
            slot_returns[slot_id] = {
                "irr": None,
                "moic": None,
            }

    return slot_returns

@st.cache_data(show_spinner=False)
def run_bid_dc_sensitivity(slot_df, deal_inputs, base_dc, base_bid):
    if bool(deal_inputs.get("use_dc_pct_sensitivity", False)):
        dc_values = build_percentage_sensitivity_range(
            base_dc,
            pct_step=0.05,
            steps_each_way=4,
            min_value=0.0,
        )
    else:
        dc_values = build_sensitivity_range(
            base_dc,
            50.0,
            4,
            min_value=0.0,
        )

    bid_values = build_sensitivity_range(
        base_bid,
        500.0,
        4,
        min_value=1.0,
    )

    irr_table = pd.DataFrame(index=bid_values, columns=dc_values)
    moic_table = pd.DataFrame(index=bid_values, columns=dc_values)

    for dc in dc_values:
        for bid in bid_values:
            sens_deal_inputs = deal_inputs.copy()
            sens_deal_inputs["use_dc_override"] = True
            sens_deal_inputs["dc_override"] = float(dc)
            sens_deal_inputs["use_bid_override"] = True
            sens_deal_inputs["bid_override"] = float(bid)

            try:
                _, _, _, _, irr, moic = run_deal_model(slot_df.copy(), sens_deal_inputs)
                irr_table.loc[bid, dc] = irr
                moic_table.loc[bid, dc] = moic
            except Exception:
                irr_table.loc[bid, dc] = None
                moic_table.loc[bid, dc] = None

    return irr_table, moic_table


@st.cache_data(show_spinner=False)
def run_oil_bid_sensitivity(slot_df, deal_inputs, oil_values, bid_values):
    irr_table = pd.DataFrame(index=bid_values, columns=oil_values, dtype=float)
    moic_table = pd.DataFrame(index=bid_values, columns=oil_values, dtype=float)

    for oil in oil_values:
        for bid in bid_values:
            sens_deal_inputs = deal_inputs.copy()
            sens_deal_inputs["oil_price"] = float(oil)
            sens_deal_inputs["use_bid_override"] = True
            sens_deal_inputs["bid_override"] = float(bid)

            try:
                _, _, _, _, irr, moic = run_deal_model(slot_df.copy(), sens_deal_inputs)
                irr_table.loc[bid, oil] = irr
                moic_table.loc[bid, oil] = moic
            except Exception:
                irr_table.loc[bid, oil] = None
                moic_table.loc[bid, oil] = None

    return irr_table, moic_table


@st.cache_data(show_spinner=False)
def run_gas_bid_sensitivity(slot_df, deal_inputs, gas_values, bid_values):
    irr_table = pd.DataFrame(index=bid_values, columns=gas_values, dtype=float)
    moic_table = pd.DataFrame(index=bid_values, columns=gas_values, dtype=float)

    for gas in gas_values:
        for bid in bid_values:
            sens_deal_inputs = deal_inputs.copy()
            sens_deal_inputs["gas_price"] = float(gas)
            sens_deal_inputs["use_bid_override"] = True
            sens_deal_inputs["bid_override"] = float(bid)

            try:
                _, _, _, _, irr, moic = run_deal_model(slot_df.copy(), sens_deal_inputs)
                irr_table.loc[bid, gas] = irr
                moic_table.loc[bid, gas] = moic
            except Exception:
                irr_table.loc[bid, gas] = None
                moic_table.loc[bid, gas] = None

    return irr_table, moic_table


@st.cache_data(show_spinner=False)
def run_tcrisk_bid_sensitivity(slot_df, deal_inputs, tc_risk_values, bid_values):
    irr_table = pd.DataFrame(index=bid_values, columns=tc_risk_values, dtype=float)
    moic_table = pd.DataFrame(index=bid_values, columns=tc_risk_values, dtype=float)

    # This is the displayed/base weighted-average TC risk.
    # At this exact value, delta = 0, so each slot keeps its actual original TC risk.
    base_tc_risk = weighted_avg_by_net_acres(slot_df, "tc_risk")

    for tc_risk in tc_risk_values:
        for bid in bid_values:
            sens_deal_inputs = deal_inputs.copy()
            sens_deal_inputs["use_bid_override"] = True
            sens_deal_inputs["bid_override"] = float(bid)

            sens_slot_df = slot_df.copy()

            # Apply sensitivity as +/- change from the base weighted average,
            # instead of setting every slot to the same absolute TC risk.
            tc_risk_delta = float(tc_risk) - float(base_tc_risk)

            sens_slot_df["tc_risk"] = (
                pd.to_numeric(sens_slot_df["tc_risk"], errors="coerce").fillna(0.0)
                + tc_risk_delta
            ).clip(lower=0.0)

            try:
                _, _, _, _, irr, moic = run_deal_model(sens_slot_df, sens_deal_inputs)
                irr_table.loc[bid, tc_risk] = irr
                moic_table.loc[bid, tc_risk] = moic
            except Exception:
                irr_table.loc[bid, tc_risk] = None
                moic_table.loc[bid, tc_risk] = None

    return irr_table, moic_table
    
@st.cache_data(show_spinner=False)
def run_ngl_yield_bid_sensitivity(slot_df, deal_inputs, ngl_yield_values, bid_values):
    irr_table = pd.DataFrame(index=bid_values, columns=ngl_yield_values, dtype=float)
    moic_table = pd.DataFrame(index=bid_values, columns=ngl_yield_values, dtype=float)

    # This is the displayed/base weighted-average NGL yield.
    # At this exact value, delta = 0, so each slot keeps its actual original NGL yield.
    base_ngl_yield = weighted_avg_by_net_acres(slot_df, "ngl_yield")

    for ngl_yield in ngl_yield_values:
        for bid in bid_values:
            sens_deal_inputs = deal_inputs.copy()
            sens_deal_inputs["use_bid_override"] = True
            sens_deal_inputs["bid_override"] = float(bid)

            sens_slot_df = slot_df.copy()

            # Apply sensitivity as +/- change from the base weighted average,
            # instead of setting every slot to the same absolute NGL yield.
            ngl_delta = float(ngl_yield) - float(base_ngl_yield)

            sens_slot_df["ngl_yield"] = (
                pd.to_numeric(sens_slot_df["ngl_yield"], errors="coerce").fillna(0.0)
                + ngl_delta
            ).clip(lower=0.0)

            try:
                _, _, _, _, irr, moic = run_deal_model(sens_slot_df, sens_deal_inputs)
                irr_table.loc[bid, ngl_yield] = irr
                moic_table.loc[bid, ngl_yield] = moic
            except Exception:
                irr_table.loc[bid, ngl_yield] = None
                moic_table.loc[bid, ngl_yield] = None

    return irr_table, moic_table

def build_heatmap(
    df,
    title,
    metric="irr",
    x_title="",
    y_title="",
    x_format="dollar",
    y_format="dollar",
    base_x=None,
    base_y=None,
):
    heatmap_df = df.copy()

    def format_axis_value(v, fmt):
        if fmt == "dollar":
            return f"${int(v):,}" if float(v).is_integer() else f"${v:,.2f}"
        if fmt == "percent":
            return f"{v:.0%}"
        if fmt == "float2":
            return f"{v:.2f}"
        if fmt == "date":
            return pd.Timestamp(v).strftime("%b-%y")
        return str(v)

    x_vals = [format_axis_value(x, x_format) for x in heatmap_df.columns]
    y_vals = [format_axis_value(y, y_format) for y in heatmap_df.index]

    def clamp01(x):
        return max(0.0, min(1.0, x))

    if metric == "irr":
        text_vals = heatmap_df.map(lambda x: f"{x:.2%}" if pd.notnull(x) else "")
        zmin = 0.0
        zmax = max(0.40, float(heatmap_df.max().max()))

        low_cut = 0.15
        high_cut = 0.25

        low_norm = clamp01((low_cut - zmin) / (zmax - zmin)) if zmax > zmin else 0.33
        high_norm = clamp01((high_cut - zmin) / (zmax - zmin)) if zmax > zmin else 0.66

        colorscale = [
            [0.00, "rgb(255,180,180)"],
            [low_norm, "rgb(255,180,180)"],
            [low_norm, "rgb(255,255,204)"],
            [high_norm, "rgb(255,255,204)"],
            [high_norm, "rgb(214,232,202)"],
            [1.00, "rgb(214,232,202)"],
        ]
    elif metric == "moic":
        text_vals = heatmap_df.map(lambda x: f"{x:.2f}x" if pd.notnull(x) else "")
        zmin = min(0.0, float(heatmap_df.min().min()))
        zmax = max(2.0, float(heatmap_df.max().max()))

        low_cut = 1.00
        high_cut = 1.50

        low_norm = clamp01((low_cut - zmin) / (zmax - zmin)) if zmax > zmin else 0.33
        high_norm = clamp01((high_cut - zmin) / (zmax - zmin)) if zmax > zmin else 0.66

        colorscale = [
            [0.00, "rgb(255,180,180)"],
            [low_norm, "rgb(255,180,180)"],
            [low_norm, "rgb(255,255,204)"],
            [high_norm, "rgb(255,255,204)"],
            [high_norm, "rgb(214,232,202)"],
            [1.00, "rgb(214,232,202)"],
        ]
    else:
        text_vals = heatmap_df.map(lambda x: f"{x}" if pd.notnull(x) else "")
        zmin = 0.0
        zmax = 1.0
        colorscale = "RdYlGn"

    fig = go.Figure(
        data=go.Heatmap(
            z=heatmap_df.values,
            x=x_vals,
            y=y_vals,
            text=text_vals.values,
            texttemplate="%{text}",
            colorscale=colorscale,
            zmin=zmin,
            zmax=zmax,
            showscale=False,
            hovertemplate=f"{x_title}: %{{x}}<br>{y_title}: %{{y}}<br>Value: %{{text}}<extra></extra>",
            textfont=dict(size=16),
        )
    )

    fig.update_layout(
        title=dict(
            text=title,
            x=0.0,
            xanchor="left",
            y=0.98,
            yanchor="top",
            font=dict(size=22, color="black"),
        ),
        font=dict(size=16, color="black"),
        xaxis=dict(
            title=dict(text=x_title, font=dict(size=18, color="black")),
            tickfont=dict(size=15, color="black"),
            side="top",
            type="category",
            automargin=True,
        ),
        yaxis=dict(
            title=dict(
                text=y_title,
                font=dict(size=18, color="black"),   # 🔥 slightly bigger
                standoff=20,                         # 🔥 THIS creates gap
            ),
            tickfont=dict(size=15, color="black"),   # 🔥 slightly bigger
            type="category",
            automargin=True,
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=90, r=20, t=90, b=60),
        height=390,
    )
    
    if base_x is not None and base_y is not None:
        try:
            x_vals_raw = list(heatmap_df.columns)
            y_vals_raw = list(heatmap_df.index)

            def find_closest_index(values, target):
                def distance(value):
                    if isinstance(value, (pd.Timestamp, date)) or isinstance(
                        target,
                        (pd.Timestamp, date),
                    ):
                        return abs(
                            (pd.Timestamp(value) - pd.Timestamp(target)).days
                        )
                    return abs(float(value) - float(target))

                return min(range(len(values)), key=lambda i: distance(values[i]))

            x_idx = find_closest_index(x_vals_raw, base_x)
            y_idx = find_closest_index(y_vals_raw, base_y)

            fig.add_shape(
                type="rect",
                x0=x_idx - 0.5,
                x1=x_idx + 0.5,
                y0=y_idx - 0.5,
                y1=y_idx + 0.5,
                line=dict(color="black", width=3),
                fillcolor="rgba(0,0,0,0)",
            )
        except Exception:
            pass

    return fig


def build_quarterly_output_table(deal_df, all_slots_df, slot_df, deal_inputs):
    import numpy as np

    deal = deal_df.copy()
    slots = all_slots_df.copy()
    slot_inputs = slot_df.copy()

    deal["date"] = pd.to_datetime(deal["date"])
    slots["date"] = pd.to_datetime(slots["date"])
    slot_inputs["drilling_spud_month"] = pd.to_datetime(slot_inputs["drilling_spud_month"])

    deal["quarter_label"] = "Q" + deal["date"].dt.quarter.astype(str) + " " + deal["date"].dt.strftime("%y")
    deal["year_label"] = deal["date"].dt.year.astype(str)

    quarter_order = [
        "Q1 26", "Q2 26", "Q3 26", "Q4 26",
        "Q1 27", "Q2 27", "Q3 27", "Q4 27",
    ]
    year_order = [str(y) for y in range(2026, 2034)]

    def quarter_days_from_label(q_label):
        q_num = int(q_label[1])
        year = 2000 + int(q_label[-2:])
        quarter_start_month = {1: 1, 2: 4, 3: 7, 4: 10}[q_num]
        start = pd.Timestamp(year=year, month=quarter_start_month, day=1)
        end = start + pd.offsets.QuarterEnd(0)
        return (end - start).days + 1

    def year_days_from_label(y_label):
        year = int(y_label)
        start = pd.Timestamp(year=year, month=1, day=1)
        end = pd.Timestamp(year=year, month=12, day=31)
        return (end - start).days + 1

    q_prices = (
    deal.groupby("quarter_label")[
        [
            "index_oil_price",
            "index_gas_price",
        ]
    ]
    .mean()
    .reindex(quarter_order)
    )
    
    y_prices = (
        deal.groupby("year_label")[
            [
                "index_oil_price",
                "index_gas_price",
            ]
        ]
        .mean()
        .reindex(year_order)
    )
    
    q_days = pd.Series({q: quarter_days_from_label(q) for q in quarter_order}, index=quarter_order, dtype=float)
    y_days = pd.Series({y: year_days_from_label(y) for y in year_order}, index=year_order, dtype=float)

    q = deal.groupby("quarter_label").sum(numeric_only=True).reindex(quarter_order)
    y = deal.groupby("year_label").sum(numeric_only=True).reindex(year_order)

    slot_metrics = slot_inputs.copy()
    slot_metrics["spud_quarter"] = "Q" + slot_metrics["drilling_spud_month"].dt.quarter.astype(str) + " " + slot_metrics["drilling_spud_month"].dt.strftime("%y")
    slot_metrics["spud_year"] = slot_metrics["drilling_spud_month"].dt.year.astype(str)

    # Gross wells do not change with a WI promote. Net wells spud should use
    # the actual effective ownership on each slot's spud date, so wells drilled
    # after the promote vests reflect the reduced WI.
    slot_metrics["gross_wells_spud"] = slot_metrics["gross_wells"]

    spud_effective_wi = (
        slots[["slot_id", "date", "effective_net_wells"]]
        .merge(
            slot_inputs[["slot_id", "drilling_spud_month"]],
            on="slot_id",
            how="inner",
        )
    )
    spud_effective_wi = spud_effective_wi[
        spud_effective_wi["date"] == spud_effective_wi["drilling_spud_month"]
    ]
    net_wells_by_slot = (
        spud_effective_wi.groupby("slot_id")["effective_net_wells"].first()
    )

    # Fallback for any unmatched row retains the original slot calculation.
    unit_acres_final = np.where(
        slot_metrics["use_calc_unit_acres"].fillna(False),
        slot_metrics["gross_wells"] * slot_metrics["lateral_length"] / 50.0,
        slot_metrics["unit_acres"],
    )
    fallback_working_interest = np.where(
        unit_acres_final != 0,
        (slot_metrics["net_acres"] / unit_acres_final) * slot_metrics["pct_unitized"],
        0.0,
    )
    fallback_net_wells = fallback_working_interest * slot_metrics["gross_wells"]

    slot_metrics["net_wells_spud"] = (
        slot_metrics["slot_id"].map(net_wells_by_slot)
        .fillna(pd.Series(fallback_net_wells, index=slot_metrics.index))
    )

    q_spud = slot_metrics.groupby("spud_quarter")[["gross_wells_spud", "net_wells_spud"]].sum().reindex(quarter_order).fillna(0.0)
    y_spud = slot_metrics.groupby("spud_year")[["gross_wells_spud", "net_wells_spud"]].sum().reindex(year_order).fillna(0.0)

    def safe_div(n, d):
        return np.where((d != 0) & pd.notnull(d), n / d, 0.0)

    def build_section(df, days, index_prices):
        out = pd.DataFrame(index=[], columns=df.index)

        oil_index_price = index_prices[
            "index_oil_price"
        ]
        
        gas_index_price = index_prices[
            "index_gas_price"
        ]

        realized_oil = safe_div(df["slot_oil_revenue"], df["slot_net_oil_production"])
        realized_gas = safe_div(df["slot_gas_revenue"], df["slot_net_gas_production"])
        realized_ngl_price = safe_div(df["slot_ngl_revenue"], df["slot_net_ngl_production"])
        realized_ngl_pct_wti = safe_div(
            realized_ngl_price,
            oil_index_price,
        )

        oil_mbbl_d = safe_div(df["slot_net_oil_production"], days)
        ngl_mbbl_d = safe_div(df["slot_net_ngl_production"], days)
        gas_mmcf_d = safe_div(df["slot_net_gas_production"], days)

        total_mcfe = (
            df["slot_net_oil_production"] * 6.0
            + df["slot_net_ngl_production"] * 6.0
            + df["slot_net_gas_production"]
        )
        total_mcfe_d = safe_div(total_mcfe, days)

        taxes_pos = -df["slot_tax"]
        loe_pos = -df["slot_loe"]
        total_opex = taxes_pos + loe_pos
        ebitda = df["slot_total_revenue"] - total_opex

        d_and_c = -df["slot_capex"]
        acquisition = -df["slot_asset_purchase"]
        total_capex = d_and_c + acquisition

        free_cash_flow = df["slot_total_cash_flow"]

        out.loc[
            "Assumed Index Pricing - Crude Oil"
        ] = oil_index_price
        
        out.loc[
            "Assumed Index Pricing - Natural Gas"
        ] = gas_index_price
        out.loc["Realized Pricing - Crude Oil"] = realized_oil
        out.loc["Realized Pricing - NGL (% of WTI)"] = realized_ngl_pct_wti
        out.loc["Realized Pricing - Natural Gas"] = realized_gas
        out.loc["Production - Crude Oil"] = oil_mbbl_d
        out.loc["Production - NGL's"] = ngl_mbbl_d
        out.loc["Production - Natural Gas"] = gas_mmcf_d
        out.loc["Production - Total (Mcfe/d)"] = total_mcfe_d
        out.loc["Revenues - Crude Oil"] = df["slot_oil_revenue"] / 1000.0
        out.loc["Revenues - NGL's"] = df["slot_ngl_revenue"] / 1000.0
        out.loc["Revenues - Natural Gas"] = df["slot_gas_revenue"] / 1000.0
        out.loc["Revenues - Total"] = df["slot_total_revenue"] / 1000.0
        out.loc["Operating Expenses - Taxes"] = taxes_pos / 1000.0
        out.loc["Operating Expenses - LOE"] = loe_pos / 1000.0
        out.loc["Operating Expenses - Total Opex"] = total_opex / 1000.0
        out.loc["Taxes / Mcfe"] = safe_div(taxes_pos, total_mcfe)
        out.loc["LOE / Mcfe"] = safe_div(loe_pos, total_mcfe)
        out.loc["EBITDA"] = ebitda / 1000.0
        out.loc["Capital Expenditures - D&C"] = d_and_c / 1000.0
        out.loc["Capital Expenditures - Acquisition"] = acquisition / 1000.0
        out.loc["Capital Expenditures - Total"] = total_capex / 1000.0
        out.loc["Free Cash Flow"] = free_cash_flow / 1000.0
        out.loc["Cumulative FCF"] = (free_cash_flow / 1000.0).cumsum()

        return out

    q_out = build_section(
        q,
        q_days,
        q_prices,
    )
    
    y_out = build_section(
        y,
        y_days,
        y_prices,
    )

    q_out.loc["Gross Wells Spud"] = q_spud["gross_wells_spud"]
    q_out.loc["Net Wells Spud"] = q_spud["net_wells_spud"]
    y_out.loc["Gross Wells Spud"] = y_spud["gross_wells_spud"]
    y_out.loc["Net Wells Spud"] = y_spud["net_wells_spud"]

    row_order = [
        "Assumed Index Pricing - Crude Oil",
        "Assumed Index Pricing - Natural Gas",
        "Realized Pricing - Crude Oil",
        "Realized Pricing - NGL (% of WTI)",
        "Realized Pricing - Natural Gas",
        "Gross Wells Spud",
        "Net Wells Spud",
        "Production - Crude Oil",
        "Production - NGL's",
        "Production - Natural Gas",
        "Production - Total (Mcfe/d)",
        "Revenues - Crude Oil",
        "Revenues - NGL's",
        "Revenues - Natural Gas",
        "Revenues - Total",
        "Operating Expenses - Taxes",
        "Operating Expenses - LOE",
        "Operating Expenses - Total Opex",
        "Taxes / Mcfe",
        "LOE / Mcfe",
        "EBITDA",
        "Capital Expenditures - D&C",
        "Capital Expenditures - Acquisition",
        "Capital Expenditures - Total",
        "Free Cash Flow",
        "Cumulative FCF",
    ]

    q_out = q_out.reindex(row_order)
    y_out = y_out.reindex(row_order)

    separator = pd.DataFrame(index=q_out.index, columns=[" "], data="")
    final = pd.concat([q_out, separator, y_out], axis=1)
    return final


def build_quarterly_output_display_table(df):
    first_col = "$ in Thousands"
    data_cols = list(df.columns)

    pct_rows = {"Realized Pricing - NGL (% of WTI)"}
    dollar_per_unit_rows = {"Taxes / Mcfe", "LOE / Mcfe"}
    price_rows = {
        "Assumed Index Pricing - Crude Oil",
        "Assumed Index Pricing - Natural Gas",
        "Realized Pricing - Crude Oil",
        "Realized Pricing - Natural Gas",
    }
    production_rows = {
        "Production - Crude Oil",
        "Production - NGL's",
        "Production - Natural Gas",
        "Production - Total (Mcfe/d)",
        "Gross Wells Spud",
        "Net Wells Spud",
    }

    def fmt_value(source_row, col):
        val = df.loc[source_row, col]

        if col == " ":
            return ""
        
        if pd.isnull(val) or val == "":
            if source_row in production_rows or source_row in pct_rows or source_row in dollar_per_unit_rows or source_row in price_rows:
                return "-"
            return "-"

        if source_row in pct_rows:
            return format_accounting_percent(val, decimals=0)
        if source_row in dollar_per_unit_rows:
            return format_accounting_number(val, decimals=2, prefix="$")
        if source_row in price_rows:
            return format_accounting_number(val, decimals=2, prefix="$")
        if source_row in production_rows:
            return format_accounting_production(val)
        return format_accounting_number(val, decimals=1, prefix="$")

    rows = []
    row_styles = []

    def add_section(label):
        row = {first_col: label}
        for c in data_cols:
            row[c] = ""
        rows.append(row)
        row_styles.append("section")

    def add_gap():
        row = {first_col: ""}
        for c in data_cols:
            row[c] = ""
        rows.append(row)
        row_styles.append("gap")

    def add_data(label, source_row, indent=False, style="normal"):
        display_label = f"\u00A0\u00A0\u00A0\u00A0{label}" if indent else label
        row = {first_col: display_label}
        for c in data_cols:
            row[c] = fmt_value(source_row, c)
        rows.append(row)
        row_styles.append(style)

    add_section("Assumed Index Pricing")
    add_data("Crude Oil", "Assumed Index Pricing - Crude Oil", indent=True)
    add_data("Natural Gas", "Assumed Index Pricing - Natural Gas", indent=True)

    add_gap()

    add_section("Realized Pricing")
    add_data("Crude Oil", "Realized Pricing - Crude Oil", indent=True)
    add_data("Natural Gas", "Realized Pricing - Natural Gas", indent=True)
    add_data("NGL (% of WTI)", "Realized Pricing - NGL (% of WTI)", indent=True)

    add_gap()

    add_data("Gross Wells Spud", "Gross Wells Spud")
    add_data("Net Wells Spud", "Net Wells Spud")

    add_gap()

    add_section("Production")
    add_data("Crude Oil", "Production - Crude Oil", indent=True)
    add_data("Natural Gas", "Production - Natural Gas", indent=True)
    add_data("NGL's", "Production - NGL's", indent=True)
    add_data("Total (Mcfe/d)", "Production - Total (Mcfe/d)", style="bold")

    add_gap()

    add_section("Revenues")
    add_data("Crude Oil", "Revenues - Crude Oil", indent=True)
    add_data("Natural Gas", "Revenues - Natural Gas", indent=True)
    add_data("NGL's", "Revenues - NGL's", indent=True)
    add_data("Total", "Revenues - Total")

    add_gap()

    add_section("Operating Expenses")
    add_data("Taxes", "Operating Expenses - Taxes", indent=True)
    add_data("LOE", "Operating Expenses - LOE", indent=True)
    add_data("Total", "Operating Expenses - Total Opex")

    add_gap()

    add_data("Taxes / Mcfe", "Taxes / Mcfe", style="italic")
    add_data("LOE / Mcfe", "LOE / Mcfe", style="italic")

    add_gap()

    add_data("EBITDA", "EBITDA", style="bold")

    add_gap()

    add_section("Capital Expenditures")
    add_data("D&C", "Capital Expenditures - D&C", indent=True)
    add_data("Acquisition", "Capital Expenditures - Acquisition", indent=True)
    add_data("Total", "Capital Expenditures - Total")

    add_gap()

    add_data("Free Cash Flow", "Free Cash Flow", style="bold")

    add_gap()

    add_data("Cumulative FCF", "Cumulative FCF", style="footer")

    display_df = pd.DataFrame(rows)
    return display_df, row_styles


def style_quarterly_output_table(display_df, row_styles):
    style_map = pd.Series(row_styles, index=display_df.index)

    first_col = display_df.columns[0]
    data_cols = list(display_df.columns[1:])

    quarter_cols = [c for c in data_cols if str(c).startswith("Q")]
    year_cols = [c for c in data_cols if str(c).isdigit()]
    separator_cols = [c for c in data_cols if str(c).strip() == ""]

    def row_style(row):
        rtype = style_map.loc[row.name]
        styles = [""] * len(row)

        if rtype == "section":
            styles = ["font-weight: 700; text-align: left;"] + [""] * (len(row) - 1)
        elif rtype == "bold":
            styles = ["font-weight: 700;"] * len(row)
        elif rtype == "italic":
            styles = ["font-style: italic;"] * len(row)
        elif rtype == "footer":
            styles = [f"background-color: {QUARTERLY_HEADER_COLOR}; color: white; font-weight: 700;"] * len(row)
        elif rtype == "gap":
            styles = [""] * len(row)

        return styles

    return (
        display_df.style
        .apply(row_style, axis=1)
        .hide(axis="index")
        .set_properties(subset=[first_col], **{
            "text-align": "left",
            "white-space": "pre",
            "background-color": "white",
        })
        .set_properties(subset=quarter_cols, **{
            "text-align": "right",
            "background-color": "white",
        })
        .set_properties(subset=year_cols, **{
            "text-align": "right",
            "background-color": "white",
        })
        .set_properties(subset=separator_cols, **{
            "background-color": QUARTERLY_HEADER_COLOR,
            "width": "4px",
            "min-width": "4px",
            "max-width": "4px",
            "padding": "0",
            "border": "none",
        })

        .set_table_styles([
            {
                "selector": "table",
                "props": [
                    ("border-collapse", "collapse"),
                    ("border-spacing", "0"),
                    ("width", "100%"),
                    ("table-layout", "fixed"),
                    ("border", "none !important"),
                ],
            },
            {
                "selector": "tbody td",
                "props": [
                    ("padding", "4px 8px"),
                    ("font-size", "12px"),
                    ("border", "none !important"),
                    ("outline", "none !important"),
                    ("box-shadow", "none !important"),
                ],
            },
            {
                "selector": "tbody td",
                "props": [
                    ("border", "none"),
                    ("padding", "4px 8px"),
                    ("font-size", "12px"),
                ],
            },
            {
                "selector": "tbody td.col0",
                "props": [
                    ("text-align", "left"),
                    ("white-space", "pre"),
                    ("width", "220px"),
                ],
            },
            {
                "selector": "tbody td:not(.col0)",
                "props": [
                    ("text-align", "right"),
                    ("width", "80px"),
                ],
            },
            {
                "selector": "tbody tr",
                "props": [
                    ("border", "none !important"),
                ],
            },
            {
                "selector": "tbody td",
                "props": [
                    ("background-clip", "padding-box"),
                ],
            },
            {
                "selector": "thead th",
                "props": [
                    ("background-color", QUARTERLY_HEADER_COLOR),
                    ("color", "white"),
                    ("font-weight", "700"),
                    ("font-size", "12px"),
                    ("padding", "4px 8px"),
                    ("border", "none !important"),
                    ("outline", "none !important"),
                    ("box-shadow", "none !important"),
                    ("text-align", "right"),
                ],
            },
            {
                "selector": "thead th.col0",
                "props": [
                    ("text-align", "left"),
                ],
            },
            {
                "selector": "tbody td",
                "props": [
                    ("border-top", "none !important"),
                    ("border-bottom", "none !important"),
                ],
            },
            {
                "selector": f"tbody tr:nth-child({len(display_df)}) td",
                "props": [
                    ("background-color", QUARTERLY_HEADER_COLOR),
                    ("color", "white"),
                    ("font-weight", "700"),
                ],
            },
        ], overwrite=False)
    )

def calc_slot_eur_metrics(slot_row, deal_inputs):
    type_curve_library = load_type_curve_library("type_curve_library.xlsx")
    deal_settings = prepare_deal_settings(deal_inputs)
    global_assumptions = prepare_global_assumptions(deal_inputs)

    slot_inputs_df = prepare_slot_inputs(pd.DataFrame([slot_row]), deal_inputs)
    slot_prepped = slot_inputs_df.iloc[0].copy()

    total_net_acres = float(slot_prepped.get("net_acres", 0.0))

    slot_prepped = calc_slot_metrics(
        slot_prepped,
        deal_settings=deal_settings,
        total_net_acres=total_net_acres,
    )

    slot_ngl = build_slot_ngl_factors(
        slot=slot_prepped,
        global_assumptions=global_assumptions,
        content_percentages=global_assumptions["content_percentages"],
        recover_ethane_percentages=global_assumptions["recover_ethane_percentages"],
        reject_ethane_percentages=global_assumptions["reject_ethane_percentages"],
        ngl_prices=global_assumptions["ngl_prices"],
        ngl_shrink_factors=global_assumptions["ngl_shrink_factors"],
    )

    one_well_df = run_single_slot_economics(
        slot=slot_prepped,
        type_curve_library=type_curve_library,
        global_assumptions=global_assumptions,
        slot_ngl=slot_ngl,
    ).copy()

    prod_df = one_well_df[one_well_df["period"].between(1, 360)].copy()

    lateral_length = float(slot_prepped["lateral_length"])
    if lateral_length == 0:
        return None, None, float(slot_ngl["shrink"])

    oil_eur_per_ft = prod_df["gross_oil_production"].sum() / lateral_length
    gas_eur_per_ft = prod_df["gross_gas_production"].sum() / lateral_length
    gas_shrink = float(slot_ngl["shrink"])

    return oil_eur_per_ft, gas_eur_per_ft, gas_shrink

def build_tc_assumptions_output_display_table(slot_df, deal_inputs, slot_returns=None):
    df = slot_df.copy()

    if slot_returns is None:
        slot_returns = {}

    if df.empty:
        return pd.DataFrame({"TC Assumptions": []}), []

    df["drilling_spud_month"] = pd.to_datetime(df["drilling_spud_month"], errors="coerce")
    display_cols = [f"Slot {int(s)}" for s in df["slot_id"]]

    rows = []
    row_styles = []

    def add_section(label):
        row = {"TC Assumptions": label}
        for c in display_cols:
            row[c] = ""
        rows.append(row)
        row_styles.append("section")

    def add_gap():
        row = {"TC Assumptions": ""}
        for c in display_cols:
            row[c] = ""
        rows.append(row)
        row_styles.append("gap")

    def add_data(label, values, style="normal"):
        row = {"TC Assumptions": label}
        row.update(values)
        rows.append(row)
        row_styles.append(style)

    slot_map = {}
    for _, r in df.iterrows():
        slot_name = f"Slot {int(r['slot_id'])}"
        slot_map[slot_name] = r

    def fmt_num(x, decimals=1, prefix="", suffix=""):
        return format_accounting_number(
            x,
            decimals=decimals,
            prefix=prefix,
            suffix=suffix,
            null_as_blank=False,
        )

    def fmt_pct(x, decimals=0):
        return format_accounting_percent(x, decimals=decimals, null_as_blank=False)

    def fmt_date(x):
        if pd.isnull(x):
            return "-"
        return pd.to_datetime(x).strftime("%m/%d/%y")

    def fmt_tc_name(x):
        if pd.isnull(x):
            return "-"

        text = str(x)
        parts = text.split("_")

        # Example:
        # lean_cond_plus_wet_gas
        # becomes:
        # lean_cond_plus<br>wet_gas
        if len(parts) >= 3:
            first_line = "_".join(parts[:-2])
            second_line = "_".join(parts[-2:])
            return f"{first_line}<br>{second_line}"

        return text

    def get_slot_return(slot_id, metric):
        slot_id = int(slot_id)
        return slot_returns.get(slot_id, {}).get(metric)
    
    gas_shrink_pct = {k: None for k in slot_map.keys()}
    oil_eur_per_ft = {k: None for k in slot_map.keys()}
    gas_eur_per_ft = {k: None for k in slot_map.keys()}

    for slot_name, v in slot_map.items():
        slot_one_row_df = pd.DataFrame([v]).copy()

        slot_inputs_prepped = prepare_slot_inputs(slot_one_row_df, deal_inputs)
        slot_prepped = slot_inputs_prepped.iloc[0].copy()

        deal_settings = prepare_deal_settings(deal_inputs)
        global_assumptions = prepare_global_assumptions(deal_inputs)
        type_curve_library = load_type_curve_library("type_curve_library.xlsx")

        total_net_acres = float(slot_prepped.get("net_acres", 0.0))

        slot_prepped = calc_slot_metrics(
            slot_prepped,
            deal_settings=deal_settings,
            total_net_acres=total_net_acres,
        )

        slot_ngl = build_slot_ngl_factors(
            slot=slot_prepped,
            global_assumptions=global_assumptions,
            content_percentages=global_assumptions["content_percentages"],
            recover_ethane_percentages=global_assumptions["recover_ethane_percentages"],
            reject_ethane_percentages=global_assumptions["reject_ethane_percentages"],
            ngl_prices=global_assumptions["ngl_prices"],
            ngl_shrink_factors=global_assumptions["ngl_shrink_factors"],
        )

        one_well_df = run_single_slot_economics(
            slot=slot_prepped,
            type_curve_library=type_curve_library,
            global_assumptions=global_assumptions,
            slot_ngl=slot_ngl,
        ).copy()

        prod_df = one_well_df[one_well_df["period"].between(1, 360)].copy()

        lateral_length = (
            float(slot_prepped["lateral_length"])
            if pd.notnull(slot_prepped["lateral_length"]) and float(slot_prepped["lateral_length"]) != 0
            else None
        )

        gas_shrink_pct[slot_name] = float(slot_ngl["shrink"])

        if lateral_length:
            oil_eur_per_ft[slot_name] = (
                prod_df["gross_oil_production"].sum() / lateral_length
            )
            
            gas_eur_per_ft[slot_name] = (
                prod_df["gross_gas_production"].sum() / lateral_length
            )
            
    add_section("Development")
    add_data("Type Curve", {k: fmt_tc_name(v["tc_name"]) for k, v in slot_map.items()})
    add_data("Gross Wells", {k: fmt_num(v["gross_wells"], decimals=2) for k, v in slot_map.items()})
    add_data("Net Acres", {k: fmt_num(v["net_acres"], decimals=1) for k, v in slot_map.items()})
    add_data("Unit Acres", {k: fmt_num(v["unit_acres"], decimals=0) for k, v in slot_map.items()})
    add_data("% Unitized", {k: fmt_pct(v["pct_unitized"], decimals=0) for k, v in slot_map.items()})
    add_data("Spud Month", {k: fmt_date(v["drilling_spud_month"]) for k, v in slot_map.items()})
    add_data("Flowback Delay", {k: fmt_num(v["flowback_delay"], decimals=0) for k, v in slot_map.items()})
    add_data("Lateral Length (ft)", {k: fmt_num(v["lateral_length"], decimals=0) for k, v in slot_map.items()})
    add_data(
        "WI Reversion Eligible",
        {k: ("Yes" if bool(v.get("dale_promote", False)) else "No") for k, v in slot_map.items()},
    )

    add_gap()

    add_section("Production")
    add_data("TC Risk", {k: fmt_pct(v["tc_risk"], decimals=0) for k, v in slot_map.items()})
    add_data("NRI", {k: fmt_pct(v["net_revenue_interest"], decimals=0) for k, v in slot_map.items()})
    add_data("Gas Shrink %", {k: fmt_pct(gas_shrink_pct[k], decimals=1) for k in slot_map.keys()})
    add_data("Oil EUR (Bbl/ft)", {k: fmt_num(oil_eur_per_ft[k], decimals=0) for k in slot_map.keys()})
    add_data("Gas EUR (Mcf/ft)", {k: fmt_num(gas_eur_per_ft[k], decimals=0) for k in slot_map.keys()})
    add_data("NGL Yield", {k: fmt_num(v["ngl_yield"], decimals=2) for k, v in slot_map.items()})
    
    add_gap()
    
    add_section("Returns")
    add_data(
        "IRR",
        {
            k: fmt_pct(get_slot_return(v["slot_id"], "irr"), decimals=1)
            for k, v in slot_map.items()
        },
    )
    add_data(
        "MOIC",
        {
            k: fmt_num(get_slot_return(v["slot_id"], "moic"), decimals=2, suffix="x")
            for k, v in slot_map.items()
        },
    )
    
    add_gap()
    
    add_section("Economics")
    add_data("D&C ($/ft)", {k: fmt_num(v["dc_costs"], decimals=0, prefix="$") for k, v in slot_map.items()})
    add_data("$/Acre Bid", {k: fmt_num(v["bid_per_acre"], decimals=0, prefix="$") for k, v in slot_map.items()})
    add_data("Oil Diff", {k: fmt_num(v["oil_diff"], decimals=2, prefix="$") for k, v in slot_map.items()})
    add_data("Gas Diff", {k: fmt_num(v["gas_diff"], decimals=2, prefix="$") for k, v in slot_map.items()})

    add_gap()

    add_section("Operating Costs")
    add_data("Oil Opex", {k: fmt_num(v["oil_opex_bbl"], decimals=2, prefix="$") for k, v in slot_map.items()})
    add_data("Gas Opex", {k: fmt_num(v["gas_opex_mcf"], decimals=2, prefix="$") for k, v in slot_map.items()})
    add_data("NGL Opex", {k: fmt_num(v["ngl_opex"], decimals=2, prefix="$") for k, v in slot_map.items()})
    add_data("Fixed LOE", {k: fmt_num(v["fixed_loe"], decimals=0, prefix="$") for k, v in slot_map.items()}, style="footer")

    display_df = pd.DataFrame(rows)
    return display_df, row_styles

def style_tc_assumptions_output_table(display_df, row_styles):
    style_map = pd.Series(row_styles, index=display_df.index)

    first_col = display_df.columns[0]
    other_cols = list(display_df.columns[1:])

    def row_style(row):
        rtype = style_map.loc[row.name]
        styles = [""] * len(row)

        if rtype == "section":
            styles = ["font-weight: 700; text-align: left;"] + [""] * (len(row) - 1)
        elif rtype == "footer":
            styles = [f"background-color: {QUARTERLY_HEADER_COLOR}; color: white; font-weight: 700;"] * len(row)
        elif rtype == "gap":
            styles = [""] * len(row)

        return styles

    return (
    display_df.style
    .apply(row_style, axis=1)
    .hide(axis="index")
    .set_properties(subset=[first_col], **{
        "text-align": "left",
        "white-space": "pre-wrap",
        "background-color": "white",
        "width": "180px",
        "min-width": "180px",
        "max-width": "180px",
    })
    .set_properties(subset=other_cols, **{
        "text-align": "right",
        "background-color": "white",
        "white-space": "normal",
        "overflow-wrap": "anywhere",
        "word-break": "break-word",
        "vertical-align": "top",
    })
    .set_table_styles([
        {
            "selector": "table",
            "props": [
                ("border-collapse", "separate"),
                ("border-spacing", "0"),
                ("width", "100%"),
                ("table-layout", "fixed"),
            ],
        },
        {
            "selector": "thead th",
            "props": [
                ("background-color", QUARTERLY_HEADER_COLOR),
                ("color", "white"),
                ("font-weight", "700"),
                ("text-align", "center"),
                ("border", "none"),
                ("padding", "4px 8px"),
                ("font-size", "12px"),
                ("white-space", "normal"),
                ("overflow-wrap", "anywhere"),
                ("word-break", "break-word"),
            ],
        },
        {
            "selector": "thead th.col0",
            "props": [
                ("text-align", "left"),
                ("width", "180px"),
                ("min-width", "180px"),
                ("max-width", "180px"),
            ],
        },
        {
            "selector": "thead th:not(.col0)",
            "props": [
                ("width", "120px"),
                ("min-width", "120px"),
                ("max-width", "120px"),
            ],
        },
        {
            "selector": "tbody td",
            "props": [
                ("border", "none"),
                ("padding", "4px 8px"),
                ("font-size", "12px"),
                ("vertical-align", "top"),
            ],
        },
        {
            "selector": "tbody td.col0",
            "props": [
                ("text-align", "left"),
                ("white-space", "pre-wrap"),
                ("width", "180px"),
                ("min-width", "180px"),
                ("max-width", "180px"),
            ],
        },
        {
            "selector": "tbody td:not(.col0)",
            "props": [
                ("text-align", "right"),
                ("white-space", "normal"),
                ("overflow-wrap", "anywhere"),
                ("word-break", "break-word"),
                ("width", "120px"),
                ("min-width", "120px"),
                ("max-width", "120px"),
            ],
        },
    ], overwrite=False)
)

def render_deal_highlight_box(title, value):
    st.markdown(
        f"""
        <div style="
            background-color: {QUARTERLY_HEADER_COLOR};
            color: white;
            padding: 14px 10px;
            border-radius: 6px;
            text-align: center;
            font-weight: 700;
        ">
            <div style="font-size: 13px; margin-bottom: 6px;">{title}</div>
            <div style="font-size: 24px;">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_production_profile_chart(deal_df, chart_view="Stacked Mcfe/d"):
    df = deal_df.copy()
    df["date"] = pd.to_datetime(df["date"])
    df = df[df["date"] <= pd.Timestamp("2040-12-31")].copy()

    # Convert monthly net volumes to Mcfe/d
    df["days_in_month"] = df["date"].dt.days_in_month
    df["net_oil_mcfe_d"] = (df["slot_net_oil_production"] * 6.0) / df["days_in_month"]
    df["net_ngl_mcfe_d"] = (df["slot_net_ngl_production"] * 6.0) / df["days_in_month"]
    df["net_gas_mcfe_d"] = df["slot_net_gas_production"] / df["days_in_month"]

    # Optional total line if you ever want it later
    df["total_mcfe_d"] = (
        df["net_oil_mcfe_d"] + df["net_ngl_mcfe_d"] + df["net_gas_mcfe_d"]
    )

    fig = go.Figure()

    oil_color = "#1f4e79"
    ngl_color = "#4e80b1"
    gas_color = "#b7cde3"

    if chart_view == "Stacked Mcfe/d":
        fig.add_trace(
            go.Scatter(
                x=df["date"],
                y=df["net_oil_mcfe_d"],
                mode="lines",
                name="Oil",
                stackgroup="one",
                line=dict(color=oil_color, width=1.5),
                hovertemplate="Date: %{x|%Y-%m-%d}<br>Net Oil: %{y:,.1f} Mcfe/d<extra></extra>",
            )
        )

        fig.add_trace(
            go.Scatter(
                x=df["date"],
                y=df["net_ngl_mcfe_d"],
                mode="lines",
                name="NGL",
                stackgroup="one",
                line=dict(color=ngl_color, width=1.5),
                hovertemplate="Date: %{x|%Y-%m-%d}<br>Net NGL: %{y:,.1f} Mcfe/d<extra></extra>",
            )
        )

        fig.add_trace(
            go.Scatter(
                x=df["date"],
                y=df["net_gas_mcfe_d"],
                mode="lines",
                name="Gas",
                stackgroup="one",
                line=dict(color=gas_color, width=1.5),
                hovertemplate="Date: %{x|%Y-%m-%d}<br>Net Gas: %{y:,.1f} Mcfe/d<extra></extra>",
            )
        )

        chart_title = "Net Production Profile (Mcfe/d)"
    else:
        fig.add_trace(
            go.Scatter(
                x=df["date"],
                y=df["net_oil_mcfe_d"],
                mode="lines",
                name="Oil",
                line=dict(color=oil_color, width=3),
                hovertemplate="Date: %{x|%Y-%m-%d}<br>Net Oil: %{y:,.1f} Mcfe/d<extra></extra>",
            )
        )

        fig.add_trace(
            go.Scatter(
                x=df["date"],
                y=df["net_ngl_mcfe_d"],
                mode="lines",
                name="NGL",
                line=dict(color=ngl_color, width=3),
                hovertemplate="Date: %{x|%Y-%m-%d}<br>Net NGL: %{y:,.1f} Mcfe/d<extra></extra>",
            )
        )

        fig.add_trace(
            go.Scatter(
                x=df["date"],
                y=df["net_gas_mcfe_d"],
                mode="lines",
                name="Gas",
                line=dict(color=gas_color, width=3),
                hovertemplate="Date: %{x|%Y-%m-%d}<br>Net Gas: %{y:,.1f} Mcfe/d<extra></extra>",
            )
        )

        chart_title = "Net Production Stream Split (Mcfe/d)"

    fig.update_layout(
        title=dict(
            text=chart_title,
            font=dict(color="black"),
        ),
        xaxis=dict(
            title=dict(text="Date", font=dict(color="black")),
            tickformat="%Y",
            dtick="M12",
            tickfont=dict(color="black"),
        ),
        yaxis=dict(
            title=dict(text="Net Production (Mcfe/d)", font=dict(color="black")),
            tickfont=dict(color="black"),
        ),
        
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.18,
            xanchor="center",
            x=0.5,
            font=dict(size=17, color="black"),
            traceorder="normal",
            entrywidth=200,
            entrywidthmode="pixels",
            tracegroupgap=28,
        ),
    )
    
    fig.update_xaxes(
        tickfont=dict(size=13, color="black"),
        title_font=dict(size=15, color="black"),
    )
        
    fig.update_yaxes(
        tickfont=dict(size=13, color="black"),
        title_font=dict(size=15, color="black"),
    )

    return fig

def build_cumulative_fcf_chart(deal_df, slot_df):
    df = deal_df.copy()
    df["date"] = pd.to_datetime(df["date"])
    df = df[df["date"] <= pd.Timestamp("2040-12-31")].copy()

    monthly_fcf = df.groupby("date", as_index=False)["slot_total_cash_flow"].sum()
    monthly_fcf["cum_fcf"] = monthly_fcf["slot_total_cash_flow"].cumsum() / 1000.0

    payback_years = None
    payback_date = None

    for i in range(1, len(monthly_fcf)):
        prev_val = monthly_fcf.loc[i - 1, "cum_fcf"]
        curr_val = monthly_fcf.loc[i, "cum_fcf"]

        if prev_val < 0 <= curr_val:
            prev_date = monthly_fcf.loc[i - 1, "date"]
            curr_date = monthly_fcf.loc[i, "date"]

            frac = 0 if curr_val == prev_val else (0 - prev_val) / (curr_val - prev_val)
            payback_date = prev_date + (curr_date - prev_date) * frac
            start_date = monthly_fcf.loc[0, "date"]
            payback_years = (payback_date - start_date).days / 365.25
            break

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=monthly_fcf["date"],
            y=monthly_fcf["cum_fcf"],
            mode="lines",
            name="Cumulative FCF",
            line=dict(color="#4E80B1", width=3),
            fill="tozeroy",
            fillcolor="rgba(78, 128, 177, 0.28)",
            hovertemplate="Date: %{x|%Y-%m-%d}<br>Cumulative FCF: %{y:,.1f}<extra></extra>",
        )
    )

    fig.add_hline(y=0, line_width=1, line_dash="dash", line_color="gray")

    slot_chart = slot_df.copy()
    slot_chart["drilling_spud_month"] = pd.to_datetime(
        slot_chart["drilling_spud_month"], errors="coerce"
    )
    slot_chart = slot_chart[
        slot_chart["drilling_spud_month"] <= pd.Timestamp("2040-12-31")
    ].copy()

    if not slot_chart.empty:
        spud_summary = (
            slot_chart.groupby("drilling_spud_month", as_index=False)["gross_wells"]
            .sum()
            .sort_values("drilling_spud_month")
        )
    
        for _, row in spud_summary.iterrows():
            spud_date = row["drilling_spud_month"]
            gross_wells = row["gross_wells"]
    
            x0 = spud_date
            x1 = spud_date + pd.offsets.MonthEnd(1)
    
            # 🔹 Soft shaded band (no outline)
            fig.add_vrect(
                x0=x0,
                x1=x1,
                fillcolor="rgba(78, 128, 177, 0.18)",  # softer fill
                line_width=0,  # removes harsh border
                layer="below",
            )
    
            # 🔹 Move Gross Wells LOWER so it never conflicts with payback
            fig.add_annotation(
                x=spud_date + pd.Timedelta(days=14),
                y=0.88,
                yref="paper",
                text=f"{gross_wells:.1f} Gross Wells",
                showarrow=False,
                textangle=-90,
                font=dict(size=10, color="black"),
                bgcolor="rgba(255,255,255,0)",
                bordercolor="rgba(0,0,0,0)",
                borderwidth=0,
                xanchor="center",
                yanchor="middle",
            )
            
    if payback_date is not None and payback_years is not None:
        fig.add_vline(
            x=payback_date,
            line_width=1,
            line_dash="dot",
            line_color="gray",
        )

        fig.add_annotation(
            x=payback_date,
            y=1.08,
            yref="paper",
            text=f"<b>Payback = {payback_years:.1f} years</b>",
            showarrow=True,
            arrowhead=2,
            ax=0,
            ay=28,
            font=dict(size=16, color="black"),  # slightly bigger
            bgcolor="rgba(255,255,255,0)",     # no box
            bordercolor="rgba(0,0,0,0)",
            borderwidth=0,
        )

    fig.update_layout(
        title=dict(
            text="<b>Cumulative Free Cash Flow</b>",
            x=0.5,
            xanchor="center",
            font=dict(size=20, color="black"),
        ),
        xaxis=dict(
            title=dict(text="Date", font=dict(size=14, color="black")),
            tickformat="%Y",
            dtick="M12",
            tickfont=dict(size=12, color="black"),
        ),
        yaxis=dict(
            title=dict(text="$ in Thousands", font=dict(size=14, color="black")),
            tickfont=dict(size=12, color="black"),
        ),
        height=525,
        margin=dict(l=50, r=40, t=95, b=45),
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=False,
    )

    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.08)")

    return fig

@st.cache_data(show_spinner=False)
def build_scenario_scatter_chart(slot_df, deal_inputs, base_bid, base_dc):
    bid_values = build_sensitivity_range(
        base_bid,
        500.0,
        3,
        min_value=1.0,
    )

    dc_cases = [
        ("Low", base_dc - 50.0),
        ("Base", base_dc),
        ("High", base_dc + 50.0),
    ]
    
    tc_risk_values = [0.80, 1.00, 1.20]

    base_oil = float(deal_inputs["oil_price"])
    base_gas = float(deal_inputs["gas_price"])

    pricing_cases = [
        ("Downside", max(0.0, base_oil - 5.0), max(0.0, base_gas - 0.25)),
        ("Base", base_oil, base_gas),
        ("Upside", base_oil + 5.0, base_gas + 0.25),
    ]

    rows = []

    for pricing_name, oil_price, gas_price in pricing_cases:
        for dc_label, dc_value in dc_cases:
            for tc_risk in tc_risk_values:
                for bid in bid_values:
                    sens_inputs = deal_inputs.copy()
                    sens_inputs["oil_price"] = float(oil_price)
                    sens_inputs["gas_price"] = float(gas_price)
                    sens_inputs["use_bid_override"] = True
                    sens_inputs["bid_override"] = float(bid)
                    sens_inputs["use_dc_override"] = True
                    sens_inputs["dc_override"] = float(dc_value)

                    sens_slot_df = slot_df.copy()
                    sens_slot_df["tc_risk"] = float(tc_risk)

                    try:
                        _, _, _, _, irr, moic = run_deal_model(sens_slot_df, sens_inputs)
                    except Exception:
                        irr, moic = None, None

                    rows.append(
                        {
                            "pricing_case": pricing_name,
                            "oil_price": oil_price,
                            "gas_price": gas_price,
                            "dc_case": dc_label,
                            "dc_value": dc_value,
                            "tc_risk": tc_risk,
                            "bid": bid,
                            "irr": irr,
                            "moic": moic,
                        }
                    )

    chart_df = pd.DataFrame(rows)
    chart_df = chart_df[pd.notnull(chart_df["irr"])].copy()

    color_map = {
        "Low": "#9ECAE1",
        "Base": "#4E80B1",
        "High": "#1F4E79",
    }

    size_map = {
        0.80: 10,
        1.00: 20,
        1.20: 32,
    }

    dc_label_map = {
        "Low": f"Low (${base_dc - 100.0:,.0f}/ft)",
        "Base": f"Base (${base_dc:,.0f}/ft)",
        "High": f"High (${base_dc + 100.0:,.0f}/ft)",
    }
    
    fig = make_subplots(
        rows=1,
        cols=3,
        shared_yaxes=True,
        horizontal_spacing=0.06,
    )
    
    panel_col_map = {"Downside": 1, "Base": 2, "Upside": 3}
    legend_seen = set()
    tc_jitter = {
        0.90: 0,
        1.00: 0,
        1.10: 0,
    }
    
    for pricing_name in ["Downside", "Base", "Upside"]:
        panel_df = chart_df[chart_df["pricing_case"] == pricing_name].copy()
        col_num = panel_col_map[pricing_name]
    
        for dc_case in ["Low", "Base", "High"]:
            dc_df = panel_df[panel_df["dc_case"] == dc_case].copy()
            if dc_df.empty:
                continue
    
            marker_sizes = [size_map.get(float(x), 14) for x in dc_df["tc_risk"]]
            show_legend = dc_case not in legend_seen
    
            fig.add_trace(
                go.Scatter(
                    x=[
                        b + tc_jitter.get(round(float(r), 2), 0)
                        for b, r in zip(dc_df["bid"], dc_df["tc_risk"])
                    ],                    
                    y=dc_df["irr"],
                    mode="markers",
                    name=dc_label_map[dc_case],
                    legendgroup="dc",
                    showlegend=show_legend,
                    marker=dict(
                        color=color_map[dc_case],
                        size=marker_sizes,
                        line=dict(color="white", width=0.5),
                        opacity=0.70,
                    ),
                    hovertemplate=(
                        "Bid: $%{x:,.0f}"
                        "<br>IRR: %{y:.1%}"
                        "<br>D&C: " + dc_label_map[dc_case] +
                        "<br>TC Risk: %{customdata[0]:.0%}"
                        "<br>Oil: $%{customdata[1]:.0f}"
                        "<br>Gas: $%{customdata[2]:.2f}"
                        "<extra></extra>"
                    ),
                    customdata=dc_df[["tc_risk", "oil_price", "gas_price"]].values,
                ),
                row=1,
                col=col_num,
            )

            legend_seen.add(dc_case)

    base_tc_risk = round(weighted_avg_by_net_acres(slot_df, "tc_risk"), 2)
    base_bid_rounded = round(float(base_bid), 2)
    
    base_points = chart_df[
        (chart_df["pricing_case"] == "Base")
        & (chart_df["dc_case"] == "Base")
        & (chart_df["tc_risk"].round(2) == base_tc_risk)
        & (chart_df["bid"].round(2) == base_bid_rounded)
    ].copy()
    
    if base_points.empty:
        base_points = chart_df[
            (chart_df["pricing_case"] == "Base")
            & (chart_df["dc_case"] == "Base")
            & (chart_df["tc_risk"].round(2) == 1.00)
            & (chart_df["bid"].round(2) == base_bid_rounded)
        ].copy()
    

    if not base_points.empty:
        fig.add_trace(
            go.Scatter(
                x=base_points["bid"],
                y=base_points["irr"],
                mode="markers",
                name="Current Base Point",
                marker=dict(
                    color="#1F4E79",
                    size=24,
                    line=dict(color="black", width=3),
                    opacity=1.0,
                ),
                hovertemplate="Current Base Point<br>Bid: $%{x:,.0f}<br>IRR: %{y:.1%}<extra></extra>",
            ),
            row=1,
            col=2,
    )

    for c in [1, 2, 3]:
        fig.update_xaxes(
            title_text="$/Acre Bid",
            tickprefix="$",
            tickformat=",.0f",
            showgrid=False,
            row=1,
            col=c,
        )

    fig.update_yaxes(
        title_text="IRR",
        tickformat=".0%",
        showgrid=True,
        gridcolor="rgba(0,0,0,0.08)",
        row=1,
        col=1,
    )

    fig.update_layout(
        title=dict(
            text="Scenario Matrix: IRR vs. $/Acre Bid<br><sup>Color = D&C | Marker Size = TC Risk</sup>",
            x=0.5,
            xanchor="center",
            font=dict(size=26, color="black"),
        ),
        height=1000,
        margin=dict(l=80, r=60, t=195, b=245),
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.20,
            xanchor="center",
            x=0.5,
            font=dict(size=18, color="black"),   # 🔥 bigger legend
            traceorder="normal",
            entrywidth=200,                      # 🔥 slightly wider spacing
            entrywidthmode="pixels",
            tracegroupgap=28,
        ),
    )
    
    fig.update_xaxes(
        tickfont=dict(size=18, color="black"),   # 🔥 bigger ticks
        title_font=dict(size=20, color="black"), # 🔥 bigger label
    )
    
    fig.update_yaxes(
        tickfont=dict(size=16, color="black"),
        title_font=dict(size=18, color="black"),
    )

    fig.add_annotation(
        x=0.165, y=1.04,
        xref="paper", yref="paper",
        text=f"<b>Downside (Oil {pricing_cases[0][1]:.0f} / Gas {pricing_cases[0][2]:.2f})</b>",
        showarrow=False,
        font=dict(size=24, color="black"),
        xanchor="center",
        yanchor="bottom",
    )
    
    fig.add_annotation(
        x=0.50, y=1.04,
        xref="paper", yref="paper",
        text=f"<b>Base (Oil {pricing_cases[1][1]:.0f} / Gas {pricing_cases[1][2]:.2f})</b>",
        showarrow=False,
        font=dict(size=24, color="black"),
        xanchor="center",
        yanchor="bottom",
    )
    
    fig.add_annotation(
        x=0.835, y=1.04,
        xref="paper", yref="paper",
        text=f"<b>Upside (Oil {pricing_cases[2][1]:.0f} / Gas {pricing_cases[2][2]:.2f})</b>",
        showarrow=False,
        font=dict(size=24, color="black"),
        xanchor="center",
        yanchor="bottom",
    )

    # Centered custom TC Risk legend below the D&C legend.
    fig.add_annotation(
        x=0.50,
        y=-0.285,
        xref="paper",
        yref="paper",
        text=(
            "<span style='font-size:14px; color:rgba(120,120,120,0.85);'>●</span>"
            "&nbsp; TC Risk 80%"
            "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            "<span style='font-size:20px; color:rgba(120,120,120,0.85);'>●</span>"
            "&nbsp; TC Risk 100%"
            "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            "<span style='font-size:30px; color:rgba(120,120,120,0.85);'>●</span>"
            "&nbsp; TC Risk 120%"
        ),
        showarrow=False,
        font=dict(size=18, color="black"),
        xanchor="center",
        yanchor="middle",
        align="center",
    )
    
    return fig

def fig_to_base64_png(fig, width=1200, height=700, chart_name="Chart"):
    try:
        img_bytes = fig.to_image(
            format="png",
            width=width,
            height=height,
            scale=1,
        )
        return base64.b64encode(img_bytes).decode("utf-8")

    except Exception as exc:
        st.session_state["email_chart_export_error"] = (
            f"{chart_name} export error: "
            f"{type(exc).__name__}: {exc}"
        )
        return None


def html_img_from_fig(fig, width=900, height=500, title="Chart", max_width_px=None):
    img_b64 = fig_to_base64_png(
        fig,
        width=width,
        height=height,
        chart_name=title,
    )

    if img_b64 is None:
        return f"""
        <div style="margin:12px 0 20px 0; padding:12px; border:1px solid #cccccc;">
            <b>{title}</b><br>
            {st.session_state.get(
                "email_chart_export_error",
                "Image export unavailable in current environment.",
            )}
        </div>
        """

    max_width_style = f"max-width:{max_width_px}px;" if max_width_px else f"max-width:{width}px;"

    return f'''
        <img src="data:image/png;base64,{img_b64}"
             style="width:100%; {max_width_style} height:auto; margin:12px 0 20px 0; display:block;">
    '''

def build_email_html(
    opportunity_name,
    deal_inputs,
    slot_df,
    irr,
    moic,
    tc_output_styler,
    quarterly_output_styler,
    irr_oil_bid_heatmap,
    irr_gas_bid_heatmap,
    irr_heatmap,
    irr_tcrisk_bid_heatmap,
    irr_ngl_yield_bid_heatmap,
    irr_spud_tcrisk_heatmap,
    cum_fcf_chart,
    prod_chart_stacked,
    scenario_scatter_chart,
):
    base_gas = float(deal_inputs["gas_price"])
    base_oil = float(deal_inputs["oil_price"])

    base_dc = (
        float(deal_inputs["dc_override"])
        if deal_inputs["use_dc_override"]
        else weighted_avg_by_net_acres(slot_df, "dc_costs")
    )
    
    base_bid = max(
        1.0,
        (
            float(deal_inputs["bid_override"])
            if deal_inputs["use_bid_override"]
            else weighted_avg_by_net_acres(slot_df, "bid_per_acre")
        ),
    )
    
    total_wells = float(slot_df["gross_wells"].sum())
    avg_ll = weighted_avg_by_net_acres(slot_df, "lateral_length")
    avg_pct_unitized = weighted_avg_by_net_acres(slot_df, "pct_unitized")
    avg_tc_risk = weighted_avg_by_net_acres(slot_df, "tc_risk")

    tc_names = slot_df["tc_name"].dropna().astype(str).unique().tolist()
    tc_name_text = ", ".join(tc_names)

    tc_table_html = tc_output_styler.to_html()
    quarterly_table_html = quarterly_output_styler.to_html()

    sensitivities_html = f"""
    <h3 style='margin-bottom:8px;'>Sensitivities:</h3>
    
    <table role="presentation" style="width:auto; border-collapse:collapse; margin:0 0 20px 0;">
        <tr>
            <td style="vertical-align:top; padding:0 8px 12px 8px;">
                {html_img_from_fig(
                    irr_oil_bid_heatmap,
                    width=1100,
                    height=450,
                    title="Oil Price IRR",
                    max_width_px=760,
                )}
            </td>
            <td style="vertical-align:top; padding:0 8px 12px 8px;">
                {html_img_from_fig(
                    irr_gas_bid_heatmap,
                    width=1100,
                    height=450,
                    title="Gas Price IRR",
                    max_width_px=760,
                )}
            </td>
        </tr>
        <tr>
            <td style="vertical-align:top; padding:0 8px 12px 8px;">
                {html_img_from_fig(
                    irr_heatmap,
                    width=1100,
                    height=450,
                    title="D&C Costs IRR",
                    max_width_px=760,
                )}
            </td>
            <td style="vertical-align:top; padding:0 8px 12px 8px;">
                {html_img_from_fig(
                    irr_tcrisk_bid_heatmap,
                    width=1100,
                    height=450,
                    title="TC Risk IRR",
                    max_width_px=760,
                )}
            </td>
        </tr>
        <tr>
            <td style="vertical-align:top; padding:0 8px 0 8px;">
                {html_img_from_fig(
                    irr_ngl_yield_bid_heatmap,
                    width=1100,
                    height=450,
                    title="NGL Yield IRR",
                    max_width_px=760,
                )}
            </td>
            <td style="vertical-align:top; padding:0 8px 0 8px;">
                {html_img_from_fig(
                    irr_spud_tcrisk_heatmap,
                    width=1100,
                    height=450,
                    title="Spud Date vs. TC Risk IRR",
                    max_width_px=760,
                )}
            </td>
        </tr>
    </table>
    """
    
    charts_html = "".join([
        "<h3 style='margin-bottom:8px;'>Charts:</h3>",
        html_img_from_fig(
            cum_fcf_chart,
            width=950,
            height=450,
            title="Cumulative FCF",
            max_width_px=950,
        ),
        html_img_from_fig(prod_chart_stacked, width=1100, height=520, title="Production in BOE/d"),
        html_img_from_fig(
            scenario_scatter_chart,
            width=2600,
            height=1100,
            max_width_px=1800,
        )
    ])

    html = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #000000; line-height: 1.45;">
        <p>Utica Team,</p>
    
        <p>Below are our contemplated economics for the {opportunity_name} opportunity.</p>
    
        <p><b>Base Case Summary:</b></p>
        <ul style="margin-top:0;">
            <li>
                Pricing: ${base_gas:.2f}/Mcf gas
                {" (dry gas, no oil)" if base_oil == 0 else f" and ${base_oil:.0f}/Bbl oil"}
            </li>
            <li>
                D&amp;C Costs: ${base_dc:,.0f}/ft
            </li>
            <li>
                Type Curves: {tc_name_text}{"s" if len(tc_names) > 1 else ""}; {avg_tc_risk:.0%} base case TC risk
            </li>
            <li>
                Development: {total_wells:,.1f} gross wells with ~{avg_ll:,.0f}' average lateral length
            </li>
            <li>
                Unitization: {avg_pct_unitized:.0%} of acres assumed to be unitized
            </li>
            <li>
                Entry: ${base_bid:,.0f}/acre
            </li>
            <li>
                Returns: <b>{irr:.1%} IRR</b> and <b>{moic:.2f}x MOIC</b>
            </li>
        </ul>
    
        <p>Let us know if you have any questions.</p>
    
        <h3 style="margin-bottom:8px;">Type Curve Assumptions:</h3>
        {tc_table_html}
    
        <h3 style="margin-top:24px; margin-bottom:8px;">Quarterly Output:</h3>
        {quarterly_table_html}
    
        <div style="margin-top:24px;">
            {sensitivities_html}
        </div>
    
        <div style="margin-top:24px;">
            {charts_html}
        </div>
    </body>
    </html>
    """
    return html

# -----------------------------
# Session state init
# -----------------------------
# Bump this value whenever the editor schema or stored model-result schema
# changes. Streamlit can retain old widget/session values across a hot reload,
# which may leave the page blank or stuck after a deployment.
APP_STATE_VERSION = "sensitivity-axis-toggle-v2"

if st.session_state.get("_app_state_version") != APP_STATE_VERSION:
    # Clear stale data-editor widget state and prior calculated outputs.
    # Preserve the user's slot table when possible, but sanitize the new
    # $1/acre minimum before the editor renders.
    existing_slot_df = st.session_state.get("slot_df")

    for stale_key in [
        "slot_editor",
        "slot_editor_v2",
        "model_deal_inputs",
        "deal_df",
        "all_slots_df",
        "model_slot_df",
        "slot_audit_df",
        "deal_audit_df",
        "irr",
        "moic",
        "deal_log_overview_df",
        "deal_log_slot_df",
        "deal_log_filename",
    ]:
        st.session_state.pop(stale_key, None)

    if isinstance(existing_slot_df, pd.DataFrame):
        existing_slot_df = existing_slot_df.copy()
        if "bid_per_acre" in existing_slot_df.columns:
            existing_slot_df["bid_per_acre"] = (
                pd.to_numeric(
                    existing_slot_df["bid_per_acre"],
                    errors="coerce",
                )
                .fillna(1.0)
                .clip(lower=1.0)
            )
        st.session_state["slot_df"] = existing_slot_df

    st.session_state["model_has_run"] = False
    st.session_state["_app_state_version"] = APP_STATE_VERSION
if "slot_df" not in st.session_state:
    st.session_state["slot_df"] = build_slot_template(2)

if "include_slot" not in st.session_state["slot_df"].columns:
    st.session_state["slot_df"].insert(0, "include_slot", True)
    
if "dale_promote" not in st.session_state["slot_df"].columns:
    st.session_state["slot_df"].insert(1, "dale_promote", False)

if "carry_enabled" not in st.session_state["slot_df"].columns:
    st.session_state["slot_df"].insert(2, "carry_enabled", False)

if "carry_dnc_pct" in st.session_state["slot_df"].columns:
    st.session_state["slot_df"] = st.session_state["slot_df"].drop(
        columns=["carry_dnc_pct"]
    )

if "carry_wi_reversion_pct" not in st.session_state["slot_df"].columns:
    st.session_state["slot_df"]["carry_wi_reversion_pct"] = 0.0

if "model_deal_inputs" not in st.session_state:
    st.session_state["model_deal_inputs"] = None

if "deal_df" not in st.session_state:
    st.session_state["deal_df"] = None

if "all_slots_df" not in st.session_state:
    st.session_state["all_slots_df"] = None

if "model_slot_df" not in st.session_state:
    st.session_state["model_slot_df"] = None

if "irr" not in st.session_state:
    st.session_state["irr"] = None

if "moic" not in st.session_state:
    st.session_state["moic"] = None

if "model_has_run" not in st.session_state:
    st.session_state["model_has_run"] = False

# If the app has old results from before model_slot_df existed,
# force the user to rerun the model instead of trying to copy None.
if st.session_state["model_has_run"] and st.session_state["model_slot_df"] is None:
    st.session_state["model_has_run"] = False
    st.session_state["deal_df"] = None
    st.session_state["all_slots_df"] = None

if "heavy_outputs_disabled" not in st.session_state:
    st.session_state["heavy_outputs_disabled"] = False

# -----------------------------
# Sidebar deal inputs
# -----------------------------
st.sidebar.header("Deal-Level Inputs")

st.sidebar.subheader("Output Mode")

heavy_outputs_label = (
    "Enable Sensitivities / Charts / Email"
    if st.session_state["heavy_outputs_disabled"]
    else "Disable Sensitivities / Charts / Email"
)

if st.sidebar.button(
    heavy_outputs_label,
    use_container_width=True,
    type="secondary",
):
    st.session_state["heavy_outputs_disabled"] = not st.session_state["heavy_outputs_disabled"]
    st.rerun()

disable_heavy_outputs = st.session_state["heavy_outputs_disabled"]

if disable_heavy_outputs:
    st.sidebar.info(
        "Fast output mode is on. Sensitivities, charts, and email export are disabled."
    )

st.sidebar.subheader("Sensitivity Setup")
use_tc_risk_as_main_sensitivity = st.sidebar.toggle(
    "Use TC Risk as Main Sensitivity Variable",
    value=False,
    help=(
        "Off = the sensitivity tables use $/Acre Bid as the main row variable. "
        "On = the sensitivity tables use TC Risk as the main row variable."
    ),
)

st.sidebar.caption(
    "Main sensitivity variable: "
    + ("TC Risk" if use_tc_risk_as_main_sensitivity else "$/Acre Bid")
)

use_dc_pct_sensitivity = st.sidebar.toggle(
    "Use 5% D&C Sensitivity Steps",
    value=False,
    help=(
        "Off = D&C sensitivity changes in $50/ft increments. "
        "On = D&C sensitivity changes in 5% increments around the base D&C. "
        "With nine points, this covers 80% to 120% of the base D&C."
    ),
)

st.sidebar.caption(
    "D&C sensitivity steps: "
    + ("5% of base D&C" if use_dc_pct_sensitivity else "$50/ft")
)

st.sidebar.subheader("Timing")
effective_date = st.sidebar.date_input("Effective Date", value=next_month_start())

st.sidebar.subheader("Pricing")

use_monthly_pricing_file = st.sidebar.toggle(
    "Use Monthly Pricing File",
    value=False,
    help=(
        "Off = use flat oil and gas prices for the entire model. "
        "On = use price_file_library.xlsx until each commodity's "
        "selected flat-pricing date."
    ),
)

pricing_mode = (
    "file"
    if use_monthly_pricing_file
    else "flat"
)

if pricing_mode == "flat":
    oil_col, gas_col = st.sidebar.columns(2)

    with oil_col:
        oil_price = st.number_input(
            "Oil Price ($/bbl)",
            min_value=0.0,
            value=60.0,
            step=1.0,
            format="%.2f",
            key="flat_mode_oil_price",
        )

    with gas_col:
        gas_price = st.number_input(
            "Gas Price ($/mcf)",
            min_value=0.0,
            value=3.75,
            step=0.05,
            format="%.3f",
            key="flat_mode_gas_price",
        )

    # These dates are ignored in flat mode.
    oil_flat_start_date = date(1900, 1, 1)
    gas_flat_start_date = date(1900, 1, 1)

else:
    try:
        pricing_preview_df = load_price_file(
            PRICE_FILE_PATH
        )

    except (FileNotFoundError, ValueError) as error:
        st.sidebar.error(str(error))
        st.stop()

    first_pricing_month = (
        pricing_preview_df["month"].min()
    )

    last_pricing_month = (
        pricing_preview_df["month"].max()
    )

    default_flat_start = (
        last_pricing_month
        + pd.offsets.MonthBegin(1)
    ).date()

    st.sidebar.caption(
        "Pricing file range: "
        f"{first_pricing_month:%b %Y} through "
        f"{last_pricing_month:%b %Y}"
    )

    oil_col, gas_col = st.sidebar.columns(2)

    with oil_col:
        st.markdown("**Oil**")

        oil_flat_start_date = st.date_input(
            "Switch to Flat",
            value=default_flat_start,
            format="MM/DD/YYYY",
            key="oil_flat_start_date",
        )

        oil_price = st.number_input(
            "Flat Oil ($/bbl)",
            min_value=0.0,
            value=60.0,
            step=1.0,
            format="%.2f",
            key="terminal_oil_price",
        )

    with gas_col:
        st.markdown("**Gas**")

        gas_flat_start_date = st.date_input(
            "Switch to Flat",
            value=default_flat_start,
            format="MM/DD/YYYY",
            key="gas_flat_start_date",
        )

        gas_price = st.number_input(
            "Flat Gas ($/mcf)",
            min_value=0.0,
            value=3.75,
            step=0.05,
            format="%.3f",
            key="terminal_gas_price",
        )

st.sidebar.subheader("Overrides")

use_acquisition_override = st.sidebar.checkbox("Use Acquisition Cost Override", value=False)
acquisition_cost_override = st.sidebar.number_input(
    "Acquisition Cost Override",
    min_value=0.0,
    value=0.0,
    step=1000.0,
    format="%.1f",
    disabled=not use_acquisition_override,
)

use_dc_override = st.sidebar.checkbox("Use D&C Override for All Slots", value=False)
dc_override = st.sidebar.number_input(
    "D&C Override ($/ft)",
    value=750.0,
    step=25.0,
    disabled=not use_dc_override,
)

use_bid_override = st.sidebar.checkbox("Use $/Acre Override for All Slots", value=False)
bid_override = st.sidebar.number_input(
    "$/Acre Override",
    min_value=1.0,
    value=8000.0,
    step=250.0,
    disabled=not use_bid_override,
)

st.sidebar.subheader("Taxes")

use_sev_tax_pct = st.sidebar.toggle(
    "Severance Tax as % of Net Revenue",
    value=False,
    help=(
        "Off = fixed $/bbl for oil and $/mcf for gas. "
        "On = percentage of net oil revenue and net gas revenue."
    ),
)

oil_tax_col, gas_tax_col = st.sidebar.columns(2)

if use_sev_tax_pct:
    with oil_tax_col:
        oil_sev_tax = st.number_input(
            "Oil Sev. Tax (%)",
            min_value=0.0,
            value=0.0,
            step=0.25,
            format="%.3f",
            key="oil_sev_tax_pct_input",
            help="Enter 5 for 5% of net oil revenue.",
        )

    with gas_tax_col:
        gas_sev_tax = st.number_input(
            "Gas Sev. Tax (%)",
            min_value=0.0,
            value=0.0,
            step=0.25,
            format="%.3f",
            key="gas_sev_tax_pct_input",
            help="Enter 5 for 5% of net gas revenue.",
        )

else:
    with oil_tax_col:
        oil_sev_tax = st.number_input(
            "Oil Sev. Tax ($/bbl)",
            min_value=0.0,
            value=0.10,
            step=0.01,
            format="%.3f",
            key="oil_sev_tax_fixed_input",
        )

    with gas_tax_col:
        gas_sev_tax = st.number_input(
            "Gas Sev. Tax ($/mcf)",
            min_value=0.0,
            value=0.025,
            step=0.005,
            format="%.3f",
            key="gas_sev_tax_fixed_input",
        )

ad_val_tax = st.sidebar.number_input(
    "Ad Valorem Tax (% of Net Revenue)",
    value=0.025,
    step=0.005,
    format="%.3f",
)

st.sidebar.subheader("Ethane / NGL")
ethane_rec = st.sidebar.checkbox("Recover Ethane", value=False)

with st.sidebar.expander("Content Percentages", expanded=False):
    content_ethane = st.number_input("Ethane Content %", value=0.50, step=0.01, format="%.3f")
    content_propane = st.number_input("Propane Content %", value=0.25, step=0.01, format="%.3f")
    content_isobutane = st.number_input("Isobutane Content %", value=0.065, step=0.005, format="%.3f")
    content_butane = st.number_input("Butane Content %", value=0.065, step=0.005, format="%.3f")
    content_pentanes = st.number_input("Pentanes Content %", value=0.12, step=0.01, format="%.3f")

with st.sidebar.expander("Recover Ethane Percentages", expanded=False):
    rec_ethane = st.number_input("Recover Ethane %", value=0.90, step=0.01, format="%.3f")
    rec_propane = st.number_input("Recover Propane %", value=0.98, step=0.01, format="%.3f")
    rec_isobutane = st.number_input("Recover Isobutane %", value=0.99, step=0.01, format="%.3f")
    rec_butane = st.number_input("Recover Butane %", value=0.99, step=0.01, format="%.3f")
    rec_pentanes = st.number_input("Recover Pentanes %", value=0.995, step=0.001, format="%.3f")

with st.sidebar.expander("Reject Ethane Percentages", expanded=False):
    rej_ethane = st.number_input("Reject Ethane %", value=0.20, step=0.01, format="%.3f")
    rej_propane = st.number_input("Reject Propane %", value=0.90, step=0.01, format="%.3f")
    rej_isobutane = st.number_input("Reject Isobutane %", value=0.98, step=0.01, format="%.3f")
    rej_butane = st.number_input("Reject Butane %", value=0.98, step=0.01, format="%.3f")
    rej_pentanes = st.number_input("Reject Pentanes %", value=0.995, step=0.001, format="%.3f")

with st.sidebar.expander("NGL Shrink Factors", expanded=False):
    shrink_ethane = st.number_input("Ethane Shrink", value=0.06634, step=0.001, format="%.5f")
    shrink_propane = st.number_input("Propane Shrink", value=0.091563, step=0.001, format="%.5f")
    shrink_isobutane = st.number_input("Isobutane Shrink", value=0.09963, step=0.001, format="%.5f")
    shrink_butane = st.number_input("Butane Shrink", value=0.103744, step=0.001, format="%.5f")
    shrink_pentanes = st.number_input("Pentanes Shrink", value=0.10968, step=0.001, format="%.5f")

with st.sidebar.expander("NGL Component Prices", expanded=False):
    price_ethane = st.number_input("Ethane Price", value=0.27, step=0.01, format="%.5f")
    price_propane = st.number_input("Propane Price", value=0.64625, step=0.01, format="%.5f")
    price_isobutane = st.number_input("Isobutane Price", value=0.84, step=0.01, format="%.5f")
    price_butane = st.number_input("Butane Price", value=0.7825, step=0.01, format="%.5f")
    price_pentanes = st.number_input("Pentanes Price", value=1.22125, step=0.01, format="%.5f")

st.sidebar.subheader("WI Reversion")

st.sidebar.caption(
    "For selected slots, the reversion transfers a percentage of our then-current "
    "working interest after cumulative distributions reach the selected multiple "
    "of acquisition cost plus funded D&C."
)

dale_promote_override = st.sidebar.checkbox(
    "WI Reversion Override - Apply to All Slots",
    value=False,
)

promote_wi_reversion_pct = st.sidebar.number_input(
    "WI Given Up at Reversion (%)",
    min_value=0.0,
    max_value=100.0,
    value=6.25,
    step=0.25,
    format="%.2f",
    help=(
        "Enter 12.5 to transfer 12.5% of our then-current WI. "
        "The transfer permanently begins the month after the hurdle is reached."
    ),
)

promote_multiple = st.sidebar.number_input(
    "Reversion Investment Multiple",
    min_value=0.01,
    value=1.00,
    step=0.05,
    format="%.2f",
    help=(
        "Cumulative positive operating distributions divided by cumulative "
        "acquisition cost plus funded D&C for reversion-eligible slots."
    ),
)

deal_inputs = {
    "use_acquisition_override": use_acquisition_override,
    "acquisition_cost_override": acquisition_cost_override,
    "effective_date": effective_date,
    "oil_price": oil_price,
    "gas_price": gas_price,
    "pricing_mode": pricing_mode,
    "pricing_file_path": PRICE_FILE_PATH,

    # Flat prices in flat mode.
    # Terminal flat prices in file mode.
    "oil_price": float(oil_price),
    "gas_price": float(gas_price),

    # Preserve the original base case for pricing sensitivities.
    "base_oil_price": float(oil_price),
    "base_gas_price": float(gas_price),

    "oil_flat_start_date": oil_flat_start_date,
    "gas_flat_start_date": gas_flat_start_date,
    "use_dc_override": use_dc_override,
    "dc_override": dc_override,
    "use_bid_override": use_bid_override,
    "bid_override": bid_override,
    "use_tc_risk_as_main_sensitivity": use_tc_risk_as_main_sensitivity,
    "use_dc_pct_sensitivity": use_dc_pct_sensitivity,
    "use_sev_tax_pct": use_sev_tax_pct,
    "oil_sev_tax": oil_sev_tax,
    "gas_sev_tax": gas_sev_tax,
    "ad_val_tax": ad_val_tax,
    "ethane_rec": ethane_rec,
    "content_ethane": content_ethane,
    "content_propane": content_propane,
    "content_isobutane": content_isobutane,
    "content_butane": content_butane,
    "content_pentanes": content_pentanes,
    "rec_ethane": rec_ethane,
    "rec_propane": rec_propane,
    "rec_isobutane": rec_isobutane,
    "rec_butane": rec_butane,
    "rec_pentanes": rec_pentanes,
    "rej_ethane": rej_ethane,
    "rej_propane": rej_propane,
    "rej_isobutane": rej_isobutane,
    "rej_butane": rej_butane,
    "rej_pentanes": rej_pentanes,
    "shrink_ethane": shrink_ethane,
    "shrink_propane": shrink_propane,
    "shrink_isobutane": shrink_isobutane,
    "shrink_butane": shrink_butane,
    "shrink_pentanes": shrink_pentanes,
    "price_ethane": price_ethane,
    "price_propane": price_propane,
    "price_isobutane": price_isobutane,
    "price_butane": price_butane,
    "price_pentanes": price_pentanes,
    "dale_promote_override": dale_promote_override,
    "promote_enabled": False,  # set right before model run based on selected slots
    "promote_wi_reversion_pct": promote_wi_reversion_pct,
    "promote_multiple": promote_multiple,
}


# -----------------------------
# Slot controls
# -----------------------------
st.subheader("Type Curve Assumptions")

file_mtime = os.path.getmtime("type_curve_library.xlsx")
tc_names = ["Choose TC"] + load_tc_names(file_mtime)

# Keep the slot-count controls in the same form as the editor.
# This lets "Load Slots" receive unsaved changes currently in the grid.
with st.form("slot_inputs_form"):
    col1, col2, col3 = st.columns([2, 1, 1])

    with col1:
        num_slots = st.number_input(
            "Number of Slots",
            min_value=1,
            step=1,
            value=len(st.session_state["slot_df"]),
        )

    with col2:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        load_slots_clicked = st.form_submit_button(
            "Load Slots",
            use_container_width=True,
            type="primary",
        )

    with col3:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        refresh_tc_clicked = st.form_submit_button(
            "Refresh Type Curves",
            use_container_width=True,
            type="secondary",
        )
        
    edited_slot_df = st.data_editor(
    st.session_state["slot_df"],
    num_rows="fixed",
    use_container_width=True,
    key="slot_editor_v2",
    column_order=[
        "include_slot",
        "dale_promote",
        "carry_enabled",
        "carry_wi_reversion_pct",
        "slot_id",
        "tc_name",
        "gross_wells",
        "net_acres",
        "unit_acres",
        "use_calc_unit_acres",
        "pct_unitized",
        "drilling_spud_month",
        "flowback_delay",
        "net_revenue_interest",
        "lateral_length",
        "dc_costs",
        "tc_risk",
        "bid_per_acre",
        "oil_diff",
        "gas_diff",
        "oil_opex_bbl",
        "gas_opex_mcf",
        "ngl_opex",
        "fixed_loe",
        "ngl_yield",
    ],
    column_config={
        "include_slot": st.column_config.CheckboxColumn(
            "Include",
            help="Include this slot in the model run.",
            default=True,
        ),
        "dale_promote": st.column_config.CheckboxColumn(
            "WI Reversion",
            help=(
                "Pool this slot's acquisition, funded D&C, and positive operating "
                "distributions in the reversion test. After the hurdle is reached, the selected "
                "portion of our then-current WI transfers permanently."
            ),
            default=False,
        ),
        "carry_enabled": st.column_config.CheckboxColumn(
            "Carry",
            help=(
                "We fund D&C at our original WI, then surrender a portion "
                "of our original WI beginning with first production."
            ),
            default=False,
        ),
        "carry_wi_reversion_pct": st.column_config.NumberColumn(
            "WI Given Up (%)",
            min_value=0.0,
            max_value=100.0,
            step=1.0,
            format="%.0f%%",
            help="Enter 35 for a 35% WI reversion.",
        ),
        "slot_id": st.column_config.NumberColumn("Slot", format="%d", disabled=True),
        "tc_name": st.column_config.SelectboxColumn("Type Curve", options=tc_names, required=True),
            "gross_wells": st.column_config.NumberColumn("Gross Wells", format="%.2f"),
            "net_acres": st.column_config.NumberColumn(
                "Net Acres",
                format="%.2f",
            ),
            "unit_acres": st.column_config.NumberColumn("Unit Acres", format="%,.0f"),
            "use_calc_unit_acres": st.column_config.CheckboxColumn("Calc Unit Acres"),
            "pct_unitized": st.column_config.NumberColumn("% Unitized", format="%.2f"),
            "drilling_spud_month": st.column_config.DateColumn("Spud Month", format="YYYY-MM-DD"),
            "flowback_delay": st.column_config.NumberColumn("Flowback Delay", format="%d"),
            "net_revenue_interest": st.column_config.NumberColumn("NRI", format="%.2f"),
            "lateral_length": st.column_config.NumberColumn("Lateral Length (ft)", format="%,d"),
            "dc_costs": st.column_config.NumberColumn("D&C ($/ft)", format="$%,.0f"),
            "tc_risk": st.column_config.NumberColumn(
                "TC Risk",
                min_value=0.0,
                step=0.01,
                format="%.2f",
                help="Enter as decimal, e.g. 0.97 = 97% TC risk.",
            ),
            "bid_per_acre": st.column_config.NumberColumn(
                "$/Acre Bid",
                min_value=1.0,
                step=250.0,
                format="$%,d",
                help="Minimum allowed bid is $1 per acre.",
            ),
            "oil_diff": st.column_config.NumberColumn("Oil Diff", format="$%.2f"),
            "gas_diff": st.column_config.NumberColumn("Gas Diff", format="$%.2f"),
            "ngl_diff": None,
            "oil_opex_bbl": st.column_config.NumberColumn("Oil Opex", format="$%.2f"),
            "gas_opex_mcf": st.column_config.NumberColumn("Gas Opex", format="$%.2f"),
            "ngl_opex": st.column_config.NumberColumn("NGL Opex", format="$%.2f"),
            "fixed_loe": st.column_config.NumberColumn("Fixed LOE", format="$%,.0f"),
            "ngl_yield": st.column_config.NumberColumn("NGL Yield", format="%.2f"),
        },
    ).copy()

    apply_slot_changes = st.form_submit_button(
        "Apply Slot Changes",
        type="primary",
    )

if load_slots_clicked or refresh_tc_clicked or apply_slot_changes:
    cleaned_slot_df = edited_slot_df.copy()

    cleaned_slot_df["tc_risk"] = pd.to_numeric(
        cleaned_slot_df["tc_risk"],
        errors="coerce",
    ).fillna(1.0).astype(float)

    cleaned_slot_df["carry_enabled"] = (
        cleaned_slot_df["carry_enabled"]
        .fillna(False)
        .astype(bool)
    )

    cleaned_slot_df["carry_wi_reversion_pct"] = (
        pd.to_numeric(
            cleaned_slot_df["carry_wi_reversion_pct"],
            errors="coerce",
        )
        .fillna(0.0)
        .clip(lower=0.0, upper=100.0)
    )

    cleaned_slot_df["bid_per_acre"] = (
        pd.to_numeric(
            cleaned_slot_df["bid_per_acre"],
            errors="coerce",
        )
        .fillna(1.0)
        .clip(lower=1.0)
    )

    cleaned_slot_df = apply_calc_unit_acres(cleaned_slot_df)

    if load_slots_clicked:
        cleaned_slot_df = resize_slot_df(
            cleaned_slot_df,
            int(num_slots),
        )

    st.session_state["slot_df"] = cleaned_slot_df
    st.session_state["model_has_run"] = False

    if refresh_tc_clicked:
        load_tc_names.clear()

    # Rebuild the editor immediately when rows or dropdown options change.
    if load_slots_clicked or refresh_tc_clicked:
        st.rerun()

slot_df = st.session_state["slot_df"].copy()

run_model_clicked = st.button("Run Model", type="primary")

if run_model_clicked:
    st.session_state["slot_df"] = slot_df

    included_slot_df = slot_df[slot_df["include_slot"].fillna(True)].copy()

    if included_slot_df.empty:
        st.warning("Please include at least one slot before running the model.")
        st.session_state["model_has_run"] = False

    elif (included_slot_df["tc_name"] == "Choose TC").any():
        st.warning("Please select a Type Curve for all included slots before running the model.")
        st.session_state["model_has_run"] = False

    else:
        model_slot_df = included_slot_df.drop(columns=["include_slot"], errors="ignore").copy()
        model_slot_df["bid_per_acre"] = (
            pd.to_numeric(
                model_slot_df["bid_per_acre"],
                errors="coerce",
            )
            .fillna(1.0)
            .clip(lower=1.0)
        )
        
        if deal_inputs.get("dale_promote_override", False):
            model_slot_df["dale_promote"] = True
        
        run_deal_inputs = deal_inputs.copy()
        run_deal_inputs["promote_enabled"] = bool(
            model_slot_df["dale_promote"].fillna(False).any()
        )
        
        if not run_deal_inputs["promote_enabled"]:
            run_deal_inputs["promote_wi_reversion_pct"] = 0.0
            run_deal_inputs["promote_multiple"] = 0.0
        
        all_slots_df, deal_df, slot_audit_df, deal_audit_df, irr, moic = run_deal_model(
            model_slot_df,
            run_deal_inputs,
        )

        # ------------------------------------------------------------
        # COMPLIANCE DEAL LOG EXPORT — SAVE EXACT INPUTS USED THIS RUN
        # ------------------------------------------------------------
        run_timestamp = pd.Timestamp.now(tz="America/Chicago")
        
        overview_rows = [
            {
                "Section": "Run Information",
                "Input": "Model Run Created",
                "Value": run_timestamp.strftime("%m/%d/%Y %I:%M %p %Z"),
            },
            {
                "Section": "Run Information",
                "Input": "IRR",
                "Value": irr,
            },
            {
                "Section": "Run Information",
                "Input": "MOIC",
                "Value": moic,
            },
        ]
        
        # Add every main sidebar / deal assumption used in the run
        for key, value in run_deal_inputs.items():
            overview_rows.append(
                {
                    "Section": "Sidebar Inputs",
                    "Input": str(key).replace("_", " ").title(),
                    "Value": value,
                }
            )
        
        overview_df = pd.DataFrame(overview_rows)
        
        # Save the exact run information. The CSV itself is assembled later
        # so it can include the opportunity name entered in the export section.
        deal_log_filename = (
            f"Utica_Deal_Log_"
            f"{run_timestamp.strftime('%Y%m%d_%H%M%S')}.csv"
        )
        
        st.session_state["deal_log_overview_df"] = overview_df
        st.session_state["deal_log_slot_df"] = model_slot_df.copy()
        st.session_state["deal_log_filename"] = deal_log_filename
        st.session_state["model_deal_inputs"] = run_deal_inputs
        st.session_state["model_slot_df"] = model_slot_df
        st.session_state["all_slots_df"] = all_slots_df
        st.session_state["deal_df"] = deal_df
        st.session_state["slot_audit_df"] = slot_audit_df
        st.session_state["deal_audit_df"] = deal_audit_df
        st.session_state["irr"] = irr
        st.session_state["moic"] = moic
        st.session_state["model_has_run"] = True


# -----------------------------
# Results
# -----------------------------
DEAL_DISPLAY_COLS = [
    "date",
    "index_oil_price",
    "index_gas_price",
    "promote_cumulative_investment",
    "promote_cumulative_distributions",
    "promote_running_multiple",
    "promote_hurdle_reached",
    "promote_active",
    "promote_hurdle_date",
    "promote_effective_date",
    "slot_net_oil_production",
    "slot_net_gas_production",
    "slot_net_ngl_production",
    "slot_oil_revenue",
    "slot_gas_revenue",
    "slot_ngl_revenue",
    "slot_total_revenue",
    "slot_loe",
    "slot_tax",
    "slot_operating_profit",
    "slot_capex",
    "slot_asset_purchase",
    "slot_total_cash_flow",
    "cum_total_cf",
]

SLOT_DISPLAY_COLS = [
    "slot_id",
    "tc_name",
    "date",
    "economic_limit_reached",
    "well_shut_in",
    "pre_shut_in_operating_cf",
    "dale_promote",
    "pre_promote_working_interest",
    "promote_wi_transferred",
    "post_promote_working_interest",
    "effective_working_interest",
    "promote_running_multiple",
    "promote_hurdle_reached",
    "promote_active",
    "promote_hurdle_date",
    "promote_effective_date",
    "slot_net_oil_production",
    "slot_net_gas_production",
    "slot_net_ngl_production",
    "slot_oil_revenue",
    "slot_gas_revenue",
    "slot_ngl_revenue",
    "slot_total_revenue",
    "slot_loe",
    "slot_tax",
    "slot_operating_profit",
    "slot_capex",
    "slot_asset_purchase",
    "slot_total_cash_flow",
    "cum_total_cf",
]

if (
    st.session_state["model_has_run"]
    and st.session_state["deal_df"] is not None
    and st.session_state["all_slots_df"] is not None
    and st.session_state["model_slot_df"] is not None
):
    all_slots_df = st.session_state["all_slots_df"]
    deal_df = st.session_state["deal_df"]
    irr = st.session_state["irr"]
    moic = st.session_state["moic"]
    deal_audit_df = st.session_state["deal_audit_df"]
    slot_audit_df = st.session_state["slot_audit_df"]
    slot_df = st.session_state["model_slot_df"].copy()
    deal_inputs = st.session_state.get("model_deal_inputs", deal_inputs).copy()

    # This toggle only changes how sensitivities are displayed, so let it
    # update immediately without requiring the base model to be rerun.
    deal_inputs["use_tc_risk_as_main_sensitivity"] = (
        use_tc_risk_as_main_sensitivity
    )
    deal_inputs["use_dc_pct_sensitivity"] = use_dc_pct_sensitivity

    deal_display_df = deal_audit_df[[col for col in DEAL_DISPLAY_COLS if col in deal_audit_df.columns]].copy()
    slot_display_df = slot_audit_df[[col for col in SLOT_DISPLAY_COLS if col in slot_audit_df.columns]].copy()

    deal_audit_display_df = format_display_df(deal_display_df)
    slot_audit_display_df = format_display_df(slot_display_df)
    audit_excel_data = to_excel_bytes(deal_audit_df, slot_audit_df)

    with st.expander("Monthly Data", expanded=False):
        st.subheader("Total Deal Monthly Data")
        st.dataframe(deal_audit_display_df, use_container_width=True)

        st.subheader("Type Curve Monthly Data")
        st.dataframe(slot_audit_display_df, use_container_width=True)

        st.download_button(
            "Download in Excel",
            audit_excel_data,
            file_name="deal_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="monthly_data_download_excel",
        )

    st.subheader("Deal Summary")

    total_net_acres = slot_df["net_acres"].sum()

    if deal_inputs["use_bid_override"]:
        total_acquisition = total_net_acres * deal_inputs["bid_override"]
    else:
        total_acquisition = (slot_df["net_acres"] * slot_df["bid_per_acre"]).sum()

    blended_bid = total_acquisition / total_net_acres if total_net_acres > 0 else 0.0

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.metric("Total Net Acres", format_accounting_number(total_net_acres, decimals=1))
    with col2:
        total_acquisition_cost = -deal_df["slot_asset_purchase"].sum()
        st.metric("Acquisition Cost", format_thousands_short(total_acquisition_cost, decimals=1))
    with col3:
        st.metric("$/Acre Bid", format_accounting_number(blended_bid, decimals=0, prefix="$"))
    with col4:
        st.metric("IRR", format_accounting_percent(irr, decimals=1, zero_as_dash=False) if irr is not None else "N/A")
    with col5:
        st.metric("MOIC", format_accounting_number(moic, decimals=2, suffix="x", zero_as_dash=False) if moic is not None else "N/A")

    if bool(deal_inputs.get("promote_enabled", False)):
        active_dates = deal_df.loc[
            deal_df.get("promote_active", False).fillna(False),
            "date",
        ] if "promote_active" in deal_df.columns else pd.Series(dtype="datetime64[ns]")

        if not active_dates.empty:
            promote_effective_date = pd.to_datetime(active_dates.iloc[0])
            transferred_pct = float(
                deal_inputs.get("promote_wi_reversion_pct", 0.0)
            )
            st.info(
                f"WI reversion becomes effective {promote_effective_date:%m/%d/%Y}: "
                f"{transferred_pct:.2f}% of our then-current WI transfers after the "
                f"{float(deal_inputs.get('promote_multiple', 0.0)):.2f}x hurdle."
            )
        else:
            st.info(
                "WI reversion is enabled but does not reach its investment "
                "multiple during the modeled period."
            )

    base_dc = (
        deal_inputs["dc_override"]
        if deal_inputs["use_dc_override"]
        else weighted_avg_by_net_acres(slot_df, "dc_costs")
    )
    
    base_bid = max(
        1.0,
        (
            deal_inputs["bid_override"]
            if deal_inputs["use_bid_override"]
            else weighted_avg_by_net_acres(slot_df, "bid_per_acre")
        ),
    )

    if not disable_heavy_outputs:
        st.subheader("Sensitivity Tables")

        scenario_scatter_chart = build_scenario_scatter_chart(
            slot_df=slot_df,
            deal_inputs=deal_inputs,
            base_bid=base_bid,
            base_dc=base_dc,
        )
        
        base_tc_risk = weighted_avg_by_net_acres(slot_df, "tc_risk")
        base_ngl_yield = weighted_avg_by_net_acres(slot_df, "ngl_yield")

        bid_values = build_sensitivity_range(
            base_bid,
            500.0,
            4,
            min_value=1.0,
        )
        tc_risk_values = [
            max(0.0, base_tc_risk + 0.05 * i)
            for i in range(-4, 5)
        ]
        if bool(deal_inputs.get("use_dc_pct_sensitivity", False)):
            dc_values = build_percentage_sensitivity_range(
                base_dc,
                pct_step=0.05,
                steps_each_way=4,
                min_value=0.0,
            )
        else:
            dc_values = build_sensitivity_range(
                base_dc,
                50.0,
                4,
                min_value=0.0,
            )

        oil_values = build_sensitivity_range(
            float(deal_inputs["oil_price"]),
            5.0,
            4,
        )
        gas_values = build_sensitivity_range(
            float(deal_inputs["gas_price"]),
            0.25,
            4,
        )
        ngl_yield_values = build_sensitivity_range(
            base_value=base_ngl_yield,
            step=0.50,
            steps_each_way=4,
            min_value=0.0,
        )
        base_spud_month = weighted_avg_spud_month_by_net_acres(slot_df)
        spud_date_values = [
            base_spud_month + pd.DateOffset(months=3 * i)
            for i in range(-4, 5)
        ]

        use_tc_risk_main = bool(
            deal_inputs.get("use_tc_risk_as_main_sensitivity", False)
        )

        if use_tc_risk_main:
            main_values = tc_risk_values
            main_variable = "tc_risk"
            main_title = "TC Risk"
            main_format = "percent"
            main_base = base_tc_risk

            cross_x_values = bid_values
            cross_x_variable = "bid"
            cross_x_title = "$/Acre Bid"
            cross_x_format = "dollar"
            cross_base_x = base_bid
        else:
            main_values = bid_values
            main_variable = "bid"
            main_title = "$/Acre Bid"
            main_format = "dollar"
            main_base = base_bid

            cross_x_values = tc_risk_values
            cross_x_variable = "tc_risk"
            cross_x_title = "TC Risk"
            cross_x_format = "percent"
            cross_base_x = base_tc_risk

        irr_sens_df, moic_sens_df = run_two_way_sensitivity(
            slot_df=slot_df,
            deal_inputs=deal_inputs,
            x_values=dc_values,
            x_variable="dc",
            y_values=main_values,
            y_variable=main_variable,
        )

        irr_heatmap = build_heatmap(
            irr_sens_df,
            "IRR Sensitivity",
            metric="irr",
            x_title="D&C Costs ($/ft)",
            y_title=main_title,
            x_format="dollar",
            y_format=main_format,
            base_x=base_dc,
            base_y=main_base,
        )

        moic_heatmap = build_heatmap(
            moic_sens_df,
            "MOIC Sensitivity",
            metric="moic",
            x_title="D&C Costs ($/ft)",
            y_title=main_title,
            x_format="dollar",
            y_format=main_format,
            base_x=base_dc,
            base_y=main_base,
        )

        irr_oil_bid_df, moic_oil_bid_df = run_two_way_sensitivity(
            slot_df=slot_df,
            deal_inputs=deal_inputs,
            x_values=oil_values,
            x_variable="oil",
            y_values=main_values,
            y_variable=main_variable,
        )

        irr_gas_bid_df, moic_gas_bid_df = run_two_way_sensitivity(
            slot_df=slot_df,
            deal_inputs=deal_inputs,
            x_values=gas_values,
            x_variable="gas",
            y_values=main_values,
            y_variable=main_variable,
        )

        irr_ngl_yield_bid_df, moic_ngl_yield_bid_df = run_two_way_sensitivity(
            slot_df=slot_df,
            deal_inputs=deal_inputs,
            x_values=ngl_yield_values,
            x_variable="ngl_yield",
            y_values=main_values,
            y_variable=main_variable,
        )

        irr_oil_bid_heatmap = build_heatmap(
            irr_oil_bid_df,
            "IRR Sensitivity",
            metric="irr",
            x_title="Oil Price ($/bbl)",
            y_title=main_title,
            x_format="dollar",
            y_format=main_format,
            base_x=deal_inputs["oil_price"],
            base_y=main_base,
        )

        moic_oil_bid_heatmap = build_heatmap(
            moic_oil_bid_df,
            "MOIC Sensitivity",
            metric="moic",
            x_title="Oil Price ($/bbl)",
            y_title=main_title,
            x_format="dollar",
            y_format=main_format,
            base_x=deal_inputs["oil_price"],
            base_y=main_base,
        )

        irr_gas_bid_heatmap = build_heatmap(
            irr_gas_bid_df,
            "IRR Sensitivity",
            metric="irr",
            x_title="Gas Price ($/mcf)",
            y_title=main_title,
            x_format="dollar",
            y_format=main_format,
            base_x=deal_inputs["gas_price"],
            base_y=main_base,
        )

        moic_gas_bid_heatmap = build_heatmap(
            moic_gas_bid_df,
            "MOIC Sensitivity",
            metric="moic",
            x_title="Gas Price ($/mcf)",
            y_title=main_title,
            x_format="dollar",
            y_format=main_format,
            base_x=deal_inputs["gas_price"],
            base_y=main_base,
        )

        irr_ngl_yield_bid_heatmap = build_heatmap(
            irr_ngl_yield_bid_df,
            "IRR Sensitivity",
            metric="irr",
            x_title="NGL Yield (GPM)",
            y_title=main_title,
            x_format="float2",
            y_format=main_format,
            base_x=base_ngl_yield,
            base_y=main_base,
        )

        moic_ngl_yield_bid_heatmap = build_heatmap(
            moic_ngl_yield_bid_df,
            "MOIC Sensitivity",
            metric="moic",
            x_title="NGL Yield (GPM)",
            y_title=main_title,
            x_format="float2",
            y_format=main_format,
            base_x=base_ngl_yield,
            base_y=main_base,
        )

        irr_tcrisk_bid_df, moic_tcrisk_bid_df = run_two_way_sensitivity(
            slot_df=slot_df,
            deal_inputs=deal_inputs,
            x_values=cross_x_values,
            x_variable=cross_x_variable,
            y_values=main_values,
            y_variable=main_variable,
        )

        irr_tcrisk_bid_heatmap = build_heatmap(
            irr_tcrisk_bid_df,
            "IRR Sensitivity",
            metric="irr",
            x_title=cross_x_title,
            y_title=main_title,
            x_format=cross_x_format,
            y_format=main_format,
            base_x=cross_base_x,
            base_y=main_base,
        )

        moic_tcrisk_bid_heatmap = build_heatmap(
            moic_tcrisk_bid_df,
            "MOIC Sensitivity",
            metric="moic",
            x_title=cross_x_title,
            y_title=main_title,
            x_format=cross_x_format,
            y_format=main_format,
            base_x=cross_base_x,
            base_y=main_base,
        )

        irr_spud_tcrisk_df, moic_spud_tcrisk_df = run_two_way_sensitivity(
            slot_df=slot_df,
            deal_inputs=deal_inputs,
            x_values=spud_date_values,
            x_variable="spud_date",
            y_values=tc_risk_values,
            y_variable="tc_risk",
        )

        irr_spud_tcrisk_heatmap = build_heatmap(
            irr_spud_tcrisk_df,
            "IRR Sensitivity",
            metric="irr",
            x_title="Spud Date",
            y_title="TC Risk",
            x_format="date",
            y_format="percent",
            base_x=base_spud_month,
            base_y=base_tc_risk,
        )

        moic_spud_tcrisk_heatmap = build_heatmap(
            moic_spud_tcrisk_df,
            "MOIC Sensitivity",
            metric="moic",
            x_title="Spud Date",
            y_title="TC Risk",
            x_format="date",
            y_format="percent",
            base_x=base_spud_month,
            base_y=base_tc_risk,
        )

        with st.expander(
            f"D&C Costs ($/ft) vs. {main_title} Sensitivity",
            expanded=True,
        ):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("### IRR Sensitivity")
                st.plotly_chart(irr_heatmap, use_container_width=True)
            with col2:
                st.markdown("### MOIC Sensitivity")
                st.plotly_chart(moic_heatmap, use_container_width=True)

        with st.expander(
            f"Oil Price vs. {main_title} Sensitivity",
            expanded=False,
        ):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("### IRR Sensitivity")
                st.plotly_chart(irr_oil_bid_heatmap, use_container_width=True)
            with col2:
                st.markdown("### MOIC Sensitivity")
                st.plotly_chart(moic_oil_bid_heatmap, use_container_width=True)

        with st.expander(
            f"NGL Yield (GPM) vs. {main_title} Sensitivity",
            expanded=False,
        ):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("### IRR Sensitivity")
                st.plotly_chart(irr_ngl_yield_bid_heatmap, use_container_width=True)
            with col2:
                st.markdown("### MOIC Sensitivity")
                st.plotly_chart(moic_ngl_yield_bid_heatmap, use_container_width=True)

        with st.expander(
            f"Gas Price vs. {main_title} Sensitivity",
            expanded=False,
        ):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("### IRR Sensitivity")
                st.plotly_chart(irr_gas_bid_heatmap, use_container_width=True)
            with col2:
                st.markdown("### MOIC Sensitivity")
                st.plotly_chart(moic_gas_bid_heatmap, use_container_width=True)

        with st.expander(
            f"{cross_x_title} vs. {main_title} Sensitivity",
            expanded=False,
        ):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("### IRR Sensitivity")
                st.plotly_chart(irr_tcrisk_bid_heatmap, use_container_width=True)
            with col2:
                st.markdown("### MOIC Sensitivity")
                st.plotly_chart(moic_tcrisk_bid_heatmap, use_container_width=True)

        with st.expander(
            "Spud Date vs. TC Risk Sensitivity",
            expanded=False,
        ):
            st.caption(
                "Spud timing shifts in 3-month increments from 12 months earlier "
                "to 12 months later. For multi-slot deals, every slot shifts by "
                "the same amount so relative timing is preserved."
            )
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("### IRR Sensitivity")
                st.plotly_chart(
                    irr_spud_tcrisk_heatmap,
                    use_container_width=True,
                )
            with col2:
                st.markdown("### MOIC Sensitivity")
                st.plotly_chart(
                    moic_spud_tcrisk_heatmap,
                    use_container_width=True,
                )

    st.subheader("Outputs")
    
    slot_returns = run_individual_slot_returns(
        slot_df=slot_df,
        deal_inputs=deal_inputs,
    )
    
    tc_output_display_df, tc_output_row_styles = build_tc_assumptions_output_display_table(
        slot_df=slot_df,
        deal_inputs=deal_inputs,
        slot_returns=slot_returns,
    )
    tc_output_styler = style_tc_assumptions_output_table(
        tc_output_display_df,
        tc_output_row_styles,
    )
    
    quarterly_output_df = build_quarterly_output_table(
        deal_df=deal_df,
        all_slots_df=all_slots_df,
        slot_df=slot_df,
        deal_inputs=deal_inputs,
    )
    
    quarterly_output_display_df, quarterly_row_styles = build_quarterly_output_display_table(quarterly_output_df)
    quarterly_output_styler = style_quarterly_output_table(
        quarterly_output_display_df,
        quarterly_row_styles,
    )
    
    with st.expander("TC Assumptions Output", expanded=False):
        st.markdown(tc_output_styler.to_html(), unsafe_allow_html=True)
    
    with st.expander("Quarterly Output", expanded=False):
        st.markdown(quarterly_output_styler.to_html(), unsafe_allow_html=True)
    
        st.markdown("### Deal Highlights")
        h1, h2, h3, h4 = st.columns(4)
    
        with h1:
            render_deal_highlight_box(
                "IRR",
                format_accounting_percent(irr, decimals=1, zero_as_dash=False) if irr is not None else "N/A",
            )
        with h2:
            render_deal_highlight_box(
                "MOIC",
                format_accounting_number(moic, decimals=2, suffix="x", zero_as_dash=False) if moic is not None else "N/A",
            )
        with h3:
            render_deal_highlight_box(
                "Net Acres",
                format_accounting_number(total_net_acres, decimals=1),
            )
        with h4:
            render_deal_highlight_box(
                "$/Acre Bid",
                format_accounting_number(blended_bid, decimals=0, prefix="$"),
            )

    if not disable_heavy_outputs:
            
        cum_fcf_chart = build_cumulative_fcf_chart(deal_df, slot_df)
    
        # Use a simpler copy for static PNG / email export.
        # The normal chart above still keeps the shaded spud bands,
        # Gross Wells labels, and payback annotation in the app.
        cum_fcf_chart_for_email = go.Figure(cum_fcf_chart)
    
        cum_fcf_chart_for_email.update_layout(
            shapes=[],
            annotations=[],
            margin=dict(l=55, r=35, t=65, b=55),
        )
    
        cum_fcf_chart_for_email.add_hline(
            y=0,
            line_width=1,
            line_dash="dash",
            line_color="gray",
        )
    
        with st.expander("Charts", expanded=False):
            chart_tab1, chart_tab2, chart_tab3 = st.tabs(
                ["Cumulative FCF", "Production", "Scenario Matrix"]
            )
        
            with chart_tab1:
                st.plotly_chart(cum_fcf_chart, use_container_width=True)
        
            with chart_tab2:
                prod_chart_view = st.radio(
                    "Production Chart View",
                    ["Stacked Mcfe/d", "Stream Split"],
                    horizontal=True,
                    key="prod_chart_view",
                )
                prod_chart = build_production_profile_chart(
                    deal_df,
                    chart_view=prod_chart_view,
                )
                st.plotly_chart(prod_chart, use_container_width=True)
        
            with chart_tab3:
                st.plotly_chart(scenario_scatter_chart, use_container_width=True)
        
        st.subheader("Email Draft Export")
        
        # Clear the old default text once, if it is still saved in session state.
        if st.session_state.get("email_opportunity_name") == "Fill Name Here":
            st.session_state["email_opportunity_name"] = ""
        
        opportunity_name = st.text_input(
            "Opportunity Name for Email Draft / Deal Log",
            placeholder="Fill Name Here",
            key="email_opportunity_name",
        )
        
        opportunity_name_clean = opportunity_name.strip()
        exports_enabled = bool(opportunity_name_clean)
        
        if not exports_enabled:
            st.caption(
                "Enter an opportunity name to enable the email and deal-log downloads."
            )
        
        prod_chart_stacked = build_production_profile_chart(
            deal_df,
            chart_view="Stacked Mcfe/d",
        )
        
        email_html = build_email_html(
            opportunity_name=opportunity_name_clean,
            deal_inputs=deal_inputs,
            slot_df=slot_df,
            irr=irr,
            moic=moic,
            tc_output_styler=tc_output_styler,
            quarterly_output_styler=quarterly_output_styler,
            irr_oil_bid_heatmap=irr_oil_bid_heatmap,
            irr_gas_bid_heatmap=irr_gas_bid_heatmap,
            irr_heatmap=irr_heatmap,
            irr_tcrisk_bid_heatmap=irr_tcrisk_bid_heatmap,
            irr_ngl_yield_bid_heatmap=irr_ngl_yield_bid_heatmap,
            irr_spud_tcrisk_heatmap=irr_spud_tcrisk_heatmap,
            cum_fcf_chart=cum_fcf_chart_for_email,
            prod_chart_stacked=prod_chart_stacked,
            scenario_scatter_chart=scenario_scatter_chart,
        )
        
        with st.expander("Preview Email Draft", expanded=False):
            st.components.v1.html(email_html, height=900, scrolling=True)
        
        st.download_button(
            label="Download Email Draft (HTML)",
            data=email_html,
            file_name="utica_email_draft.html",
            mime="text/html",
            key="download_email_html",
            disabled=not exports_enabled,
        )
        
        if (
            "deal_log_overview_df" in st.session_state
            and "deal_log_slot_df" in st.session_state
        ):
            deal_log_csv = build_deal_log_csv(
                opportunity_name=opportunity_name_clean,
                overview_df=st.session_state["deal_log_overview_df"],
                model_slot_df=st.session_state["deal_log_slot_df"],
            )
        
            st.download_button(
                label="Download Deal Log CSV",
                data=deal_log_csv,
                file_name=st.session_state["deal_log_filename"],
                mime="text/csv",
                disabled=not exports_enabled,
            )

    else:
        st.info(
            "Sensitivities, charts, production graphs, and email export are disabled. "
            "TC Assumptions Output and Quarterly Output are still available above."
        )

else:
    st.info("Set your deal assumptions and slot inputs, then click Run Model.")
