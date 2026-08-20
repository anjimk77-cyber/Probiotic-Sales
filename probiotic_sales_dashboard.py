import io
import pandas as pd
import streamlit as st
from datetime import date, timedelta

import gspread
from google.oauth2.service_account import Credentials

# =========================================================================
# CONFIG
#
# This is a SEPARATE, standalone Streamlit app: "Probiotic Sales Dashboard".
# It shares the same Google Sheets and "Customer List.xlsx" as the
# data-entry app (app.py) and the manager app, but is entirely VIEW +
# DOWNLOAD ONLY — it never writes anything back to either Google Sheet.
#
# What it shows:
#   1) A date-range picker (Sales Details "Date" column).
#   2) Zone-wise tables. Each table lists every currently RUNNING
#      Customer/Farm (same "Running" definition as the manager app's
#      Running List: at least one pond not yet Full Harvested), with:
#        Customer Name | Farm Name with Code | <one column per Probiotic
#        item> | Total
#      Each Probiotic column = summed Sales "Quantity" for that farm's
#      Customer Code, for sales rows whose "Item No." starts with "PRO"
#      and whose Date falls inside the selected range.
#   3) A CSV download button for the combined table.
#
# Deploy this as its own Streamlit app (its own URL/link), separate from
# app.py and the manager app.
# =========================================================================
st.set_page_config(page_title="Probiotic Sales Dashboard - KMN", layout="wide", page_icon="🧪")

CUSTOMER_FILE = "Customer List.xlsx"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Second Google Sheet — Sales Details. Same spreadsheet key used by the
# manager app; the same service account must be shared (Viewer is enough,
# since this app never writes) on this sheet.
SALES_SHEET_ID = "1S3csAE-E_hN8vstuHR0KkeAN7yCVQTFe4AkEVlw4vQw"

# Must match app.py's COLUMN_ORDER exactly, since this app reads the same
# main WaterQualityData sheet (only to work out which farms are currently
# "Running", i.e. not yet fully harvested).
COLUMN_ORDER = [
    "Timestamp", "Customer", "Farm Name with Code", "Zone", "Area",
    "Pond Number", "Date", "Species Culture", "Cycle Type",
    "DOC", "Density", "Feed Per Day", "ABW",
    "Expect Harvest (KG)", "Survival QTY",
    "Issues", "Water Color", "Grade", "Remark", "Technician",
    "Harvest Date", "Harvest Type", "Harvest KG", "Harvest ABW",
    "Harvest Date 2", "Harvest Type 2", "Harvest KG 2", "Harvest ABW 2",
    "Deleted",
]

# Expected columns in the Sales Details Google Sheet.
SALES_COLUMN_ORDER = [
    "Date", "Item No.", "Item Description", "Customer Code",
    "Customer Name", "Quantity", "Sales Amt", "Settle",
]

# Item No. prefix that identifies a Probiotic line item.
PROBIOTIC_PREFIX = "PRO"

# =========================================================================
# LOGIN GATE
# Standard username/password login screen with a blurred background image.
# Nothing below this block runs until the user is authenticated.
# =========================================================================
LOGIN_USERNAME = "Lakshani"
LOGIN_PASSWORD = "2000"
LOGIN_BACKGROUND_IMAGE_URL = (
    "https://images.unsplash.com/photo-1717737852821-1bea137cab50"
    "?auto=format&fit=crop&w=1740&q=80&blur=60"
)

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

def _render_login_page():
    st.markdown(
        f"""
        <style>
        [data-testid="stAppViewContainer"] {{
            background-image: linear-gradient(rgba(0, 0, 0, 0.65), rgba(0, 0, 0, 0.65)),
                url("{LOGIN_BACKGROUND_IMAGE_URL}");
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
            background-attachment: fixed;
        }}
        [data-testid="stAppViewContainer"] > .main {{
            background: transparent;
        }}
        [data-testid="stHeader"] {{
            background: rgba(0,0,0,0);
        }}
        div[data-testid="stForm"] {{
            background: rgba(255, 255, 255, 0.18);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
            border: 1px solid rgba(255, 255, 255, 0.35);
            border-radius: 18px;
            padding: 2.2rem 2rem 1.4rem 2rem;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.35);
        }}
        div[data-testid="stForm"] label p {{
            color: #ffffff !important;
            font-weight: 600;
        }}
        .login-title, .login-title * {{
            text-align: center;
            color: #ffffff !important;
            text-shadow: 0 2px 8px rgba(0,0,0,0.5);
            margin-bottom: 0.2rem;
        }}
        .login-subtitle, .login-subtitle * {{
            text-align: center;
            color: #ffffff !important;
            text-shadow: 0 1px 6px rgba(0,0,0,0.5);
            margin-bottom: 1.6rem;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    _c1, _c2, _c3 = st.columns([1, 1.1, 1])
    with _c2:
        st.markdown("<h1 class='login-title'>🧪 Probiotic Sales Dashboard</h1>", unsafe_allow_html=True)
        st.markdown("<p class='login-subtitle'>KMN Aqua Services — please sign in</p>", unsafe_allow_html=True)
        with st.form("login_form", clear_on_submit=False):
            _username = st.text_input("Username")
            _password = st.text_input("Password", type="password")
            _submitted = st.form_submit_button("🔐 Login", use_container_width=True)

        if _submitted:
            if _username == LOGIN_USERNAME and _password == LOGIN_PASSWORD:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("❌ Incorrect username or password.")

if not st.session_state["authenticated"]:
    _render_login_page()
    st.stop()

st.markdown("<h3 style='text-align:center;'>Hi Welcome, Probiotic Sales Dashboard</h3>", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center;'>🧪 Probiotic Sales Dashboard</h1>",
            unsafe_allow_html=True)
st.subheader("KMN Aqua Services")
st.caption("View & download only — this dashboard never writes back to any Google Sheet.")
st.markdown("---")

# =========================================================================
# GOOGLE SHEETS BACKEND — entirely read-only.
# =========================================================================
def _gsheet_configured():
    return "gcp_service_account" in st.secrets and "gsheet" in st.secrets and "sheet_id" in st.secrets["gsheet"]

@st.cache_resource(show_spinner=False)
def get_worksheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet_id = st.secrets["gsheet"]["sheet_id"]
    worksheet_name = st.secrets["gsheet"].get("worksheet_name", "WaterQualityData")
    sh = client.open_by_key(sheet_id)
    return sh.worksheet(worksheet_name)

@st.cache_resource(show_spinner=False)
def get_sales_worksheet():
    """Separate spreadsheet (Sales Details) — same service account creds,
    different spreadsheet key. Worksheet/tab name can be overridden via
    st.secrets["gsheet"]["sales_worksheet_name"] (defaults to the first
    sheet/tab in the spreadsheet if not set)."""
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SALES_SHEET_ID)
    worksheet_name = st.secrets.get("gsheet", {}).get("sales_worksheet_name", "")
    if worksheet_name:
        return sh.worksheet(worksheet_name)
    return sh.sheet1

def load_data():
    """Reads the main WaterQualityData sheet fresh (no caching) — used only
    to work out which farms are currently 'Running'. Mirrors the manager
    app's load_data() (Deleted / Harvest Status filtering)."""
    ws = get_worksheet()
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for c in COLUMN_ORDER:
        if c not in df.columns:
            df[c] = ""
    if "Harvest Status" not in df.columns:
        df["Harvest Status"] = ""
    if len(df) > 0:
        df = df[COLUMN_ORDER + ["Harvest Status"]]
    df = df.astype(str).replace("nan", "")
    if "Deleted" in df.columns:
        is_deleted = df["Deleted"].astype(str).str.strip().str.lower().isin(["yes", "true", "1"])
        df = df[~is_deleted].reset_index(drop=True)
    if "Harvest Status" in df.columns:
        is_harvest_hidden = df["Harvest Status"].astype(str).str.strip().str.upper() == "H"
        df = df[~is_harvest_hidden].reset_index(drop=True)
    return df

def load_sales_data():
    """Reads the Sales Details sheet fresh (no caching)."""
    ws = get_sales_worksheet()
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for c in SALES_COLUMN_ORDER:
        if c not in df.columns:
            df[c] = ""
    return df

if not _gsheet_configured():
    st.error("❌ Google Sheets is not configured yet. This app needs the same "
             "`.streamlit/secrets.toml` (the `[gcp_service_account]` and `[gsheet]` sections) "
             "used by the data-entry app.")
    st.stop()

try:
    get_worksheet()
except Exception as e:
    st.error(f"❌ Could not connect to the Google Sheet. Check your secrets and sharing settings.\n\n{e}")
    st.stop()

# =========================================================================
# LOAD CUSTOMER LIST (for Zone + Customer Code lookups)
# =========================================================================
@st.cache_data
def load_customer_data():
    return pd.read_excel(CUSTOMER_FILE)

try:
    customer_df = load_customer_data()
except Exception as e:
    st.error(f"❌ Could not load '{CUSTOMER_FILE}'. Make sure it's in the app folder. ({e})")
    st.stop()

REQUIRED_COLS = ["Customer Name", "Farm Name with Code", "Zone"]
missing_cols = [c for c in REQUIRED_COLS if c not in customer_df.columns]
if missing_cols:
    st.error(f"❌ 'Customer List.xlsx' is missing required column(s): {', '.join(missing_cols)}")
    st.stop()

for _col in REQUIRED_COLS:
    customer_df[_col] = customer_df[_col].apply(
        lambda v: "" if pd.isna(v) else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))
    )

# Customer Code may live under any of a few likely column names — the
# first one that exists and has a non-blank value for a given farm wins.
_CUSTOMER_CODE_COLUMN_CANDIDATES = [
    "Customer Code", "Customer ID", "Customer Code with Code", "Code", "Cust Code",
]

def _customer_code_for(cust_name, farm_name):
    _match = customer_df[
        (customer_df["Customer Name"] == cust_name) & (customer_df["Farm Name with Code"] == farm_name)
    ]
    if len(_match) == 0:
        return ""
    for _cand in _CUSTOMER_CODE_COLUMN_CANDIDATES:
        if _cand in customer_df.columns:
            _val = str(_match.iloc[0].get(_cand, "")).strip()
            if _val and _val.lower() != "nan":
                return _val
    return ""

# =========================================================================
# LOAD SALES DATA (loaded early — before the date picker — so the date
# range selector can be bounded to the actual dates present in the Sales
# Details sheet's "Date" column, instead of defaulting to today's date).
# =========================================================================
try:
    df_sales = load_sales_data()
except Exception as e:
    st.error(f"❌ Could not connect to the Sales Details Google Sheet. Check sharing settings.\n\n{e}")
    st.stop()

if len(df_sales) == 0:
    st.info("No sales records found in the Sales Details sheet.")
    st.stop()

df_sales = df_sales.copy()
df_sales["Quantity"] = pd.to_numeric(df_sales["Quantity"], errors="coerce").fillna(0)
df_sales["_ParsedDate"] = pd.to_datetime(df_sales["Date"], errors="coerce")

_valid_sales_dates = df_sales["_ParsedDate"].dropna()
if len(_valid_sales_dates) == 0:
    st.error("❌ No valid dates found in the Sales Details sheet's 'Date' column.")
    st.stop()

_sales_min_date = _valid_sales_dates.min().date()
_sales_max_date = _valid_sales_dates.max().date()

# =========================================================================
# DATE RANGE SELECTOR
# Bounded (min_value/max_value) to the earliest and latest dates actually
# present in the Sales Details "Date" column, so the user can't pick a
# range with no possible data.
# =========================================================================
st.subheader("📅 Select Date Range")

col_d1, col_d2 = st.columns(2)
with col_d1:
    start_date = st.date_input(
        "From",
        value=_sales_min_date,
        min_value=_sales_min_date,
        max_value=_sales_max_date,
        key="probiotic_start_date",
    )
with col_d2:
    end_date = st.date_input(
        "To",
        value=_sales_max_date,
        min_value=_sales_min_date,
        max_value=_sales_max_date,
        key="probiotic_end_date",
    )

if start_date > end_date:
    st.error("❌ 'From' date must be on or before 'To' date.")
    st.stop()

if st.button("🔄 Refresh"):
    st.rerun()

st.markdown("---")

# =========================================================================
# WORK OUT WHICH FARMS ARE CURRENTLY "RUNNING"
# Running = at least one pond NOT YET Full Harvested. Same pond-status
# rules used by the manager app's Pond Layout / Running List sections
# (latest saved record per pond; a pond keeps counting as Partial H if
# ANY of its saved records ever had a Partial harvest).
# =========================================================================
try:
    df_all = load_data()
except Exception as e:
    st.error(f"❌ Could not connect to the main Google Sheet. Check sharing settings.\n\n{e}")
    st.stop()

_running_required = {"Customer", "Farm Name with Code", "Pond Number", "Date",
                      "Harvest Type", "Harvest Type 2"}
running_farms = pd.DataFrame(columns=["Customer", "Farm Name with Code"])

if len(df_all) > 0 and _running_required.issubset(df_all.columns):
    df_all = df_all.copy()
    df_all["_ParsedDate"] = pd.to_datetime(df_all["Date"], errors="coerce")

    _latest_per_pond = (
        df_all.dropna(subset=["_ParsedDate"])
        .sort_values("_ParsedDate")
        .groupby(["Customer", "Farm Name with Code", "Pond Number"], as_index=False)
        .last()
    )

    _partial_hist = (
        df_all.assign(
            _HasPartial=(
                df_all.get("Harvest Type", pd.Series("", index=df_all.index))
                .astype(str).str.lower().str.contains("partial")
                | df_all.get("Harvest Type 2", pd.Series("", index=df_all.index))
                .astype(str).str.lower().str.contains("partial")
            )
        )
        .groupby(["Customer", "Farm Name with Code", "Pond Number"])["_HasPartial"]
        .any()
    )

    def _pond_status(prow):
        _h_type = (str(prow.get("Harvest Type 2", "")).strip()
                   or str(prow.get("Harvest Type", "")).strip()).lower()
        _key = (prow.get("Customer", ""), prow.get("Farm Name with Code", ""), prow.get("Pond Number", ""))
        _has_partial = bool(_partial_hist.get(_key, False))
        if "full" in _h_type:
            return "Full H"
        elif "partial" in _h_type or _has_partial:
            return "Partial H"
        else:
            return "Running"

    _latest_per_pond["_PondStatus"] = _latest_per_pond.apply(_pond_status, axis=1)

    _farm_pond_summary = (
        _latest_per_pond.groupby(["Customer", "Farm Name with Code"])
        .agg(
            **{
                "No of Ponds": ("Pond Number", "nunique"),
                "Full Harvested Ponds": ("_PondStatus", lambda s: (s == "Full H").sum()),
            }
        )
        .reset_index()
    )

    running_farms = _farm_pond_summary[
        _farm_pond_summary["Full Harvested Ponds"] < _farm_pond_summary["No of Ponds"]
    ][["Customer", "Farm Name with Code"]].reset_index(drop=True)

if len(running_farms) == 0:
    st.info("No running farms found — every farm's ponds are fully harvested.")
    st.stop()

# Attach Zone from the customer list.
_zone_lookup = customer_df[["Customer Name", "Farm Name with Code", "Zone"]].drop_duplicates(
    subset=["Customer Name", "Farm Name with Code"]
).rename(columns={"Customer Name": "Customer"})
running_farms = running_farms.merge(_zone_lookup, on=["Customer", "Farm Name with Code"], how="left")
running_farms["Zone"] = running_farms["Zone"].fillna("")
running_farms["Customer Code"] = running_farms.apply(
    lambda r: _customer_code_for(r["Customer"], r["Farm Name with Code"]), axis=1
)

# =========================================================================
# FILTER SALES DATA TO PROBIOTIC ITEMS WITHIN THE SELECTED DATE RANGE
# (df_sales was already loaded above, before the date picker.)
# =========================================================================
_start_ts = pd.Timestamp(start_date)
_end_ts = pd.Timestamp(end_date)

df_probiotic = df_sales[
    df_sales["Item No."].astype(str).str.strip().str.upper().str.startswith(PROBIOTIC_PREFIX)
    & df_sales["_ParsedDate"].between(_start_ts, _end_ts)
].copy()

# All Probiotic item descriptions seen in the selected range, so every
# zone table shares the same set of item columns (in the order first seen).
_probiotic_items = list(dict.fromkeys(df_probiotic["Item Description"].astype(str).str.strip()))

if not _probiotic_items:
    st.info(
        f"No Probiotic sales (Item No. starting with '{PROBIOTIC_PREFIX}') found "
        f"between {start_date} and {end_date}."
    )
    st.stop()

# Sum of Quantity per Customer Code + Item Description, for the selected
# date range. NOTE: sales are recorded per Customer Code, not per farm —
# if the same Customer Code covers more than one farm, that customer's
# totals appear identically on each of that customer's running-farm rows
# (same convention already used by the manager app's feed-purchase columns).
_qty_by_code_item = (
    df_probiotic.groupby(["Customer Code", "Item Description"])["Quantity"]
    .sum()
)

def _qty_for(code, item):
    if not code:
        return 0
    return _qty_by_code_item.get((code, item), 0)

# =========================================================================
# BUILD THE ZONE-WISE TABLES
# =========================================================================
st.subheader(f"🌍 Probiotic Sales — Zone Wise ({start_date} to {end_date})")

_zones_present = sorted(
    [z for z in running_farms["Zone"].astype(str).str.strip().unique() if z and z.lower() != "nan"]
)

_display_cols = ["Customer Name", "Farm Name with Code"] + _probiotic_items + ["Total"]
_all_zone_tables = []  # list of (zone_name, DataFrame) — collected for the download below

def _build_zone_table(zone_farms_df):
    _rows = []
    for _, _f in zone_farms_df.iterrows():
        _code = _f["Customer Code"]
        _row = {
            "Customer Name": _f["Customer"],
            "Farm Name with Code": _f["Farm Name with Code"],
        }
        _row_total = 0
        for _item in _probiotic_items:
            _val = _qty_for(_code, _item)
            _row[_item] = _val
            _row_total += _val
        _row["Total"] = _row_total
        _rows.append(_row)
    return pd.DataFrame(_rows, columns=_display_cols)

if _zones_present:
    _selected_zones = st.multiselect(
        "Select Zone(s)", options=_zones_present, default=_zones_present, key="probiotic_zone_filter"
    )
    if not _selected_zones:
        st.info("Select at least one zone above to display probiotic sales.")
    else:
        for _zone in _selected_zones:
            _zone_farms = running_farms[running_farms["Zone"].astype(str).str.strip() == _zone]
            _zone_table = _build_zone_table(_zone_farms)
            st.markdown(f"**{_zone}** ({len(_zone_table)} running farm(s))")
            st.dataframe(_zone_table, use_container_width=True, hide_index=True)
            _all_zone_tables.append((_zone, _zone_table.copy()))
else:
    st.info("No Zone information found on the customer list — showing unfiltered.")
    _table = _build_zone_table(running_farms)
    st.dataframe(_table, use_container_width=True, hide_index=True)
    _all_zone_tables.append(("All Farms", _table.copy()))

# =========================================================================
# DOWNLOAD — one CSV covering every zone table shown above, laid out as:
#   Row 1: the selected date range
#   Then, for each zone in turn: the zone name, its column headers, and
#   its data rows — followed by a blank line before the next zone.
# =========================================================================
if _all_zone_tables:
    st.markdown("---")
    _csv_buffer = io.StringIO()
    _csv_buffer.write(f"Selected Date Range,{start_date} to {end_date}\n")
    _csv_buffer.write("\n")
    for _zone_name, _zone_df in _all_zone_tables:
        _csv_buffer.write(f"{_zone_name}\n")
        _zone_df.to_csv(_csv_buffer, index=False)
        _csv_buffer.write("\n")
    st.download_button(
        label="⬇️ Download as CSV",
        data=_csv_buffer.getvalue(),
        file_name=f"probiotic_sales_{start_date}_to_{end_date}.csv",
        mime="text/csv",
    )

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>KMN Aqua Services - Probiotic Sales Dashboard "
    "(View & Download only)</p>",
    unsafe_allow_html=True,
)
