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
#   0) NEW — "Probiotic Timeline" section: pick a Customer + Farm and see
#      that farm's Probiotic purchases laid out on a timeline (Before
#      Stocking / First 30 Days / After 30 Days / After 60 Days), plus the
#      farm's All PL Stocking Density for V (Vannamei), M (Monodon) and
#      Total.
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

#st.markdown("<h3 style='text-align:center;'>Hi Welcome, Probiotic Sales Dashboard</h3>", unsafe_allow_html=True)

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
# NEW SECTION — PROBIOTIC TIMELINE (per Customer + Farm)
#
# Pick a Customer + Farm and this shows that farm's Probiotic purchases
# (Sales Details rows whose "Item No." starts with "PRO", matched on the
# farm's Customer Code) laid out on a timeline:
#
#   Before Stocking   -> purchase dates BEFORE the Cycle Started Date
#   Cycle Started Date (box)
#   First 30 Days     -> Cycle Started Date  ..  +29 days
#   After 30 Days     -> +30 days            ..  +59 days
#   After 60 Days     -> +60 days            ..  onwards
#
# Cycle Started Date = the OLDEST pond start date on that farm. A pond's
# start date = its latest saved record's Date minus its DOC (identical to
# the Marketing Manager view's "Started on" date, i.e. today - DOC Today).
# Ponds whose Cycle Type is "Soon to be" haven't started, so are skipped.
#
# Each purchase date is one block listing only Probiotic Item Description
# + Quantity. This section ignores the date-range picker below — it always
# shows the farm's full probiotic history. Nothing is written anywhere.
#
# The All PL Stocking Density figures reuse the Marketing Manager view's
# feed-limit density logic: each pond's latest saved Density, summed per
# Species Culture, INCLUDING Full H ponds (V = Vannamei, M = Monodon).
# =========================================================================
def _pt_esc(v):
    return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def _pt_blocks_html(date_items, show_total=False):
    """date_items: list of (Timestamp, [(item_description, qty), ...]).
    show_total=True adds a brace + a "Total" box to the right listing each
    item's summed Quantity across every date in this section."""
    if not date_items:
        return "<div style='color:#888;font-size:0.85rem;padding:14px 16px;'>No probiotic purchases</div>"
    _html = ""
    for _d, _items in date_items:
        _rows = "".join(
            "<tr>"
            f"<td style='border:1px solid #888;padding:3px 12px;text-align:center;'>{_pt_esc(_i)}</td>"
            f"<td style='border:1px solid #888;padding:3px 12px;text-align:right;min-width:60px;'>{_q:,.0f}</td>"
            "</tr>"
            for _i, _q in _items
        )
        _html += (
            "<div style='display:flex;align-items:center;margin:16px 0;'>"
            "<div style='width:130px;text-align:center;font-size:0.9rem;'>"
            f"{_d.strftime('%Y-%m-%d')}<br><span style='font-size:1.3rem;line-height:1;'>⟶</span></div>"
            f"<table style='border-collapse:collapse;font-size:0.9rem;'>{_rows}</table>"
            "</div>"
        )
    if not show_total:
        return _html

    # Item totals across the whole section (in the order items first appear).
    _totals = {}
    for _d, _items in date_items:
        for _i, _q in _items:
            _totals[_i] = _totals.get(_i, 0) + _q
    _total_rows = "".join(
        "<tr>"
        f"<td style='border:1px solid #888;padding:3px 12px;text-align:center;'>{_pt_esc(_i)}</td>"
        f"<td style='border:1px solid #888;padding:3px 12px;text-align:right;min-width:60px;font-weight:bold;'>{_q:,.0f}</td>"
        "</tr>"
        for _i, _q in _totals.items()
    )
    _brace = (
        "<div style='align-self:stretch;width:14px;margin:16px 14px 16px 24px;border:2px solid #888;"
        "border-left:none;border-radius:0 14px 14px 0;'></div>"
    )
    _total_box = (
        "<div style='display:flex;align-items:center;'>"
        "<table style='border-collapse:collapse;font-size:0.9rem;'>"
        "<tr><td colspan='2' style='border:1px solid #888;padding:3px 12px;text-align:center;"
        "font-weight:bold;'>Total</td></tr>"
        f"{_total_rows}</table></div>"
    )
    return f"<div style='display:flex;align-items:stretch;'><div>{_html}</div>{_brace}{_total_box}</div>"

def _pt_section_html(label, blocks_html, milestone_title=None, milestone_date=None):
    _milestone = ""
    if milestone_title:
        _milestone = (
            "<div style='display:inline-block;border:2px solid #888;padding:6px 14px;font-weight:bold;"
            f"margin:8px 0 0 0;'>{_pt_esc(milestone_title)}"
            f"<div style='font-size:0.75rem;font-weight:normal;color:#888;'>{_pt_esc(milestone_date)}</div></div>"
        )
    return (
        "<div style='display:flex;align-items:stretch;'>"
        "<div style='width:120px;display:flex;align-items:center;justify-content:center;text-align:center;"
        f"font-weight:bold;border-right:3px solid #888;padding-right:8px;margin-right:0;'>{_pt_esc(label)}</div>"
        "<div style='flex:1;border-left:2px solid #888;margin-left:40px;padding-left:0;'>"
        f"<div style='margin-left:-2px;'>{_milestone}</div>"
        f"<div style='padding-left:20px;'>{blocks_html}</div>"
        "</div></div>"
    )

st.subheader("🧪 Probiotic Timeline — Customer & Farm")

_pt_customers = sorted(customer_df["Customer Name"].replace("", pd.NA).dropna().unique().tolist())
_pt_c1, _pt_c2 = st.columns(2)
with _pt_c1:
    _pt_customer = st.selectbox("Customer Name", _pt_customers, key="pt_customer_select")

_pt_farm_options = sorted(
    customer_df.loc[customer_df["Customer Name"] == _pt_customer, "Farm Name with Code"]
    .replace("", pd.NA).dropna().unique().tolist()
)
if not _pt_farm_options:
    _pt_farm_options = ["-- No farms found for this customer --"]
with _pt_c2:
    _pt_farm = st.selectbox("Farm Name with Code", _pt_farm_options, key=f"pt_farm_select_{_pt_customer}")

try:
    _pt_all = load_data()
except Exception as e:
    _pt_all = pd.DataFrame(columns=COLUMN_ORDER)
    st.error(f"❌ Could not connect to the main Google Sheet. Check sharing settings.\n\n{e}")

_pt_needed = {"Customer", "Farm Name with Code", "Pond Number", "Date", "DOC", "Density",
              "Species Culture", "Cycle Type"}
_pt_latest = pd.DataFrame()
if len(_pt_all) > 0 and _pt_needed.issubset(_pt_all.columns):
    _pt_farm_df = _pt_all[
        (_pt_all["Customer"] == _pt_customer) & (_pt_all["Farm Name with Code"] == _pt_farm)
    ].copy()
    if len(_pt_farm_df) > 0:
        _pt_farm_df["_ParsedDate"] = pd.to_datetime(_pt_farm_df["Date"], errors="coerce")
        # Latest saved record per pond (same basis as the Marketing Manager view).
        _pt_latest = (
            _pt_farm_df.dropna(subset=["_ParsedDate"])
            .sort_values("_ParsedDate")
            .groupby("Pond Number", as_index=False)
            .last()
        )

# ---- All PL Stocking Density: V / M / Total (Full H ponds INCLUDED) ----
_pt_density_v = 0.0
_pt_density_m = 0.0
_pt_density_total = 0.0
if len(_pt_latest) > 0:
    _pt_dens = pd.to_numeric(_pt_latest["Density"], errors="coerce")
    _pt_species = _pt_latest["Species Culture"].astype(str).str.strip().str.lower()
    _pt_density_v = float(_pt_dens[_pt_species.str.contains("vannamei")].sum())
    _pt_density_m = float(_pt_dens[_pt_species.str.contains("monodon")].sum())
    _pt_density_total = float(_pt_dens.sum())

st.markdown("**🦐 All PL Stocking Density**")
_pt_m1, _pt_m2, _pt_m3 = st.columns(3)
_pt_m1.metric("V (Vannamei)", f"{_pt_density_v:,.0f}")
_pt_m2.metric("M (Monodon)", f"{_pt_density_m:,.0f}")
_pt_m3.metric("Total", f"{_pt_density_total:,.0f}")
st.caption("Sum of each pond's latest saved Density (Full H ponds included), split by Species Culture.")

# ---- Cycle Started Date = OLDEST pond start date ----
def _pt_pond_start(prow):
    if str(prow.get("Cycle Type", "")).strip() == "Soon to be":
        return pd.NaT
    try:
        _doc = int(float(prow.get("DOC")))
    except (TypeError, ValueError):
        return pd.NaT
    return prow["_ParsedDate"] - pd.Timedelta(days=_doc)

_pt_cycle_start = pd.NaT
_pt_started_ponds = 0
if len(_pt_latest) > 0:
    _pt_starts = _pt_latest.apply(_pt_pond_start, axis=1).dropna()
    _pt_started_ponds = len(_pt_starts)
    if _pt_started_ponds > 0:
        _pt_cycle_start = _pt_starts.min().normalize()

_pt_code = _customer_code_for(_pt_customer, _pt_farm)

if not _pt_code:
    st.info("No Customer Code found for this farm in 'Customer List.xlsx', so probiotic purchases "
            "can't be matched.")
elif pd.isna(_pt_cycle_start):
    st.info("No started pond (valid DOC + Date) found for this farm yet, so the Cycle Started Date "
            "can't be worked out.")
else:
    _pt_sales = df_sales[
        (df_sales["Customer Code"].astype(str).str.strip().str.lower() == _pt_code.strip().lower())
        & df_sales["Item No."].astype(str).str.strip().str.upper().str.startswith(PROBIOTIC_PREFIX)
        & df_sales["_ParsedDate"].notna()
    ].copy()
    _pt_sales["_Day"] = _pt_sales["_ParsedDate"].dt.normalize()
    _pt_sales["_Item"] = _pt_sales["Item Description"].astype(str).str.strip()

    _pt_by_day = {}
    for (_day, _item), _qty in _pt_sales.groupby(["_Day", "_Item"], sort=False)["Quantity"].sum().items():
        _pt_by_day.setdefault(_day, []).append((_item, _qty))

    _pt_d30 = _pt_cycle_start + pd.Timedelta(days=30)
    _pt_d60 = _pt_cycle_start + pd.Timedelta(days=60)

    _pt_before, _pt_first30, _pt_after30, _pt_after60 = [], [], [], []
    for _day in sorted(_pt_by_day):
        _entry = (_day, _pt_by_day[_day])
        if _day < _pt_cycle_start:
            _pt_before.append(_entry)
        elif _day < _pt_d30:
            _pt_first30.append(_entry)
        elif _day < _pt_d60:
            _pt_after30.append(_entry)
        else:
            _pt_after60.append(_entry)

    _pt_fmt = lambda d: d.strftime("%Y-%m-%d")
    st.markdown(
        _pt_section_html("Before Stocking", _pt_blocks_html(_pt_before))
        + _pt_section_html("First 30 Days", _pt_blocks_html(_pt_first30, show_total=True),
                           "Cycle Started Date", _pt_fmt(_pt_cycle_start))
        + _pt_section_html("After 30 Days", _pt_blocks_html(_pt_after30, show_total=True),
                           "After 30 Days Date", _pt_fmt(_pt_d30))
        + _pt_section_html("After 60 Days", _pt_blocks_html(_pt_after60, show_total=True),
                           "After 60 Days Date", _pt_fmt(_pt_d60)),
        unsafe_allow_html=True,
    )
    st.caption(
        f"Cycle Started Date = oldest pond start date ({_pt_started_ponds} started pond(s); "
        "start = latest record Date − DOC). Probiotic items only (Item No. starting with "
        f"'{PROBIOTIC_PREFIX}'), Customer Code '{_pt_code}'."
    )

    # ---- NEW SECTION: POND-WISE PROBIOTIC TIMELINE -------------------------
    # Same selected Customer + Farm as the timeline above. One vertical "time
    # frame" line (Started Date / After 30 days / After 60 Days / Today), one
    # vertical line per started pond (its own Started / After 30 / After 60
    # dates, ending in a circle showing that pond's DOC today — or its DOC at
    # Full Harvest, same rule as the Marketing Manager view). Positions are
    # proportional to the dates. Probiotic purchases are recorded per customer
    # (not per pond), so they are drawn once, as blocks hanging off the time
    # frame line. View-only; nothing is written anywhere.
    import streamlit.components.v1 as _pt_components

    st.markdown("---")
    st.subheader("🗓️ Pond-wise Probiotic Timeline")

    _ptw_today = pd.Timestamp(date.today())
    _ptw_ponds = []
    for _, _ptw_pr in _pt_latest.iterrows():
        _ptw_start = _pt_pond_start(_ptw_pr)
        if pd.isna(_ptw_start):
            continue
        _ptw_start = _ptw_start.normalize()
        _ptw_end, _ptw_full = _ptw_today, False
        _ptw_t2 = str(_ptw_pr.get("Harvest Type 2", "")).strip().lower()
        _ptw_t1 = str(_ptw_pr.get("Harvest Type", "")).strip().lower()
        _ptw_fd_str = ""
        if "full" in _ptw_t2:
            _ptw_fd_str = str(_ptw_pr.get("Harvest Date 2", "")).strip()
        elif "full" in _ptw_t1:
            _ptw_fd_str = str(_ptw_pr.get("Harvest Date", "")).strip()
        if _ptw_fd_str:
            _ptw_fd = pd.to_datetime(_ptw_fd_str, errors="coerce")
            if pd.notna(_ptw_fd):
                _ptw_end, _ptw_full = _ptw_fd.normalize(), True
        _ptw_ponds.append({
            "name": str(_ptw_pr.get("Pond Number", "")),
            "start": _ptw_start,
            "end": _ptw_end,
            "full": _ptw_full,
            "doc": (_ptw_end - _ptw_start).days,   # = DOC Today (or DOC at Full H)
        })

    _ptw_all_days = [p["start"] for p in _ptw_ponds] + list(_pt_by_day.keys())
    _ptw_tmin = min(_ptw_all_days)
    _ptw_tmax = max([_ptw_today] + [p["end"] for p in _ptw_ponds] + list(_pt_by_day.keys()))
    _PPD, _Y0 = 8, 80                      # pixels per day, top offset
    _AX, _BX = 170, 200                    # time-frame line x, probiotic blocks x
    _POND_X0, _POND_DX = 560, 190
    _ptw_y = lambda d: _Y0 + (d - _ptw_tmin).days * _PPD

    _ptw_svg = []
    def _ptw_text(x, y, txt, anchor="start", weight="normal", size=12, fill="#222", extra=""):
        _ptw_svg.append(
            f"<text x='{x}' y='{y}' text-anchor='{anchor}' font-weight='{weight}' font-size='{size}' "
            f"fill='{fill}' {extra}>{_pt_esc(txt)}</text>"
        )

    # -- probiotic purchase blocks (collision-avoiding, connected to the axis)
    _ptw_prev_bottom = _Y0 - 30
    for _ptw_day in sorted(_pt_by_day):
        _ptw_items = _pt_by_day[_ptw_day]
        _ptw_yp = _ptw_y(_ptw_day)
        _ptw_bh = 20 * (len(_ptw_items) + 1)
        _ptw_top = max(_ptw_yp - 10, _ptw_prev_bottom + 12)
        _ptw_prev_bottom = _ptw_top + _ptw_bh
        _ptw_svg.append(f"<line x1='{_AX}' y1='{_ptw_yp}' x2='{_BX}' y2='{_ptw_top + 10}' stroke='#aaa'/>")
        _ptw_svg.append(f"<circle cx='{_AX}' cy='{_ptw_yp}' r='3' fill='#555'/>")
        _ptw_svg.append(f"<rect x='{_BX}' y='{_ptw_top}' width='250' height='20' fill='#e8eefb' stroke='#888'/>")
        _ptw_text(_BX + 125, _ptw_top + 14, _ptw_day.strftime("%Y-%m-%d"), "middle", "bold")
        for _k, (_ptw_item, _ptw_q) in enumerate(_ptw_items):
            _ry = _ptw_top + 20 * (_k + 1)
            _ptw_svg.append(f"<rect x='{_BX}' y='{_ry}' width='200' height='20' fill='#fff' stroke='#888'/>")
            _ptw_svg.append(f"<rect x='{_BX + 200}' y='{_ry}' width='50' height='20' fill='#fff' stroke='#888'/>")
            _ptw_text(_BX + 100, _ry + 14, _ptw_item, "middle")
            _ptw_text(_BX + 244, _ry + 14, f"{_ptw_q:,.0f}", "end")

    # -- time frame line (left): Started / After 30 / After 60 / Today
    _ptw_ax_bottom = _ptw_y(_ptw_tmax) + 20
    _ptw_svg.append(f"<line x1='{_AX}' y1='{_Y0 - 20}' x2='{_AX}' y2='{_ptw_ax_bottom}' stroke='#222' stroke-width='1.5'/>")
    _ptw_marks = [(_pt_cycle_start, "Started Date"),
                  (_pt_cycle_start + pd.Timedelta(days=30), "After 30 days Date"),
                  (_pt_cycle_start + pd.Timedelta(days=60), "After 60 Days Date"),
                  (_ptw_today, "Today Date")]
    _ptw_prev_label_y = -999
    for _ptw_md, _ptw_ml in sorted([m for m in _ptw_marks if m[0] <= _ptw_tmax], key=lambda m: m[0]):
        _ptw_my = _ptw_y(_ptw_md)
        _ptw_ly = max(_ptw_my, _ptw_prev_label_y + 32)
        _ptw_prev_label_y = _ptw_ly
        _ptw_svg.append(f"<line x1='{_AX - 8}' y1='{_ptw_my}' x2='{_AX + 8}' y2='{_ptw_my}' stroke='#222'/>")
        _ptw_text(_AX - 14, _ptw_ly - 2, _ptw_ml, "end", "bold")
        _ptw_text(_AX - 14, _ptw_ly + 12, _ptw_md.strftime("%Y-%m-%d"), "end")
        if _ptw_ml == "Today Date":
            _ptw_svg.append(
                f"<line x1='{_AX}' y1='{_ptw_my}' x2='{_POND_X0 + _POND_DX * len(_ptw_ponds)}' y2='{_ptw_my}' "
                "stroke='#c33' stroke-dasharray='4 4'/>"
            )

    # -- "First 30 Days" bracket
    _ptw_b1 = _ptw_y(_pt_cycle_start)
    _ptw_b2 = _ptw_y(min(_pt_cycle_start + pd.Timedelta(days=30), _ptw_tmax))
    _ptw_svg.append(f"<path d='M62,{_ptw_b1} H50 V{_ptw_b2} H62' fill='none' stroke='#222'/>")
    _ptw_text(42, (_ptw_b1 + _ptw_b2) // 2, "First 30 Days", "middle", "normal", 12, "#222",
              f"transform='rotate(-90 42 {(_ptw_b1 + _ptw_b2) // 2})'")

    # -- one vertical line per pond
    for _i, _p in enumerate(_ptw_ponds):
        _px = _POND_X0 + _i * _POND_DX
        _y1, _y2 = _ptw_y(_p["start"]), _ptw_y(_p["end"])
        _ptw_text(_px, 40, f"Pond {_p['name']}", "middle", "bold", 14)
        _ptw_svg.append(f"<line x1='{_px}' y1='{_y1}' x2='{_px}' y2='{_y2}' stroke='#222' stroke-width='1.5'/>")
        for _ml, _md in [("Started Date", _p["start"]),
                         ("After 30 days Date", _p["start"] + pd.Timedelta(days=30)),
                         ("After 60 Days Date", _p["start"] + pd.Timedelta(days=60))]:
            if _md <= _p["end"]:
                _my = _ptw_y(_md)
                _ptw_svg.append(f"<line x1='{_px - 8}' y1='{_my}' x2='{_px + 8}' y2='{_my}' stroke='#222'/>")
                _ptw_text(_px + 12, _my - 2, _ml, "start", "bold", 11)
                _ptw_text(_px + 12, _my + 11, _md.strftime("%Y-%m-%d"), "start", "normal", 11)
        _ptw_svg.append(f"<circle cx='{_px}' cy='{_y2}' r='16' fill='#4472c4' stroke='#1f3864'/>")
        _ptw_text(_px, _y2 + 4, str(_p["doc"]), "middle", "bold", 12, "#ffd966")
        _ptw_text(_px, _y2 + 32, "Full H DOC" if _p["full"] else "Today DOC", "middle", "normal", 11)

    _ptw_width = _POND_X0 + _POND_DX * len(_ptw_ponds) + 40
    _ptw_height = int(max(_ptw_ax_bottom + 80, _ptw_prev_bottom + 40))
    _pt_components.html(
        "<div style='background:#fff;color:#222;border-radius:8px;padding:8px;overflow-x:auto;"
        "font-family:sans-serif;'>"
        f"<svg width='{_ptw_width}' height='{_ptw_height}' xmlns='http://www.w3.org/2000/svg' "
        f"font-family='sans-serif'>{''.join(_ptw_svg)}</svg></div>",
        height=_ptw_height + 30,
        scrolling=True,
    )
    st.caption(
        "Positions are proportional to the dates. Each pond line runs from its Started Date (latest record "
        "Date − DOC) to today, or to its Full H date; the circle shows its DOC today (DOC at Full H). "
        "Probiotic purchases are recorded per customer, so they appear once beside the time frame line."
    )

st.markdown("---")

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
