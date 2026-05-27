import io
import re
from datetime import timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

APP_TITLE = "Suta Customer Analytics Dashboard"
DATA_DIR = Path("data")
COMBINED_FILE = DATA_DIR / "customer_data.xlsx"
WEBSITE_FILE = DATA_DIR / "website_data.xlsx"
STORE_FILE = DATA_DIR / "store_data.xlsx"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

CSS = """
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 1rem;}
.main-title {font-size: 2rem; font-weight: 850; color: #5f1230; margin-bottom: .1rem;}
.sub-title {font-size: .95rem; color: #666; margin-bottom: 1rem;}
.metric-card {background: linear-gradient(180deg,#fff,#fbf7f8); border:1px solid #ead9de; padding: 1rem; border-radius: 18px; box-shadow:0 3px 16px rgba(95,18,48,.07);} 
.metric-label {font-size:.82rem; color:#6b5b63; margin-bottom:.35rem;}
.metric-value {font-size:1.55rem; font-weight:850; color:#431023;}
.metric-help {font-size:.75rem; color:#8a7a82; margin-top:.25rem;}
.small-note {font-size:.85rem; color:#6b5b63;}
.dataframe {font-size: 0.85rem;}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def normalize_col_name(x: object) -> str:
    """Normalize headers so ' Sum of NET_AMOUNT ' and 'sum_of_net_amount' can be matched."""
    s = str(x).strip().upper()
    s = re.sub(r"[\.\-_]+", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s


def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    lookup = {normalize_col_name(c): c for c in df.columns}
    for cand in candidates:
        key = normalize_col_name(cand)
        if key in lookup:
            return lookup[key]
    # soft contains fallback
    for cand in candidates:
        key = normalize_col_name(cand)
        for normalized, original in lookup.items():
            if key == normalized or key in normalized or normalized in key:
                return original
    return None


def pick_series(df: pd.DataFrame, candidates: list[str], default=None):
    col = find_col(df, candidates)
    if col is None:
        return default
    return df[col]


def clean_mobile_value(x: object) -> str | None:
    if pd.isna(x):
        return None
    if isinstance(x, float):
        if pd.isna(x):
            return None
        x = str(int(x))
    else:
        x = str(x)
    digits = re.sub(r"\D", "", x)
    if len(digits) >= 10:
        return digits[-10:]
    return None


def money_fmt(x: float) -> str:
    try:
        return f"₹{x:,.0f}"
    except Exception:
        return "₹0"


def num_fmt(x: float) -> str:
    try:
        return f"{int(x):,}"
    except Exception:
        return "0"


def pct_fmt(x: float) -> str:
    try:
        return f"{x:.1f}%"
    except Exception:
        return "0.0%"


def metric_card(label: str, value: str, help_text: str = ""):
    st.markdown(
        f"""
        <div class="metric-card">
          <div class="metric-label">{label}</div>
          <div class="metric-value">{value}</div>
          <div class="metric-help">{help_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def frequency_bucket(bills: float) -> str:
    try:
        bills = int(bills)
    except Exception:
        bills = 0
    if bills == 1:
        return "1 Time"
    if bills == 2:
        return "2 Times"
    if bills == 3:
        return "3 Times"
    if 4 <= bills <= 5:
        return "4-5 Times"
    if 6 <= bills <= 10:
        return "6-10 Times"
    if bills >= 11:
        return "11+ Times"
    return "No Purchase"


@st.cache_data(show_spinner="Reading and preparing Excel data...")
def load_customer_master(combined_file: str, website_file: str, store_file: str) -> tuple[pd.DataFrame, dict]:
    combined_path = Path(combined_file)
    website_path = Path(website_file)
    store_path = Path(store_file)

    diagnostics = {
        "mode": None,
        "combined_file": str(combined_path),
        "website_file": str(website_path),
        "store_file": str(store_path),
        "website_columns": [],
        "store_columns": [],
        "rows_raw": {},
        "mapped_columns": {},
    }

    def read_preferred_sheet(path: Path, preferred_sheet: str) -> pd.DataFrame:
        xls = pd.ExcelFile(path, engine="openpyxl")
        sheet_lookup = {normalize_col_name(s): s for s in xls.sheet_names}
        sheet = sheet_lookup.get(normalize_col_name(preferred_sheet), xls.sheet_names[0])
        return pd.read_excel(path, sheet_name=sheet, engine="openpyxl")

    try:
        if website_path.exists() and store_path.exists():
            diagnostics["mode"] = "Separate files"
            website = read_preferred_sheet(website_path, "WEBSITE")
            store = read_preferred_sheet(store_path, "STORE")
        elif combined_path.exists():
            diagnostics["mode"] = "Combined workbook"
            xls = pd.ExcelFile(combined_path, engine="openpyxl")
            sheet_lookup = {normalize_col_name(s): s for s in xls.sheet_names}
            website_sheet = sheet_lookup.get("WEBSITE")
            store_sheet = sheet_lookup.get("STORE")
            if website_sheet is None or store_sheet is None:
                diagnostics["error"] = f"Required sheets WEBSITE and STORE not found. Found sheets: {', '.join(xls.sheet_names)}"
                return pd.DataFrame(), diagnostics
            website = pd.read_excel(combined_path, sheet_name=website_sheet, engine="openpyxl")
            store = pd.read_excel(combined_path, sheet_name=store_sheet, engine="openpyxl")
        else:
            diagnostics["error"] = "No data files found. Add either data/website_data.xlsx + data/store_data.xlsx, or data/customer_data.xlsx."
            return pd.DataFrame(), diagnostics
    except Exception as e:
        diagnostics["error"] = str(e)
        return pd.DataFrame(), diagnostics

    website.columns = [str(c).strip() for c in website.columns]
    store.columns = [str(c).strip() for c in store.columns]

    diagnostics["website_columns"] = website.columns.tolist()
    diagnostics["store_columns"] = store.columns.tolist()
    diagnostics["rows_raw"] = {"website": len(website), "store": len(store)}

    # Website mapping: supports both your actual pivot-like headers and simple headers.
    w = pd.DataFrame(index=website.index)
    w["Source"] = "WEBSITE"
    w["Txn_Date"] = pick_series(website, ["Date", "Txn_Date", "Order Date"])
    w["Customer_Name"] = pick_series(website, ["Customer Name", "NAME", "Customer_Name"])
    w["Customer_Mobile"] = pick_series(website, ["Customer Mobile", "CUSTOMER_MOBILE", "Mobile", "Phone"])
    w["Store_Name"] = "WEBSITE"
    w["Bills"] = pick_series(website, ["BILLS", "Distinct Count of Bill No.", "Distinct Count of Bill No", "Distinct Count of BILL_NO", "Order ID", "Orders"], 0)
    w["Qty"] = pick_series(website, ["Qty.", "Sum of Qty.", "Qty", "Quantity", "Sum of Qty"], 0)
    w["Gross_Amt"] = pick_series(website, ["Gross Amt.", "Sum of Gross Amt.", "Gross Amt", "Sum of Gross Amt"], 0)
    w["Taxable_Amt"] = pick_series(website, ["Taxable Amt.", "Sum of Taxable Amt.", "Taxable Amt", "Sum of Taxable Amt"], 0)
    w["Discount_Amt"] = pick_series(website, ["Discount", "Sum of Discount", "Discount Amt", "Discount_Amt"], 0)
    w["Net_Amt"] = pick_series(website, ["Net Amount", "Sum of Net Amount", "NET_AMOUNT", "Net_Amt"], 0)

    # Store mapping.
    s = pd.DataFrame(index=store.index)
    s["Source"] = "STORE"
    s["Txn_Date"] = pick_series(store, ["BILL_DATE", "Bill Date", "Txn_Date", "Date"])
    s["Customer_Name"] = pick_series(store, ["CUSTOMER_NAME", "Customer Name", "NAME"])
    s["Customer_Mobile"] = pick_series(store, ["CUSTOMER_MOBILE", "Customer Mobile", "Mobile", "Phone"])
    s["Store_Name"] = pick_series(store, ["NAME", "SITE_NAME", "Store Name", "Store_Name"], "STORE")
    s["Bills"] = pick_series(store, ["BILLS", "Distinct Count of BILL_NO", "Distinct Count of Bill No.", "Distinct Count of Bill No", "BILL_NO"], 0)
    s["Qty"] = pick_series(store, ["QTY", "Sum of NET_QTY", "NET_QTY", "Qty", "Quantity"], 0)
    s["Gross_Amt"] = pick_series(store, ["GROSS AMOUMT", "GROSS AMOUNT", "Sum of GROSS AMOUMT", "Sum of GROSS AMOUNT", "Gross Amt"], 0)
    s["Taxable_Amt"] = pick_series(store, ["TAXABLE_AMOUNT", "Sum of TAXABLE_AMOUNT", "Taxable Amt"], 0)
    s["Discount_Amt"] = pick_series(store, ["BILL_DISCOUNT_AMOUNT", "Sum of BILL_DISCOUNT_AMOUNT", "Discount"], 0)
    s["Net_Amt"] = pick_series(store, ["NET_AMOUNT", "Sum of NET_AMOUNT", "Net Amount"], 0)

    diagnostics["mapped_columns"] = {
        "website": {"rows": len(w)},
        "store": {"rows": len(s)},
    }

    df = pd.concat([w, s], ignore_index=True)
    df["Txn_Date"] = pd.to_datetime(df["Txn_Date"], errors="coerce")
    df["Customer_Name"] = df["Customer_Name"].fillna("").astype(str).str.strip()
    df["Store_Name"] = df["Store_Name"].fillna("UNKNOWN").astype(str).str.strip().replace("", "UNKNOWN")
    df["Clean_Mobile"] = df["Customer_Mobile"].apply(clean_mobile_value)

    for col in ["Bills", "Qty", "Gross_Amt", "Taxable_Amt", "Discount_Amt", "Net_Amt"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    before_clean = len(df)
    df = df.dropna(subset=["Txn_Date", "Clean_Mobile"])
    df = df[df["Clean_Mobile"].astype(str).str.len() == 10]
    df["Month"] = df["Txn_Date"].dt.to_period("M").dt.to_timestamp()
    df["Quarter"] = df["Txn_Date"].dt.to_period("Q").astype(str)

    diagnostics["rows_after_clean"] = len(df)
    diagnostics["rows_removed_invalid_date_mobile"] = before_clean - len(df)
    diagnostics["source_counts"] = df["Source"].value_counts().to_dict() if not df.empty else {}
    return df, diagnostics

def to_excel_bytes(tables: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet, table in tables.items():
            safe_sheet = sheet[:31]
            table.to_excel(writer, sheet_name=safe_sheet, index=False)
            worksheet = writer.sheets[safe_sheet]
            for idx, col in enumerate(table.columns):
                width = min(max(len(str(col)) + 2, 12), 36)
                worksheet.set_column(idx, idx, width)
    return output.getvalue()


st.markdown(f"<div class='main-title'>{APP_TITLE}</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-title'>Online vs Retail Overlap · Repeat Customers · Buying Frequency · Store-wise Analysis</div>", unsafe_allow_html=True)

st.sidebar.markdown("---")
with st.sidebar.expander("Deployment Debug", expanded=False):
    st.write("Current folder:", str(Path.cwd()))
    st.write("Data folder exists:", DATA_DIR.exists())
    if DATA_DIR.exists():
        st.write("Files inside data folder:")
        st.write([p.name for p in DATA_DIR.iterdir()])
    st.write("website_data.xlsx found:", WEBSITE_FILE.exists())
    st.write("store_data.xlsx found:", STORE_FILE.exists())
    st.write("customer_data.xlsx found:", COMBINED_FILE.exists())

if not ((WEBSITE_FILE.exists() and STORE_FILE.exists()) or COMBINED_FILE.exists()):
    st.error("Excel data files not found in the GitHub/Streamlit app folder.")
    found_files = [str(p) for p in DATA_DIR.rglob("*")] if DATA_DIR.exists() else []
    st.markdown(
        """
        The app is looking for exactly one of these options:

        **Recommended separate files**
        - `data/website_data.xlsx`
        - `data/store_data.xlsx`

        **Or old combined file**
        - `data/customer_data.xlsx` with sheets `WEBSITE` and `STORE`
        """
    )
    st.write("Files currently found inside `data` folder:", found_files if found_files else "No files found")
    st.stop()

with st.spinner("Loading dashboard data. Large files can take a little time..."):
    df, diagnostics = load_customer_master(str(COMBINED_FILE), str(WEBSITE_FILE), str(STORE_FILE))

if df.empty:
    st.error("No valid data found. Check sheet names, date columns, and mobile number columns.")
    with st.expander("Show data health details"):
        st.json(diagnostics)
    st.stop()

# Sidebar filters
st.sidebar.header("Filters")
min_date = df["Txn_Date"].min().date()
max_date = df["Txn_Date"].max().date()
quick_range = st.sidebar.radio("Quick Date View", ["Full data", "Last 90 days", "Custom"], horizontal=False)
if quick_range == "Last 90 days":
    default_start = max(min_date, max_date - timedelta(days=90))
    default_end = max_date
else:
    default_start = min_date
    default_end = max_date

selected_range = st.sidebar.date_input(
    "Date Range",
    value=(default_start, default_end),
    min_value=min_date,
    max_value=max_date,
)
if isinstance(selected_range, tuple) and len(selected_range) == 2:
    start_date, end_date = selected_range
else:
    start_date, end_date = default_start, default_end

store_options = sorted(df.loc[df["Source"] == "STORE", "Store_Name"].dropna().unique().tolist())
selected_stores = st.sidebar.multiselect("Store", options=store_options, default=[])
selected_sources = st.sidebar.multiselect("Source", options=["WEBSITE", "STORE"], default=["WEBSITE", "STORE"])
search_text = st.sidebar.text_input("Search mobile/name", "").strip().lower()

with st.sidebar.expander("Data Health"):
    st.write(f"Website rows loaded: {num_fmt(diagnostics.get('source_counts', {}).get('WEBSITE', 0))}")
    st.write(f"Store rows loaded: {num_fmt(diagnostics.get('source_counts', {}).get('STORE', 0))}")
    st.write(f"Date range: {min_date} to {max_date}")
    st.write(f"Store options: {len(store_options)}")

# Main filtered data
mask = (df["Txn_Date"].dt.date >= start_date) & (df["Txn_Date"].dt.date <= end_date)
if selected_sources:
    mask &= df["Source"].isin(selected_sources)
if selected_stores:
    mask &= ((df["Source"] == "WEBSITE") | (df["Store_Name"].isin(selected_stores)))
if search_text:
    mask &= (
        df["Clean_Mobile"].astype(str).str.contains(search_text, na=False)
        | df["Customer_Name"].str.lower().str.contains(search_text, na=False)
    )
filtered = df.loc[mask].copy()

# Overlap base uses date/store/search filters but ignores Source slicer so both sides can still compare.
overlap_mask = (df["Txn_Date"].dt.date >= start_date) & (df["Txn_Date"].dt.date <= end_date)
if selected_stores:
    overlap_mask &= ((df["Source"] == "WEBSITE") | (df["Store_Name"].isin(selected_stores)))
if search_text:
    overlap_mask &= (
        df["Clean_Mobile"].astype(str).str.contains(search_text, na=False)
        | df["Customer_Name"].str.lower().str.contains(search_text, na=False)
    )
overlap_base = df.loc[overlap_mask].copy()
website_mobiles = set(overlap_base.loc[overlap_base["Source"] == "WEBSITE", "Clean_Mobile"].dropna().unique())
store_mobiles = set(overlap_base.loc[overlap_base["Source"] == "STORE", "Clean_Mobile"].dropna().unique())
overlap_mobiles = website_mobiles.intersection(store_mobiles)

website_customers = len(website_mobiles)
store_customers = len(store_mobiles)
overlap_customers = len(overlap_mobiles)
website_to_store_pct = (overlap_customers / website_customers * 100) if website_customers else 0
store_to_website_pct = (overlap_customers / store_customers * 100) if store_customers else 0

store_only = overlap_base[overlap_base["Source"] == "STORE"].copy()
if not store_only.empty:
    store_customer_bills = store_only.groupby("Clean_Mobile", as_index=False).agg(
        Customer_Name=("Customer_Name", "last"),
        Bills=("Bills", "sum"),
        Net_Amt=("Net_Amt", "sum"),
        First_Visit=("Txn_Date", "min"),
        Last_Visit=("Txn_Date", "max"),
    )
else:
    store_customer_bills = pd.DataFrame(columns=["Clean_Mobile", "Customer_Name", "Bills", "Net_Amt", "First_Visit", "Last_Visit"])

repeat_customers = int((store_customer_bills["Bills"] > 1).sum()) if not store_customer_bills.empty else 0
new_customers = int((store_customer_bills["Bills"] == 1).sum()) if not store_customer_bills.empty else 0
repeat_pct = (repeat_customers / store_customers * 100) if store_customers else 0

st.caption(f"Loaded {num_fmt(len(df))} valid rows · Showing {num_fmt(len(filtered))} filtered rows · Data from {min_date} to {max_date}")
if len(filtered) == 0:
    st.warning("No rows for the selected filters. Clear the search box or change Date Range to Full data / Last 90 days.")

# KPI cards
c1, c2, c3, c4 = st.columns(4)
with c1:
    metric_card("Website Customers", num_fmt(website_customers), "Unique clean mobile numbers")
with c2:
    metric_card("Store Customers", num_fmt(store_customers), "Unique clean mobile numbers")
with c3:
    metric_card("Overlap Customers", num_fmt(overlap_customers), "Customers found in both")
with c4:
    metric_card("Total Revenue", money_fmt(filtered["Net_Amt"].sum()) if not filtered.empty else "₹0", "Filtered net amount")

c5, c6, c7, c8 = st.columns(4)
with c5:
    metric_card("Website → Store %", pct_fmt(website_to_store_pct), "Online customers who also shopped store")
with c6:
    metric_card("Store → Website %", pct_fmt(store_to_website_pct), "Store customers who also shopped online")
with c7:
    metric_card("Repeat Customers", num_fmt(repeat_customers), "Store customers with Bills > 1")
with c8:
    metric_card("Repeat %", pct_fmt(repeat_pct), "Repeat customers / store customers")

st.markdown("---")

tab1, tab2, tab3, tab4 = st.tabs(["📌 Overview", "🔁 Overlap", "👥 Repeat & Frequency", "⬇️ Export Tables"])

with tab1:
    left, right = st.columns(2)
    monthly = filtered.groupby(["Month", "Source"], as_index=False).agg(
        Customers=("Clean_Mobile", "nunique"), Revenue=("Net_Amt", "sum"), Bills=("Bills", "sum")
    ) if not filtered.empty else pd.DataFrame()
    with left:
        st.subheader("Monthly Customers")
        if not monthly.empty:
            st.line_chart(monthly.pivot(index="Month", columns="Source", values="Customers").fillna(0))
        else:
            st.info("No data for selected filters.")
    with right:
        st.subheader("Monthly Revenue")
        if not monthly.empty:
            st.line_chart(monthly.pivot(index="Month", columns="Source", values="Revenue").fillna(0))
        else:
            st.info("No data for selected filters.")

    st.subheader("Store-wise Summary")
    if not store_only.empty:
        store_summary = store_only.groupby("Store_Name", as_index=False).agg(
            Store_Customers=("Clean_Mobile", "nunique"),
            Bills=("Bills", "sum"),
            Revenue=("Net_Amt", "sum"),
        )
        store_overlap_rows = []
        for store_name, g in store_only.groupby("Store_Name"):
            sm = set(g["Clean_Mobile"].dropna().unique())
            ov = len(sm.intersection(website_mobiles))
            store_overlap_rows.append({"Store_Name": store_name, "Overlap_Customers": ov})
        store_overlap_df = pd.DataFrame(store_overlap_rows)
        store_summary = store_summary.merge(store_overlap_df, on="Store_Name", how="left")
        store_summary["Overlap_%"] = (store_summary["Overlap_Customers"] / store_summary["Store_Customers"] * 100).fillna(0).round(1)
        st.dataframe(store_summary.sort_values("Revenue", ascending=False), width="stretch", hide_index=True)
    else:
        store_summary = pd.DataFrame()
        st.info("No store data for selected filters.")

with tab2:
    st.subheader("Online vs Retail Overlap")
    overlap_customers_df = overlap_base[overlap_base["Clean_Mobile"].isin(overlap_mobiles)].copy()
    ov_left, ov_right = st.columns(2)
    with ov_left:
        st.write("Monthly Overlap Trend")
        ov_month = overlap_customers_df.groupby("Month", as_index=False).agg(Overlap_Customers=("Clean_Mobile", "nunique")) if not overlap_customers_df.empty else pd.DataFrame()
        if not ov_month.empty:
            st.line_chart(ov_month.set_index("Month"))
        else:
            st.info("No overlap found for selected filters.")
    with ov_right:
        st.write("Quarterly Overlap Trend")
        ov_quarter = overlap_customers_df.groupby("Quarter", as_index=False).agg(Overlap_Customers=("Clean_Mobile", "nunique")) if not overlap_customers_df.empty else pd.DataFrame()
        if not ov_quarter.empty:
            st.bar_chart(ov_quarter.set_index("Quarter"))
        else:
            st.info("No overlap found for selected filters.")

    st.subheader("Overlap Customer List")
    if not overlap_customers_df.empty:
        overlap_customer_table = overlap_customers_df.groupby("Clean_Mobile", as_index=False).agg(
            Customer_Name=("Customer_Name", "last"),
            First_Date=("Txn_Date", "min"),
            Last_Date=("Txn_Date", "max"),
            Sources=("Source", lambda x: ", ".join(sorted(set(x)))),
            Stores=("Store_Name", lambda x: ", ".join(sorted(set([v for v in x if v != "WEBSITE"])))[:200]),
            Total_Bills=("Bills", "sum"),
            Revenue=("Net_Amt", "sum"),
        )
        st.dataframe(overlap_customer_table, width="stretch", hide_index=True)
    else:
        overlap_customer_table = pd.DataFrame()
        st.info("No overlap customers for selected filters.")

with tab3:
    st.subheader("Repeat Customers & Buying Frequency")
    if not store_customer_bills.empty:
        store_customer_bills["Frequency_Bucket"] = store_customer_bills["Bills"].apply(frequency_bucket)
        bucket_order = ["1 Time", "2 Times", "3 Times", "4-5 Times", "6-10 Times", "11+ Times"]
        bucket_summary = store_customer_bills.groupby("Frequency_Bucket", as_index=False).agg(Customers=("Clean_Mobile", "nunique"))
        bucket_summary["Frequency_Bucket"] = pd.Categorical(bucket_summary["Frequency_Bucket"], categories=bucket_order, ordered=True)
        bucket_summary = bucket_summary.sort_values("Frequency_Bucket")
        st.bar_chart(bucket_summary.set_index("Frequency_Bucket"))

        st.subheader("Customer-level Frequency Table")
        st.dataframe(store_customer_bills.sort_values("Bills", ascending=False), width="stretch", hide_index=True)
    else:
        st.info("No store data for selected filters.")

with tab4:
    st.subheader("Export Filtered Data")
    export_cols = ["Source", "Txn_Date", "Customer_Name", "Clean_Mobile", "Store_Name", "Bills", "Qty", "Gross_Amt", "Taxable_Amt", "Discount_Amt", "Net_Amt"]
    filtered_export = filtered[export_cols].copy() if not filtered.empty else pd.DataFrame(columns=export_cols)
    if not filtered_export.empty:
        filtered_export["Txn_Date"] = filtered_export["Txn_Date"].dt.date
    st.dataframe(filtered_export.head(1000), width="stretch", hide_index=True)
    st.caption("Preview shows first 1,000 rows only. Export includes all filtered rows.")

    csv_bytes = filtered_export.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Download Filtered CSV",
        data=csv_bytes,
        file_name="suta_filtered_customer_data.csv",
        mime="text/csv",
    )

    tables = {"Filtered Data": filtered_export}
    if 'store_summary' in locals() and not store_summary.empty:
        tables["Store Summary"] = store_summary
    if 'overlap_customer_table' in locals() and not overlap_customer_table.empty:
        tables["Overlap Customers"] = overlap_customer_table
    if not store_customer_bills.empty:
        tables["Customer Frequency"] = store_customer_bills
    excel_bytes = to_excel_bytes(tables)
    st.download_button(
        "Download Excel Report",
        data=excel_bytes,
        file_name="suta_customer_analytics_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("---")
st.markdown("<div class='small-note'>Data source is controlled from the <b>data</b> folder. Use <b>website_data.xlsx</b> + <b>store_data.xlsx</b>, or one combined <b>customer_data.xlsx</b>. Viewers cannot upload or edit source data from this app.</div>", unsafe_allow_html=True)
