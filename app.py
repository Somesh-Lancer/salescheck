import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="GST Reconciliation Tool", layout="wide")

st.title("📊 Noscomp Transaction Reconciliation 2025")
st.markdown("Upload your files to generate separate comparisons for B2B and B2C against the Main Noscomp Report.")

# 1. File Uploaders
col1, col2, col3 = st.columns(3)
with col1:
    b2b_file = st.file_uploader("Upload B2B Form (Excel)", type=['xlsx'])
with col2:
    b2c_file = st.file_uploader("Upload B2C Form (Excel)", type=['xlsx'])
with col3:
    noscomp_file = st.file_uploader("Upload Main Noscomp Report (Excel)", type=['xlsx'])


def _clean_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


def _ensure_order_id(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # normalize any "order id" variations to "Order ID"
    for c in list(df.columns):
        if str(c).strip().lower() == "order id":
            df.rename(columns={c: "Order ID"}, inplace=True)
    if "Order ID" in df.columns:
        df["Order ID"] = df["Order ID"].astype(str).str.strip()
    return df


def _find_type_col(df: pd.DataFrame) -> str | None:
    # possible column names for transaction type
    candidates = ["Transaction Type", "transaction type", "Type", "type", "Transaction type"]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _filter_by_transaction_type(df: pd.DataFrame, want: str) -> pd.DataFrame:
    """
    Keep rows where transaction type contains `want` (case-insensitive).
    If no type column exists, return df unchanged.
    """
    df = df.copy()
    tcol = _find_type_col(df)
    if not tcol:
        return df
    mask = df[tcol].astype(str).str.strip().str.lower().str.contains(want.lower(), na=False)
    return df[mask].copy()


# --- CORE PROCESSING FUNCTION ---
def process_reconciliation(sales_df: pd.DataFrame, main_df: pd.DataFrame) -> pd.DataFrame:
    sales_df = _ensure_order_id(_clean_cols(sales_df))
    main_df = _ensure_order_id(_clean_cols(main_df))

    # ✅ Apply your rule BEFORE merge:
    # Sales (B2B/B2C) -> only Shipment
    sales_df = _filter_by_transaction_type(sales_df, "shipment")

    # Main Noscomp -> only Order
    main_df = _filter_by_transaction_type(main_df, "order")

    # Outer merge to catch missing order ids too
    merged = pd.merge(sales_df, main_df, on="Order ID", how="outer", suffixes=("_sales", "_main"))

    # Status Logic
    def check_status(row):
        if pd.isna(row.get("Tax Exclusive Gross")) or pd.isna(row.get("product sales")):
            return "Order ID Not Found"

        try:
            gross_match = round(float(row["Tax Exclusive Gross"]), 2) == round(float(row["product sales"]), 2)
            tax_match = round(float(row["Total Tax Amount"]), 2) == round(float(row["Total sales tax liable"]), 2)

            if gross_match and tax_match:
                return "Values Match"
            return "Values Not Match"
        except (ValueError, TypeError):
            return "Values Not Match"

    merged["Status"] = merged.apply(check_status, axis=1)

    # ✅ Keep only ONE record per Order ID:
    # - If a match exists, keep the first matched row
    # - Else keep a "best" row: prefer non-zero Tax Exclusive Gross, else first row
    final_rows = []
    for order_id, group in merged.groupby("Order ID", dropna=False):
        matched = group[group["Status"] == "Values Match"]
        if not matched.empty:
            final_rows.append(matched.iloc[[0]])
            continue

        # prefer non-zero gross rows (avoid blank/zero)
        gross_num = pd.to_numeric(group.get("Tax Exclusive Gross"), errors="coerce")
        non_zero = group[(gross_num.notna()) & (gross_num != 0)]
        if not non_zero.empty:
            final_rows.append(non_zero.iloc[[0]])
        else:
            final_rows.append(group.iloc[[0]])

    final_df = pd.concat(final_rows, ignore_index=True)

    # Exact column order requested
    exact_columns_order = [
        # From B2B/B2C
        "Seller Gstin", "Invoice Number", "Invoice Date", "Transaction Type", "Order ID",
        "Quantity", "Item Description", "Invoice Amount", "Tax Exclusive Gross", "Total Tax Amount",
        # From Main Excel
        "product sales", "shipping credits", "gift wrap credits", "promotional rebates",
        "Total sales tax liable", "TCS-CGST", "TCS-SGST", "TCS-IGST", "TDS (Section 194-O)",
        "selling fees", "fba fees", "other transaction fees", "other", "total",
        "date/time", "settlement id",
        # Status
        "Status",
    ]

    final_cols = [c for c in exact_columns_order if c in final_df.columns]
    final_df = final_df[final_cols].fillna("-")

    return final_df


# Helper function to convert dataframe to Excel in memory
@st.cache_data
def convert_df_to_excel(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


# --- EXECUTION ---
if noscomp_file:
    df_nos = pd.read_excel(noscomp_file)
    df_nos = _ensure_order_id(_clean_cols(df_nos))

    # Normalize "Total sales tax liable" column name if it varies
    for col in list(df_nos.columns):
        if "total sales tax liable" in str(col).lower():
            df_nos.rename(columns={col: "Total sales tax liable"}, inplace=True)

    st.divider()
    left_col, right_col = st.columns(2)

    # B2B
    with left_col:
        st.subheader("🏢 B2B Comparison")
        if b2b_file:
            with st.spinner("Processing B2B vs Main..."):
                df_b2b = pd.read_excel(b2b_file)
                df_b2b = _ensure_order_id(_clean_cols(df_b2b))

                try:
                    result_b2b = process_reconciliation(df_b2b, df_nos)
                    st.success("B2B Processing Complete!")
                    st.dataframe(result_b2b, height=300, use_container_width=True)

                    excel_b2b = convert_df_to_excel(result_b2b, "B2B_Reconciliation")
                    st.download_button(
                        label="📥 Download B2B Result (Excel)",
                        data=excel_b2b,
                        file_name="B2B_vs_Main_Reconciliation.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_b2b",
                    )
                except KeyError as e:
                    st.error(f"Missing Column in B2B: {e}")
        else:
            st.info("Upload the B2B form to generate this report.")

    # B2C
    with right_col:
        st.subheader("🛒 B2C Comparison")
        if b2c_file:
            with st.spinner("Processing B2C vs Main..."):
                df_b2c = pd.read_excel(b2c_file)
                df_b2c = _ensure_order_id(_clean_cols(df_b2c))

                try:
                    result_b2c = process_reconciliation(df_b2c, df_nos)
                    st.success("B2C Processing Complete!")
                    st.dataframe(result_b2c, height=300, use_container_width=True)

                    excel_b2c = convert_df_to_excel(result_b2c, "B2C_Reconciliation")
                    st.download_button(
                        label="📥 Download B2C Result (Excel)",
                        data=excel_b2c,
                        file_name="B2C_vs_Main_Reconciliation.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_b2c",
                    )
                except KeyError as e:
                    st.error(f"Missing Column in B2C: {e}")
        else:
            st.info("Upload the B2C form to generate this report.")
else:
    st.warning("⚠️ The Main Noscomp Report is required to run any comparisons. Please upload it above.")