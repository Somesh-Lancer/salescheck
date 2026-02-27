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

# --- CORE PROCESSING FUNCTION ---
def process_reconciliation(sales_df, main_df):
    # Standardize Order IDs before merging
    if 'Order ID' in sales_df.columns:
        sales_df['Order ID'] = sales_df['Order ID'].astype(str).str.strip()
    if 'Order ID' in main_df.columns:
        main_df['Order ID'] = main_df['Order ID'].astype(str).str.strip()

    # Merge side-by-side using OUTER join to catch missing records on both ends
    merged = pd.merge(sales_df, main_df, on='Order ID', how='outer')

    # Status Logic
    def check_status(row):
        # If either side is missing completely, the Order ID wasn't found in one of the files
        if pd.isna(row.get('Tax Exclusive Gross')) or pd.isna(row.get('product sales')):
            return "Order ID Not Found"
        
        # Check for value matches
        try:
            gross_match = round(float(row['Tax Exclusive Gross']), 2) == round(float(row['product sales']), 2)
            tax_match = round(float(row['Total Tax Amount']), 2) == round(float(row['Total sales tax liable']), 2)
            
            if gross_match and tax_match:
                return "Values Match"
            else:
                return "Values Not Match"
        except (ValueError, TypeError):
            return "Values Not Match"

    merged['Status'] = merged.apply(check_status, axis=1)

    # Exact column order requested, with Status at the very end
    exact_columns_order = [
        # From B2B/B2C
        'Seller Gstin', 'Invoice Number', 'Invoice Date', 'Transaction Type', 'Order ID', 
        'Quantity', 'Item Description', 'Invoice Amount', 'Tax Exclusive Gross', 'Total Tax Amount',
        # From Main Excel
        'product sales', 'shipping credits', 'gift wrap credits', 'promotional rebates', 
        'Total sales tax liable', 'TCS-CGST', 'TCS-SGST', 'TCS-IGST', 'TDS (Section 194-O)', 
        'selling fees', 'fba fees', 'other transaction fees', 'other', 'total', 
        'date/time', 'settlement id', 
        # The new Status column
        'Status'
    ]
    
    # Keep only columns that exist and enforce the order
    final_cols = [c for c in exact_columns_order if c in merged.columns]
    final_df = merged[final_cols].fillna('-')
    
    return final_df

# Helper function to convert dataframe to Excel in memory
@st.cache_data
def convert_df_to_excel(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


# --- EXECUTION ---
if noscomp_file:
    # Load and clean the main file once
    df_nos = pd.read_excel(noscomp_file)
    df_nos.columns = df_nos.columns.astype(str).str.strip()
    
    for col in df_nos.columns:
        if str(col).lower() == "order id":
            df_nos.rename(columns={col: "Order ID"}, inplace=True)
        if "Total sales tax liable" in str(col):
            df_nos.rename(columns={col: "Total sales tax liable"}, inplace=True)

    st.divider()
    
    # Split the screen into two halves for B2B and B2C
    left_col, right_col = st.columns(2)

    # ==========================================
    # B2B COMPARISON SECTION
    # ==========================================
    with left_col:
        st.subheader("🏢 B2B Comparison")
        if b2b_file:
            with st.spinner('Processing B2B vs Main...'):
                df_b2b = pd.read_excel(b2b_file)
                df_b2b.columns = df_b2b.columns.astype(str).str.strip()
                for col in df_b2b.columns:
                    if str(col).lower() == "order id":
                        df_b2b.rename(columns={col: "Order ID"}, inplace=True)

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
                        key="download_b2b"
                    )
                except KeyError as e:
                    st.error(f"Missing Column in B2B: {e}")
        else:
            st.info("Upload the B2B form to generate this report.")

    # ==========================================
    # B2C COMPARISON SECTION
    # ==========================================
    with right_col:
        st.subheader("🛒 B2C Comparison")
        if b2c_file:
            with st.spinner('Processing B2C vs Main...'):
                df_b2c = pd.read_excel(b2c_file)
                df_b2c.columns = df_b2c.columns.astype(str).str.strip()
                for col in df_b2c.columns:
                    if str(col).lower() == "order id":
                        df_b2c.rename(columns={col: "Order ID"}, inplace=True)

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
                        key="download_b2c"
                    )
                except KeyError as e:
                    st.error(f"Missing Column in B2C: {e}")
        else:
            st.info("Upload the B2C form to generate this report.")

else:
    st.warning("⚠️ The Main Noscomp Report is required to run any comparisons. Please upload it above.")