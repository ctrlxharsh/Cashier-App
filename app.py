import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Cashier Reconciliation App", page_icon="💸", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans:wght@400;600&family=Noto+Sans+Devanagari:wght@400;600&display=swap');

/* Add a bit of custom styling to make it look professional */
html, body, [class*="st-"] {
    font-family: 'Noto Sans', 'Noto Sans Devanagari', sans-serif;
}
.main {
    background-color: #f8f9fa;
}
.stButton button {
    background-color: #0066cc;
    color: white;
    font-weight: bold;
    border-radius: 5px;
    padding: 0.5rem 1rem;
    transition: all 0.3s;
}
.stButton button:hover {
    background-color: #0052a3;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
}
h1, h2, h3 {
    color: #2c3e50;
    font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

def clean_data(df):
    """
    Remove special characters and trim spaces.
    Only letters, numbers, and spaces are kept.
    """
    cleaned_df = df.copy()
    for col in cleaned_df.columns:
        if cleaned_df[col].dtype == 'object' or cleaned_df[col].dtype.name == 'string':
            # Remove special chars, keeping alphanumeric and spaces, then trim
            cleaned_df[col] = cleaned_df[col].astype(str).replace(r'[^a-zA-Z0-9\s]', '', regex=True).str.strip()
    return cleaned_df

def load_excel_safely(uploaded_file):
    try:
        # First try normal modern Excel format
        uploaded_file.seek(0)
        return pd.ExcelFile(uploaded_file, engine='openpyxl')
    except Exception as e:
        err_str = str(e).lower()
        if "zip" in err_str or "format" in err_str:
            # Might be an old XLS or an HTML file masked as XLS (common from banking portals)
            uploaded_file.seek(0)
            try:
                # Try reading as old XLS
                return pd.ExcelFile(uploaded_file, engine='xlrd')
            except Exception:
                pass
                
            uploaded_file.seek(0)
            try:
                # Try reading as HTML table
                dfs = pd.read_html(uploaded_file)
                if dfs:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        dfs[0].to_excel(writer, index=False, sheet_name='Data')
                    output.seek(0)
                    return pd.ExcelFile(output, engine='openpyxl')
            except Exception:
                pass
        raise Exception("Could not parse file. It is not a valid modern Excel, legacy Excel, or HTML format.")

def match_entries(new_df, old_df, name_col, skip_cols):
    missing_in_old = []
    mismatches = []
    
    # Track used old indices to handle duplicates properly
    matched_old_indices = set()
    
    compare_cols = [c for c in new_df.columns if c != name_col and c not in skip_cols and c in old_df.columns]
    
    for new_idx, new_row in new_df.iterrows():
        name = str(new_row[name_col]).strip() if pd.notna(new_row[name_col]) else ""
        if not name:
            continue
            
        old_candidates = old_df[old_df[name_col].astype(str).str.strip() == name]
        
        if old_candidates.empty:
            missing_in_old.append(new_row.to_dict())
            continue
            
        # 1. Try to find an exact match on all compare_cols first
        exact_match_found = False
        for old_idx, old_row in old_candidates.iterrows():
            if old_idx in matched_old_indices:
                continue
                
            is_exact = True
            for col in compare_cols:
                if str(new_row[col]).strip() != str(old_row[col]).strip():
                    is_exact = False
                    break
            
            if is_exact:
                exact_match_found = True
                matched_old_indices.add(old_idx)
                break
                
        if exact_match_found:
            continue
            
        # 2. If no exact match found, use the first available candidate
        partial_match_index = None
        for old_idx, old_row in old_candidates.iterrows():
            if old_idx not in matched_old_indices:
                partial_match_index = old_idx
                break
                
        if partial_match_index is not None:
            matched_old_indices.add(partial_match_index)
            old_row = old_df.loc[partial_match_index]
            
            mismatch_details = {"Name": name, "Differences": {}}
            for col in compare_cols:
                new_val = str(new_row[col]).strip()
                old_val = str(old_row[col]).strip()
                if new_val != old_val:
                    mismatch_details["Differences"][col] = f"Old: {old_val} -> New: {new_val}"
            
            if mismatch_details["Differences"]:
                # Also include context about which row this is, if available
                if 'S.No.' in new_row:
                    mismatch_details["S.No."] = new_row["S.No."]
                elif 'S.No ' in new_row:
                    mismatch_details["S.No."] = new_row["S.No "]
                mismatches.append(mismatch_details)
        else:
            # All candidates were matched to other new_rows with this name
            missing_in_old.append(new_row.to_dict())

    return missing_in_old, mismatches

st.title("💸 Cashier Reconciliation Portal")
st.markdown("Easily compare New spreadsheet entries against Old spreadsheet entries to identify new accounts and mismatches.")

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Upload Old Data")
    old_file = st.file_uploader("Select Old Excel File", type=["xlsx", "xls", "csv"], key="old_file")

with col2:
    st.subheader("📁 Upload New Data")
    new_file = st.file_uploader("Select New Excel File", type=["xlsx", "xls", "csv"], key="new_file")

if old_file and new_file:
    try:
        old_xl = load_excel_safely(old_file)
        new_xl = load_excel_safely(new_file)
        
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            old_sheet = st.selectbox("Select Sheet for Old Data", old_xl.sheet_names, key="old_sheet")
            old_header_row = st.number_input("Header Row for Old Data (starts at 13)", min_value=1, value=13, key="old_header")
        with col_s2:
            new_sheet = st.selectbox("Select Sheet for New Data", new_xl.sheet_names, key="new_sheet")
            new_header_row = st.number_input("Header Row for New Data (starts at 13)", min_value=1, value=13, key="new_header")

        old_df_raw = old_xl.parse(old_sheet, skiprows=old_header_row - 1)
        new_df_raw = new_xl.parse(new_sheet, skiprows=new_header_row - 1)
        
        st.divider()
        st.subheader("⚙️ Configuration")
        
        default_skips = ['Net Payble', 'Amount', 'S.No.', 'S.No ']
        all_cols = list(new_df_raw.columns)
        
        name_col_options = [c for c in all_cols if 'name' in str(c).lower()]
        default_name_index = all_cols.index(name_col_options[0]) if name_col_options else 0
        name_col = st.selectbox("Select the 'Name' identifier column:", all_cols, index=default_name_index)
        
        available_skips = [c for c in all_cols if c != name_col]
        prefilled_skips = [c for c in default_skips if c in available_skips]
        
        skip_cols = st.multiselect("Select fields to SKIP during comparison:", available_skips, default=prefilled_skips)
        
        st.write("")
        if st.button("🚀 Run Reconciliation Analysis"):
            with st.spinner("Cleaning and Matching Data..."):
                old_df = clean_data(old_df_raw)
                new_df = clean_data(new_df_raw)
                
                missing_in_old, mismatches = match_entries(new_df, old_df, name_col, skip_cols)
                
                st.success("✅ Analysis Complete!")
                
                tab1, tab2, tab3 = st.tabs(["🆕 Missing in Old (New Entries)", "⚠️ Mismatched Details", "📥 Download Cleaned Data"])
                
                with tab1:
                    if missing_in_old:
                        st.info(f"Found {len(missing_in_old)} entries present in New but missing in Old:")
                        st.dataframe(pd.DataFrame(missing_in_old), use_container_width=True)
                    else:
                        st.success("No missing entries found. All names in New are present in Old.")
                        
                with tab2:
                    if mismatches:
                        st.warning(f"Found {len(mismatches)} entries with mismatched fields (excluding skipped fields):")
                        for m in mismatches:
                            with st.expander(f"Mismatch for Name: {m['Name']}", expanded=True):
                                if "S.No." in m:
                                    st.write(f"**Row S.No.:** {m['S.No.']}")
                                for col, diff in m['Differences'].items():
                                    st.markdown(f"- **{col}**: {diff}")
                    else:
                        st.success("No mismatched fields found for matched names.")
                        
                with tab3:
                    st.write("Download the cleaned version of the entire New Excel file (all sheets, all rows, special characters removed).")
                    
                    import openpyxl
                    import re
                    
                    wb = openpyxl.load_workbook(new_file)
                    for sheet in wb.worksheets:
                        for row in sheet.iter_rows():
                            for cell in row:
                                if isinstance(cell.value, str):
                                    cell.value = re.sub(r'[^a-zA-Z0-9\s]', '', cell.value).strip()
                                    
                    output = io.BytesIO()
                    wb.save(output)
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label="⬇️ Download Cleaned New Sheet",
                        data=excel_data,
                        file_name="cleaned_new_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                
    except Exception as e:
        error_msg = str(e)
        st.error(f"Error processing files: {error_msg}")
else:
    st.info("Please upload both Old and New Excel files to begin.")
