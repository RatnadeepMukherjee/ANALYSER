import streamlit as st
import pdfplumber
import pandas as pd
import openpyxl
import io
from io import BytesIO

st.set_page_config(page_title="Multi-File to Spreadsheet", layout="wide")

st.title("Multi-File (CSV/Excel/PDF) to Spreadsheet Converter")

# Session State
if 'master_df' not in st.session_state:
    st.session_state.master_df = pd.DataFrame()
if 'logs' not in st.session_state:
    st.session_state.logs = []

def extract_pdf_data(pdf_file):
    data_dfs = []
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    data_dfs.append(df)
                else:
                    text = page.extract_text()
                    if text:
                        lines = [line.split() for line in text.split('\n') if line.strip()]
                        if lines:
                            df = pd.DataFrame(lines[1:], columns=lines[0])
                            data_dfs.append(df)
        if data_dfs:
            return pd.concat(data_dfs, ignore_index=True)
    except Exception as e:
        st.session_state.logs.append(f"PDF Error in {pdf_file.name}: {str(e)}")
    return pd.DataFrame()

def process_excel(file):
    data_dfs = []
    try:
        excel_dfs = pd.read_excel(file, sheet_name=None)
        for sheet_name, df in excel_dfs.items():
            data_dfs.append(df)
        if data_dfs:
            return pd.concat(data_dfs, ignore_index=True)
    except Exception as e:
        st.session_state.logs.append(f"Excel Error in {file.name}: {str(e)}")
    return pd.DataFrame()

def process_csv(file):
    try:
        df = pd.read_csv(file)
        return df
    except Exception as e:
        st.session_state.logs.append(f"CSV Error in {file.name}: {str(e)}")
    return pd.DataFrame()

def clean_data(df):
    if df.empty:
        return df
    df = df.dropna(how='all')
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    df = df.fillna('')
    return df

def process_payroll_data(df):
    if df.empty:
        return df
    col_map = {
        'employee_id': ['employee', 'emp', 'id'],
        'ssn': ['ssn', 'social'],
        'first_name': ['first', 'fname', 'name first'],
        'last_name': ['last', 'lname', 'name last']
    }
    selected_cols = {}
    for target, keywords in col_map.items():
        for col in df.columns:
            if any(k in str(col).lower() for k in keywords):
                selected_cols[target] = col
                break
    
    payroll_df = pd.DataFrame()
    if selected_cols:
        for target, col in selected_cols.items():
            payroll_df[target] = df[col]
    
    code_col = next((c for c in df.columns if 'code' in str(c).lower()), None)
    amount_col = next((c for c in df.columns if any(x in str(c).lower() for x in ['amount', 'deduct', 'earn'])), None)
    
    if code_col and amount_col:
        filtered = df[df[code_col].astype(str).str.lower().str.contains('post|pre|401k', na=False) & 
                      ~df[code_col].astype(str).str.lower().str.contains('non.*taxable', na=False, regex=True)]
        if not filtered.empty:
            group_cols = list(selected_cols.values())
            if group_cols:
                ded_earn = filtered.groupby(group_cols)[amount_col].sum().reset_index()
                ded_earn.rename(columns={amount_col: 'filtered_deductions_earnings'}, inplace=True)
                payroll_df = pd.merge(payroll_df, ded_earn, left_on=group_cols, right_on=group_cols, how='left')
    
    return payroll_df

def process_file(file, apply_payroll):
    if file.name.lower().endswith('.pdf'):
        raw = extract_pdf_data(file)
    elif file.name.lower().endswith(('.xlsx', '.xls')):
        raw = process_excel(file)
    elif file.name.lower().endswith('.csv'):
        raw = process_csv(file)
    else:
        st.session_state.logs.append(f"Unsupported file type: {file.name}")
        return pd.DataFrame()
    
    cleaned = clean_data(raw)
    processed = process_payroll_data(cleaned) if apply_payroll else cleaned
    processed['source'] = file.name
    return processed

tab1, tab2, tab3 = st.tabs(["Upload Files", "Preview", "Download"])

with tab1:
    uploaded_files = st.file_uploader("Upload CSV, Excel, or PDF files", type=['csv', 'xlsx', 'xls', 'pdf'], accept_multiple_files=True)
    apply_payroll = st.checkbox("Apply Payroll Processing (filters & maps payroll data)", value=False)
    
    if st.button("Process Files", type="primary") and uploaded_files:
        st.session_state.master_df = pd.DataFrame()
        st.session_state.logs = []
        progress = st.progress(0)
        total_files = len(uploaded_files)
        for i, file in enumerate(uploaded_files):
            processed_df = process_file(file, apply_payroll)
            rows_added = len(processed_df)
            st.session_state.master_df = pd.concat([st.session_state.master_df, processed_df], ignore_index=True)
            st.session_state.logs.append(f"Processed {file.name}: {rows_added} rows")
            progress.progress((i + 1) / total_files)
        st.success(f"Processed {total_files} files. Total rows: {len(st.session_state.master_df)}")
        st.rerun()

with tab2:
    if not st.session_state.master_df.empty:
        st.metric("Total Files Processed", len(st.session_state.logs))
        st.metric("Total Rows", len(st.session_state.master_df))
        st.dataframe(st.session_state.master_df, use_container_width=True)
    else:
        st.info("Upload and process files first.")

with tab3:
    if not st.session_state.master_df.empty:
        xlsx_buffer = BytesIO()
        with pd.ExcelWriter(xlsx_buffer, engine='openpyxl') as writer:
            st.session_state.master_df.to_excel(writer, index=False, sheet_name='Merged Data')
        st.download_button(
            label="Download Merged Excel",
            data=xlsx_buffer.getvalue(),
            file_name="merged_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No data to download.")

if st.session_state.logs:
    st.sidebar.title("Processing Logs")
    for log in st.session_state.logs:
        st.sidebar.write(log)
