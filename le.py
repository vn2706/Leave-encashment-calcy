import streamlit as st
import pandas as pd
import numpy as np
import re
import calendar
from datetime import datetime
import io

# --- STATE LIMITS DICTIONARY ---
STATE_LIMITS = {
    'Rajasthan': 45, 'Chhattisgarh': 90, 'Maharashtra': 45, 'Uttarakhand': 45,
    'Delhi': 45, 'Karnataka': 45, 'Chandigarh': 45, 'Gujarat': 63,
    'Haryana': 45, 'Punjab': 45, 'Madhya Pradesh': 90, 'Goa': 45,
    'Assam': 45, 'Bihar': 45, 'Odisha': 45, 'West Bengal': 45,
    'Jharkhand': 45, 'Kerala': 45, 'Tamil Nadu': 45, 'Andhra Pradesh': 60,
    'Telangana': 60, 'Uttar Pradesh': 45
}

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Leave Encashment calcy", page_icon="🌴", layout="wide")
st.markdown("""
    <style>
    .stApp { background-color: #F4F6F9; }
    .main-header {
        background: linear-gradient(-45deg, #5E239D, #9D4EDD, #FF007F, #5E239D);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        text-align: center; font-size: 3.5rem; font-weight: 900; margin-bottom: -15px; padding-top: 10px;
    }
    .sub-header { text-align: center; color: #4B5563; font-size: 1.2rem; font-weight: 600; margin-bottom: 30px; }
    div.stButton > button { border-radius: 10px !important; font-weight: 700 !important; background-color: #5E239D !important; color: white !important; }
    </style>
""", unsafe_allow_html=True)

if 'current_page' not in st.session_state: st.session_state.current_page = "Home"
states_to_init = {'leave_preview_ready': False, 'lapse_calc_df': None, 'lapse_approved': False, 'pivot_df': None, 'pivot_approved': False}
for key, value in states_to_init.items():
    if key not in st.session_state: st.session_state[key] = value

def nav_home(): st.session_state.current_page = "Home"
def nav_leave(): st.session_state.current_page = "Leave"

def convert_df_to_styled_excel(report_df, lapse_df, master_df=None, master_sheet_name="Master Report"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        workbook = writer.book
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'fg_color': '#5E239D', 'font_color': 'white', 'border': 1})
        body_format = workbook.add_format({'border': 1})

        # --- TAB 1: MASTER CONSOLIDATED REPORT ---
        if master_df is not None:
            master_df.to_excel(writer, index=False, sheet_name=master_sheet_name)
            worksheet_m = writer.sheets[master_sheet_name]
            for col_num, value in enumerate(master_df.columns.values):
                worksheet_m.write(0, col_num, value, header_format)
            for row_num in range(1, len(master_df) + 1):
                for col_num in range(len(master_df.columns)):
                    val = master_df.iloc[row_num-1, col_num]
                    worksheet_m.write(row_num, col_num, "" if pd.isna(val) else val, body_format)

        # --- TAB 2: REPORT ---
        if report_df is not None:
            report_df.to_excel(writer, index=False, sheet_name='Report')
            worksheet1 = writer.sheets['Report']
            for col_num, value in enumerate(report_df.columns.values):
                worksheet1.write(0, col_num, value, header_format)
            for row_num in range(1, len(report_df) + 1):
                for col_num in range(len(report_df.columns)):
                    val = report_df.iloc[row_num-1, col_num]
                    worksheet1.write(row_num, col_num, "" if pd.isna(val) else val, body_format)
                    
        # --- TAB 3: LAPSE CALCULATION ---
        if lapse_df is not None:
            lapse_df.to_excel(writer, index=False, sheet_name='Lapse calculation')
            worksheet2 = writer.sheets['Lapse calculation']
            for col_num, value in enumerate(lapse_df.columns.values):
                worksheet2.write(0, col_num, value, header_format)
            for row_num in range(1, len(lapse_df) + 1):
                for col_num in range(len(lapse_df.columns)):
                    val = lapse_df.iloc[row_num-1, col_num]
                    worksheet2.write(row_num, col_num, "" if pd.isna(val) else val, body_format)
                    
    return output.getvalue()

def load_file(uploaded_file):
    uploaded_file.seek(0)
    if uploaded_file.name.endswith('.csv'):
        return pd.read_csv(uploaded_file, encoding='utf-8-sig', on_bad_lines='skip')
    else:
        return pd.read_excel(uploaded_file, engine='openpyxl')

def standardize_id(df, possible_names, report_name):
    def scrub(val): return re.sub(r'[^a-zA-Z0-9]', '', str(val)).lower()
    clean_possible = [scrub(p) for p in possible_names]
    for col in df.columns:
        if scrub(col) in clean_possible:
            df = df.copy()
            df['Base_ID'] = df[col].astype(str).str.strip()
            return df
    st.error(f"🚨 Could not find ID column in {report_name}"); st.stop()

st.markdown("<div class='main-header'>Leave Encashment calculator</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>Automated leave encashment calculation</div>", unsafe_allow_html=True)

with st.sidebar:
    st.button("🏠 Home", on_click=nav_home, use_container_width=True)
    st.button("🌴 2.1 Leave Encashment", on_click=nav_leave, use_container_width=True)

if st.session_state.current_page == "Leave":
    st.markdown("### 🌴 Step 2.1: Leave Encashment Calculator")
    with st.container(border=True):
        col1, col2, col3 = st.columns(3)
        with col1: l_time = st.file_uploader("1. Time Account Overview", type=["csv", "xlsx"])
        with col2: l_hc = st.file_uploader("2. HC Report", type=["csv", "xlsx"])
        with col3: l_ncp = st.file_uploader("3. NCP Input Sheet", type=["csv", "xlsx"])
            
    if l_time and l_hc and l_ncp:
        df_time_raw, df_hc_raw = load_file(l_time), load_file(l_hc)
        df_ncp = standardize_id(load_file(l_ncp), ["employeeid", "userid", "empid", "employee id"], "NCP Sheet")

        with st.expander("⚙️ NCP Month Detection & Per Day Calculation", expanded=not st.session_state.lapse_approved):
            accrual_val = st.number_input("Enter Default Value", value=1.25, format="%.10f")
            ncp_cols = [c for c in df_ncp.columns if 'ncp' in c.lower()]
            month_map = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}
            current_year = datetime.now().year
            if st.button("Proceed to Lapse Calculation"):
                lapse_df = df_ncp[['Base_ID'] + ncp_cols].copy()
                for col in ncp_cols:
                    m_match = re.search(r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)', col.lower())
                    if m_match:
                        m_name = m_match.group()
                        pd_val = accrual_val / calendar.monthrange(current_year, month_map[m_name])[1]
                        lapse_df[f"Lapse ({m_name})"] = pd.to_numeric(lapse_df[col], errors='coerce').fillna(0) * pd_val
                lapse_df['Total Lapse'] = lapse_df.filter(like='Lapse (').sum(axis=1).round(3)
                st.session_state.lapse_calc_df = lapse_df
                st.session_state.leave_preview_ready = True

        if st.session_state.leave_preview_ready:
            st.dataframe(st.session_state.lapse_calc_df, use_container_width=True)
            if st.button("Approve Lapse Calculation ✅"): st.session_state.lapse_approved = True

        if st.session_state.lapse_approved:
            st.markdown("### 📊 Step 2.2: Leave Balance Pivot Table")
            pivot_raw = pd.pivot_table(df_time_raw, index='External Person ID', columns='Time Account Type', values='Current Balance', aggfunc='sum', fill_value=0).reset_index()
            al_col = next((c for c in pivot_raw.columns if 'Annual' in str(c) or 'AL' in str(c)), None)
            cl_col = next((c for c in pivot_raw.columns if 'Casual' in str(c) or 'CL' in str(c)), None)
            final_pivot = pd.DataFrame({'Employee ID': pivot_raw['External Person ID'].astype(str).str.strip()})
            if al_col: final_pivot['AL balance'] = pivot_raw[al_col].round(3)
            if cl_col: final_pivot['CL balance'] = pivot_raw[cl_col].round(3)
            st.session_state.pivot_df = final_pivot
            st.dataframe(st.session_state.pivot_df, use_container_width=True)
            if st.button("Approve Pivot Table ✅"): st.session_state.pivot_approved = True

        if st.session_state.pivot_approved:
            st.markdown("### 📋 Step 2.3: Final Leave Encashment Report")
            report = pd.DataFrame({'Employee ID': df_ncp['Base_ID'].unique()})
            df_hc = standardize_id(df_hc_raw, ["user/employee id", "userid", "employee id"], "HC Report")
            df_hc['Clean_State'] = df_hc['Statutory State'].apply(lambda x: str(x).split('-')[1].strip() if '-' in str(x) else str(x).strip())
            
            report = pd.merge(report, df_hc[['Base_ID', 'Clean_State']], left_on='Employee ID', right_on='Base_ID', how='left')
            report = pd.merge(report, st.session_state.pivot_df, on='Employee ID', how='left')
            report = pd.merge(report, st.session_state.lapse_calc_df[['Base_ID', 'Total Lapse']], on='Base_ID', how='left')
            
            # --- CALCULATIONS ---
            report['CL balance'] = np.minimum(report['CL balance'].fillna(0), 0).round(3)
            report['AL balance'] = report['AL balance'].fillna(0).round(3)
            report['Leaves to be lapsed'] = report['Total Lapse'].fillna(0).round(3)
            
            report['AL to be considered for NCP adjustment'] = (report['AL balance'] + report['CL balance']).round(3)
            report['AL post lapse'] = (report['AL to be considered for NCP adjustment'] - report['Leaves to be lapsed']).round(3)
            report['Max LE days(state)'] = report['Clean_State'].map(STATE_LIMITS).fillna(0)
            
            # Final LE Metric
            report['Final LE'] = np.minimum(report['AL post lapse'], report['Max LE days(state)']).round(3)
            
            report.rename(columns={'Clean_State': 'Statutory State'}, inplace=True)
            final_cols = ['Employee ID', 'Statutory State', 'AL balance', 'CL balance', 'AL to be considered for NCP adjustment', 'Leaves to be lapsed', 'AL post lapse', 'Max LE days(state)', 'Final LE']
            
            st.session_state.final_report_df = report[final_cols].copy()
            st.dataframe(st.session_state.final_report_df, use_container_width=True)
            
            excel_data = convert_df_to_styled_excel(st.session_state.final_report_df, st.session_state.lapse_calc_df)
            st.download_button("📥 Download Final Encashment Report", data=excel_data, file_name="Final_Leave_Encashment.xlsx")

            # --- STEP 2.4: MERGE WITH CONSOLIDATED REPORT ---
            st.markdown("---")
            st.markdown("### 🔗 Step 2.4: Merge with Consolidated Report")
            l_cons = st.file_uploader("4. Upload Consolidated Report", type=["csv", "xlsx"])
            
            if l_cons:
                df_cons_raw = load_file(l_cons)
                df_cons_raw = df_cons_raw.loc[:, ~df_cons_raw.columns.str.contains('^Unnamed')]
                df_cons = standardize_id(df_cons_raw, ["employeeid", "userid", "empid", "employee id", "emp id"], "Consolidated Report")
                
                # Dynamic column check for Leave Encashment
                target_col = next((c for c in df_cons.columns if 'leave encashment' in str(c).lower()), None)
                
                if target_col is None:
                    st.warning("⚠️ 'Leave Encashment' column not found in input sheet. It will be generated automatically.")
                    target_col = 'Leave Encashment'
                
                if st.button("Perform Lookup & Generate Multi-Tab Master Sheet"):
                    calc_subset = st.session_state.final_report_df[['Employee ID', 'Final LE']]
                    
                    # Left merge to map values
                    merged_df = pd.merge(df_cons, calc_subset, left_on='Base_ID', right_on='Employee ID', how='left')
                    
                    # Explicit lookup mapping into the targeted column
                    if target_col in df_cons.columns:
                        merged_df[target_col] = merged_df['Final LE'].fillna(merged_df[target_col])
                    else:
                        merged_df[target_col] = merged_df['Final LE']
                    
                    # EXPLICIT CLEANUP: Remove standardizing metrics AND the matched 'Employee ID' column 
                    # so that it never trail-renders after the original data columns (like Remarks)
                    cols_to_drop = ['Base_ID', 'Employee ID', 'Final LE']
                    merged_df = merged_df.drop(columns=[c for c in cols_to_drop if c in merged_df.columns])
                    
                    st.success("✅ Master VLOOKUP completed successfully!")
                    st.dataframe(merged_df.head(), use_container_width=True)
                    
                    # Produces the final multi-tab workbook
                    merged_xlsx = convert_df_to_styled_excel(
                        report_df=st.session_state.final_report_df, 
                        lapse_df=st.session_state.lapse_calc_df, 
                        master_df=merged_df,
                        master_sheet_name="Consolidated Master"
                    )
                    
                    st.download_button(
                        "📥 Download Merged Consolidated Master Workbook", 
                        data=merged_xlsx, 
                        file_name="Merged_Master_Consolidated_Report.xlsx", 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

elif st.session_state.current_page == "Home":
    st.write("### 👋 Welcome to Leave Encashment Calculator")
    st.button("Launch Leave Encashment", on_click=nav_leave)
