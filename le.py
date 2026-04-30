import streamlit as st
import pandas as pd
import numpy as np
import time
import re
import io

# ==========================================
# 🛠️ UNIVERSAL COLUMN MAPPER (SF SPECIFIC)
# ==========================================
COLUMN_MAPPER = {
    'DOJ': ['employment details date of joining', 'date of joining', 'joining date', 'doj'],
    'Legal entity DOJ': ['employment details legal entity date of joining', 'legal entity date of joining', 'legal entity doj'],
    'Group DOJ': ['employment details group date of joining', 'group date of joining', 'group doj'],
    'Statutory State': ['statutory state', 'state', 'location state', 'base state'],
    'Event Reason': ['event reason', 'reason for exit', 'exit reason', 'separation reason'],
    'Employment Details Actual Exit Date': ['employment details actual exit date', 'actual exit date', 'exit date', 'lwd', 'last working day', 'dol'],
    'Employment Details Date of Resignation': ['employment details date of resignation', 'date of resignation', 'resignation date'],
    'Department': ['department', 'dept', 'business unit', 'function', 'department (label)'],
    'Position Title': ['position title', 'job title', 'designation'],
    'Employee Type': ['employee type', 'worker type'],
    'Piece Rate': ['piece rate/ non-piece rate', 'piece rate', 'worker type']
}

def translate_columns(df):
    if df is None or df.empty: return df
    clean_cols = {c: str(c).strip().lower() for c in df.columns}
    rename_dict = {}
    mapped_standards = set() 
    for standard_name, possible_names in COLUMN_MAPPER.items():
        for actual_col, clean_col in clean_cols.items():
            if clean_col in possible_names and standard_name not in mapped_standards:
                rename_dict[actual_col] = standard_name
                mapped_standards.add(standard_name)
                break 
    return df.rename(columns=rename_dict)

# --- PAGE CONFIGURATION & PREMIUM CSS ---
st.set_page_config(page_title="Leave Encashment Calculator", page_icon="🌴", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #F4F6F9; }
    @keyframes gradientBG {
        0% {background-position: 0% 50%;} 50% {background-position: 100% 50%;} 100% {background-position: 0% 50%;}
    }
    .main-header {
        background: linear-gradient(-45deg, #5E239D, #9D4EDD, #FF007F, #5E239D);
        background-size: 300% 300%; animation: gradientBG 8s ease infinite;
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        text-align: center; font-size: 3.5rem; font-weight: 900; margin-bottom: -15px; padding-top: 10px;
    }
    .sub-header { text-align: center; color: #4B5563; font-size: 1.2rem; font-weight: 600; margin-bottom: 30px; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E5E7EB; box-shadow: 2px 0 15px rgba(0,0,0,0.05); }
    
    div.stButton > button {
        border-radius: 10px !important; font-weight: 700 !important; 
        background-color: #5E239D !important; color: white !important;
        border: none !important; transition: all 0.3s ease !important;
    }
    div.stButton > button:hover {
        background-color: #4A1B7D !important; transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(94, 35, 157, 0.3) !important;
    }
    [data-testid="stMetricValue"] { color: #5E239D !important; font-weight: 900 !important; font-size: 2.2rem !important; }
    .alert-popup {
        background: linear-gradient(135deg, #FFF1F2 0%, #FFE4E6 100%);
        border: 2px solid #F43F5E; border-radius: 12px; padding: 25px; margin: 20px 0;
        text-align: center; animation: pulse-red 2s infinite;
    }
    @keyframes pulse-red {
        0% { box-shadow: 0 0 0 0 rgba(244, 63, 94, 0.4); }
        70% { box-shadow: 0 0 0 15px rgba(244, 63, 94, 0); }
        100% { box-shadow: 0 0 0 0 rgba(244, 63, 94, 0); }
    }
    .feature-card {
        background: white; padding: 25px; border-radius: 12px; border-top: 4px solid #5E239D;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); transition: transform 0.3s ease; height: 100%;
    }
    </style>
""", unsafe_allow_html=True)

# --- SESSION STATE ---
if 'current_page' not in st.session_state: st.session_state.current_page = "Home"
states = ['factor_preview_ready', 'factor_calc_done', 'reports_processed', 'edits_confirmed', 'validation_run']
for s in states:
    if s not in st.session_state: st.session_state[s] = False
if 'absconding_decision' not in st.session_state: st.session_state.absconding_decision = 'pending'
if 'analytics_df' not in st.session_state: st.session_state.analytics_df = None

def nav_home(): st.session_state.current_page = "Home"
def nav_factor(): st.session_state.current_page = "Factor"
def nav_master(): st.session_state.current_page = "Master"
def nav_analytics(): st.session_state.current_page = "Analytics"
def nav_validations(): st.session_state.current_page = "Validations"

def exclude_absconding():
    mask = st.session_state.final_master_df['Event Reason'].astype(str).str.contains('33|absconding', case=False, na=False)
    st.session_state.final_master_df = st.session_state.final_master_df[~mask]
    st.session_state.absconding_decision = 'exclude'
    st.session_state.edits_confirmed = True

def include_absconding():
    st.session_state.absconding_decision = 'include'

# --- ROBUST FILE LOADER ---
def load_file(uploaded_file):
    uploaded_file.seek(0)
    if uploaded_file.name.endswith('.csv'):
        encodings = ['utf-8-sig', 'latin1', 'cp1252']
        for enc in encodings:
            for separator in [',', '\t']:
                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, sep=separator, encoding=enc, on_bad_lines='skip')
                    if df.shape[1] > 2: return df
                except: continue
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, engine='python', on_bad_lines='skip')
    else: 
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file)

# --- SMART ID & HEADER LOCATOR ---
def standardize_id(df, possible_names, report_name):
    def scrub(val): return re.sub(r'[^a-zA-Z0-9]', '', str(val)).lower()
    clean_possible = [scrub(p) for p in possible_names]
    
    def find_indices(headers_list):
        headers_scrubbed = [scrub(h) for h in headers_list]
        return [i for i, h in enumerate(headers_scrubbed) if h in clean_possible]

    # Check top rows for header (handles Row 1 vs Row 3)
    id_indices = find_indices(df.columns)
    if not id_indices or len(df.columns) < 2:
        for i, row in df.head(50).iterrows():
            found_indices = find_indices(row.values)
            if found_indices:
                df.columns = [str(v).strip() for v in row.values]
                df = df.iloc[i+1:].reset_index(drop=True)
                id_indices = found_indices
                break
                
    if id_indices:
        base_id_series = pd.Series(np.nan, index=df.index)
        cols_to_drop = []
        for idx in id_indices:
            orig_name = df.columns[idx]
            if orig_name != 'Base_ID': cols_to_drop.append(orig_name)
            temp_col = df.iloc[:, idx].astype(str).str.strip().replace(r'(?i)^(nan|none|)$', np.nan, regex=True)
            base_id_series = base_id_series.fillna(temp_col)
        
        df = df.copy()
        df['Base_ID'] = base_id_series
        df = df.drop(columns=[c for c in cols_to_drop if c in df.columns])
        df = df.dropna(subset=['Base_ID']).reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]
        return df.loc[:, ~df.columns.duplicated()]
    else:
        st.error(f"🚨 Error: Could not locate the ID column in the **{report_name}**. Please ensure the column header is named 'Employee ID' or similar."); st.stop()

def merge_clean(base_df, new_df):
    cols_to_use = ['Base_ID'] + [c for c in new_df.columns if c not in base_df.columns and c != 'Base_ID']
    return base_df.merge(new_df[cols_to_use], on='Base_ID', how='left')

# --- MAIN UI ---
st.markdown("<div class='main-header'>Leave Encashment Calculator</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>Automated Settlement Calculation & Consolidation</div>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h2 style='color:#5E239D; text-align:center;'>🧭 Workspace</h2>", unsafe_allow_html=True)
    st.button("📊 1. Factor Calculator", on_click=nav_factor, use_container_width=True)
    st.button("📂 2. Master Builder", on_click=nav_master, use_container_width=True)
    
    # NEW: External Link Button
    st.link_button("🌴 3. Leave Encashment Calcy", 
                   "https://leave-encashment-calcy-fpt6j8bbzgdenee2jigjgl.streamlit.app/", 
                   use_container_width=True)
    
    st.button("📈 4. Exit Analytics", on_click=nav_analytics, use_container_width=True)
    st.button("✅ 5. Validations", on_click=nav_validations, use_container_width=True)

if st.session_state.current_page != "Home" and st.session_state.current_page != "Validations":
    st.button("🏠 Back to Home", on_click=nav_home)

# ==========================================
# MODULE 1: FACTOR DATA CALCULATION 
# ==========================================
if st.session_state.current_page == "Factor":
    st.markdown("### 📊 Step 1: Factor Data Calculator")
    
    with st.container(border=True):
        st.write("Upload your **CTC Report** and **Raw Factor Data**. We will extract the data and let you preview/edit it before calculating.")
        
        col1, col2 = st.columns(2)
        with col1:
            ctc_file = st.file_uploader("1. Upload CTC Report", type=["xlsx", "csv"], key="ctc_upload")
        with col2:
            raw_factor_files = st.file_uploader("2. Raw Factor Data (Can select multiple)", type=["xlsx", "csv"], key="raw_factor_upload", accept_multiple_files=True)

    # --- PHASE 1: EXTRACT & PREVIEW ---
    if ctc_file and raw_factor_files:
        col_calc1, col_calc2, col_calc3 = st.columns([1, 2, 1])
        with col_calc2:
            if st.button("🔍 Extract & Preview Inputs", use_container_width=True):
                with st.spinner("Extracting and Merging..."):
                    df_ctc_raw = load_file(ctc_file)
                    df_ctc = standardize_id(df_ctc_raw, ['Employee Code', 'Emp ID', 'Employee ID'], "CTC Report")
                    
                    ctc_rename_map = {'Basic Pay Sales Master': 'Basic Pay', 'HRA Sales Master': 'HRA Sales'}
                    for old_col, new_col in ctc_rename_map.items():
                        if old_col in df_ctc.columns: df_ctc.rename(columns={old_col: new_col}, inplace=True)
                    
                    ctc_req_cols = ['Base_ID', 'Basic Pay', 'HRA Sales', 'Consistency Allowance', 'Adv Stt Bonus SalesMaster', 'Sales Linked Comm. Master', 'Mobile Allow Sales Master']
                    for c in ctc_req_cols:
                        if c not in df_ctc.columns: df_ctc[c] = 0
                    df_ctc_clean = df_ctc[ctc_req_cols]

                    def get_factor_values(df, factor_letter):
                        clean_cols = [str(c).strip().lower() for c in df.columns]
                        valid_names = [f'part {factor_letter}', f'part {factor_letter} factor', f'part{factor_letter}']
                        idx = [i for i, c in enumerate(clean_cols) if c in valid_names]
                        if idx: return pd.to_numeric(df.iloc[:, idx[0]], errors='coerce').fillna(0).values
                        return np.zeros(len(df))

                    all_factor_dfs = []
                    for file in raw_factor_files:
                        df_factor_raw = load_file(file)
                        df_factor = standardize_id(df_factor_raw, ['EMP Code/BT ID', 'Emp ID', 'Employee Code', 'Employee ID'], f"Factor Report ({file.name})")
                        all_factor_dfs.append(pd.DataFrame({'Base_ID': df_factor['Base_ID'].values, 'Part A': get_factor_values(df_factor, 'a'), 'Part B': get_factor_values(df_factor, 'b'), 'Part C': get_factor_values(df_factor, 'c')}))

                    df_merged = pd.concat(all_factor_dfs, ignore_index=True).drop_duplicates(subset=['Base_ID'], keep='last').merge(df_ctc_clean, on='Base_ID', how='left').fillna(0)
                    df_merged.rename(columns={'Base_ID': 'Emp ID'}, inplace=True)
                    st.session_state.raw_factor_merged = df_merged
                    st.session_state.factor_preview_ready = True
                    st.session_state.factor_calc_done = False
                    st.toast("Data extracted successfully! Ready for review.", icon="✅")

    # --- PHASE 2: EDITING GRID & CALCULATE ---
    if st.session_state.factor_preview_ready:
        st.write("---")
        st.markdown("### 📝 Pre-Calculation Editor")
        edited_raw_data = st.data_editor(st.session_state.raw_factor_merged, key="factor_pre_calc_editor", use_container_width=True, num_rows="dynamic")
        
        col_conf1, col_conf2, col_conf3 = st.columns([1,2,1])
        with col_conf2:
            if st.button("✔️ Confirm Data & Run Final Calculations", use_container_width=True):
                with st.spinner("Calculating..."):
                    df = edited_raw_data.copy()
                    numeric_cols = ['Basic Pay', 'HRA Sales', 'Consistency Allowance', 'Adv Stt Bonus SalesMaster', 'Sales Linked Comm. Master', 'Mobile Allow Sales Master', 'Part A', 'Part B', 'Part C']
                    for col in numeric_cols: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    
                    df['A. Basic Pay'] = df['Basic Pay'] / 12
                    df['Final Basic pay'] = (df['Part A'] * df['A. Basic Pay']).round(0)
                    df['Final HRA'] = np.where(df['HRA Sales'] > 0, (0.05 * (df['A. Basic Pay'] + (df['Consistency Allowance']/12))).round(0), 0)
                    df['Final Const. Bonus'] = (df['Part B'] * (df['Consistency Allowance']/12)).round(0)
                    df['Final Sales linked'] = (df['Part C'] * (df['Sales Linked Comm. Master']/12)).round(0)
                    df['Final Mobile allowance'] = (df['Part A'] * (df['Mobile Allow Sales Master']/12)).round(0)
                    df['Final Adv. stat bonus'] = np.where(df['Adv Stt Bonus SalesMaster'] > 0, (0.0833 * (df['Final Basic pay'] + df['Final Const. Bonus'])).round(0), 0)

                    final_columns = ['Emp ID', 'Final Basic pay', 'Final HRA', 'Final Const. Bonus', 'Final Adv. stat bonus', 'Final Sales linked', 'Final Mobile allowance']
                    df_final = df[final_columns].copy()
                    for col in final_columns[1:]: df_final[col] = df_final[col].fillna(0).astype(int)
                    st.session_state.calculated_factor_df = df_final
                    st.session_state.factor_calc_done = True
                    st.rerun()

    # --- PHASE 3: DASHBOARD & DOWNLOAD ---
    if st.session_state.factor_calc_done and st.session_state.calculated_factor_df is not None:
        st.write("---")
        st.markdown("#### 📈 Output Dashboard")
        dcol1, dcol2, dcol3, dcol4 = st.columns(4)
        df_final = st.session_state.calculated_factor_df
        dcol1.metric("Total Employees", f"{len(df_final)}")
        dcol2.metric("Total Basic Payout", f"₹ {df_final['Final Basic pay'].sum():,.0f}")
        dcol3.metric("Total HRA Payout", f"₹ {df_final['Final HRA'].sum():,.0f}")
        dcol4.metric("Total Const. Bonus", f"₹ {df_final['Final Const. Bonus'].sum():,.0f}")
        with st.expander("👀 View Final Calculated Payout Data", expanded=True):
            st.dataframe(df_final.head(15), use_container_width=True)
        st.download_button(label="📥 Download Calculated Factor Data", data=df_final.to_csv(index=False), file_name="Calculated_Factor_Data.csv", mime="text/csv", use_container_width=True)

# ==========================================
# MODULE 2: MASTER BUILDER
# ==========================================
elif st.session_state.current_page == "Master":
    st.markdown("### 📂 Step 2: Master Report Builder")
    col1, col2 = st.columns(2)
    with col1:
        f_fact = st.file_uploader("1. Calculated Factor Data", type=["csv", "xlsx"])
        f_hc = st.file_uploader("2. HC Report", type=["csv", "xlsx"])
        f_exit = st.file_uploader("3. Exit Report", type=["csv", "xlsx"])
    with col2:
        f_ffs = st.file_uploader("4. FFS Input Report", type=["csv", "xlsx"])
        f_inv = st.file_uploader("5. Inventory Tracker", type=["csv", "xlsx"])

    if all([f_fact, f_hc, f_exit, f_ffs, f_inv]):
        if st.button("🚀 Run Consolidation Engine", use_container_width=True):
            with st.spinner("Consolidating..."):
                id_search = ["employeeuserid", "useremployeeid", "employeeid", "userid", "empid", "personid"]
                df_fact = standardize_id(load_file(f_fact), id_search, "Factor Data")
                df_hc_raw = standardize_id(load_file(f_hc), id_search, "HC Report")
                df_exit = standardize_id(load_file(f_exit), id_search, "Exit Report")
                df_ffs = standardize_id(load_file(f_ffs), id_search, "FFS Input Report")
                df_inv = standardize_id(load_file(f_inv), id_search, "Inventory Tracker")

                doj_targets = ['employmentdetailsdateofjoining', 'employmentdetailslegalentitydateofjoining', 'employmentdetailsgroupdateofjoining', 'dateofjoining', 'groupdateofjoining', 'legalentitydateofjoining']
                valid_dojs = []
                for col in df_hc_raw.columns:
                    if re.sub(r'[^a-zA-Z0-9]', '', str(col)).lower() in doj_targets:
                        df_hc_raw[col] = pd.to_datetime(df_hc_raw[col], errors='coerce')
                        valid_dojs.append(col)
                if valid_dojs: df_hc_raw['Minimum Date of Joining'] = df_hc_raw[valid_dojs].min(axis=1)

                state_col = next((c for c in df_hc_raw.columns if re.sub(r'[^a-zA-Z0-9]', '', str(c)).lower() == 'statutorystate'), None)
                if state_col: df_hc_raw['Statutory State'] = df_hc_raw[state_col].astype(str).str.split('-').str[-1].str.strip()

                if 'First Name' not in df_hc_raw.columns and 'Employee Full Name' in df_hc_raw.columns:
                    df_hc_raw['First Name'] = df_hc_raw['Employee Full Name'].astype(str).str.split().str[0]
                    df_hc_raw['Last Name'] = df_hc_raw['Employee Full Name'].astype(str).str.split().str[-1]

                df_hc, df_exit, df_ffs, df_inv = translate_columns(df_hc_raw), translate_columns(df_exit), translate_columns(df_ffs), translate_columns(df_inv)
                merged = df_fact.copy()
                for d in [df_hc, df_exit, df_ffs, df_inv]: merged = merge_clean(merged, d)
                
                merged.rename(columns={'Base_ID': 'Emp ID'}, inplace=True)
                final_req_cols = ['Emp ID', 'Final Basic pay', 'Final HRA', 'Final Const. Bonus', 'Final Adv. stat bonus', 'Final Sales linked', 'Final Mobile allowance', 'First Name', 'Last Name', 'Employee Type', 'Minimum Date of Joining', 'Position Title', 'Statutory State', 'Event Reason', 'Employment Details Date of Resignation', 'Employment Details Actual Exit Date', 'Field Allowance Payable', 'Joining Bonus Recovery', 'Notice Period Buyout Amount Recovery', 'Notice Period days to be Recovered', 'Payment in lieu of Notice Period (In Days)', 'Retention Bonus Recovery', 'ED : Handset Allowance to be Recovered', 'FC : Financial Recovery Amount', 'FC : Financial Recovery Reason', 'SC : Security Clearance', 'SC : Financial Recovery Amount', 'ITAC : Final Recovery', 'Inventory Recovery Input']
                for col in final_req_cols:
                    if col not in merged.columns: merged[col] = 0
                    else: merged[col] = merged[col].fillna(0)
                if 'FC : Financial Recovery Reason' in merged.columns:
                    merged['FC : Financial Recovery Reason'] = merged['FC : Financial Recovery Reason'].astype(str).replace(to_replace=r'(?i)^(nil|nill|none|nan|0|0.0)$', value='0', regex=True)

                st.session_state.final_master_df = merged[final_req_cols].loc[:, ~merged[final_req_cols].columns.duplicated()]
                st.session_state.reports_processed = True; st.rerun()

    if st.session_state.reports_processed:
        curr = st.session_state.final_master_df
        abs_mask = curr['Event Reason'].astype(str).str.contains('33|absconding', case=False, na=False)
        abs_count = len(curr[abs_mask])
        if abs_count > 0 and st.session_state.absconding_decision == 'pending':
            st.markdown(f"<div class='alert-popup'><h2>⚠️ {abs_count} Absconding Cases detected!</h2></div>", unsafe_allow_html=True)
            ac1, ac2 = st.columns(2)
            with ac1: st.button("✂️ Exclude", on_click=exclude_absconding, use_container_width=True)
            with ac2: st.button("📝 Review", on_click=include_absconding, use_container_width=True)
        elif st.session_state.absconding_decision == 'include' and not st.session_state.edits_confirmed:
            edited = st.data_editor(curr[abs_mask], use_container_width=True, key="abs_ed")
            if st.button("✔️ Save Edits"):
                curr.update(edited); st.session_state.final_master_df = curr; st.session_state.edits_confirmed = True; st.rerun()
        else:
            st.dataframe(curr.head(20), use_container_width=True)
            st.download_button("📥 Download Master csv", data=curr.to_csv(index=False), file_name="Master_FnF_Consolidated.csv", use_container_width=True)

# ==========================================
# MODULE 3: EXIT ANALYTICS
# ==========================================
elif st.session_state.current_page == "Analytics":
    st.markdown("### 📈 Exit Analytics Dashboard")
    if st.session_state.analytics_df is None:
        tracker_file = st.file_uploader("Upload Tracker", type=["xlsx", "csv"])
        if tracker_file:
            xls = pd.ExcelFile(tracker_file)
            target = next((s for s in xls.sheet_names if 'Master' in s), xls.sheet_names[0])
            st.session_state.analytics_df = translate_columns(pd.read_excel(xls, sheet_name=target)); st.rerun()
    else:
        df_f = st.session_state.analytics_df
        st.button("🔄 Reset Analytics", on_click=lambda: st.session_state.update({'analytics_df': None}))
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Piece Rate Breakdown**")
            pr_col = next((c for c in df_f.columns if re.sub(r'[^a-zA-Z0-9]', '', str(c)).lower() == 'pieceratenonpiecerate'), None)
            if pr_col: st.bar_chart(df_f[pr_col].astype(str).replace('nan', 'Unspecified').value_counts(), color="#9D4EDD")
        with c2:
            st.markdown("**Exits by Reason**")
            if 'Event Reason' in df_f.columns: st.bar_chart(df_f['Event Reason'].astype(str).value_counts().head(5), color="#FF007F")
        st.write("---")
        st.markdown("#### 📞 Recovery Funnel")
        def cnt(col): return len(df_f[~df_f[col].astype(str).replace(r'(?i)^(nan|none|)$', '', regex=True).eq('')]) if col in df_f.columns else 0
        mails = pd.DataFrame({"Stage": ["Mail 1", "Mail 2", "Mail 3"], "Count": [cnt('Reminder Email 1'), cnt('Reminder Email 2'), cnt('Reminder Email 3')]}).set_index("Stage")
        calls = pd.DataFrame({"Stage": ["Call 1", "Call 2", "Call 3"], "Count": [cnt('Follow Up Call -1'), cnt('Follow Up Call -2'), cnt('Follow Up Call -3')]}).set_index("Stage")
        sc1, sc2 = st.columns(2)
        with sc1: st.bar_chart(mails, color="#5E239D")
        with sc2: st.bar_chart(calls, color="#F43F5E")
        with st.expander("👀 View Tracker"): st.dataframe(df_f.head(100), use_container_width=True)

# ==========================================
# MODULE 4: VALIDATIONS
# ==========================================
elif st.session_state.current_page == "Validations":
    st.markdown("### ✅ Step 4: Report Validations")
    if not st.session_state.validation_run:
        v1, v2, v3 = st.columns(3)
        with v1: f1 = st.file_uploader("File 1", key="v1")
        with v2: f2 = st.file_uploader("File 2", key="v2")
        with v3: f3 = st.file_uploader("File 3", key="v3")
        valid = {k: v for k, v in {'File 1': f1, 'File 2': f2, 'File 3': f3}.items() if v is not None}
        if len(valid) >= 2:
            if st.button("🔍 Run Audit", use_container_width=True):
                dfs = {n: standardize_id(load_file(f), ["employeeid", "userid", "empid"], n).set_index('Base_ID') for n, f in valid.items()}
                ids = set.union(*[set(d.index) for d in dfs.values()])
                common_cols = set.intersection(*[set(d.columns) for d in dfs.values()])
                mismatches = []
                for eid in ids:
                    for col in common_cols:
                        v = [str(dfs[n].at[eid, col]).strip().lower().replace('.0', '') if eid in dfs[n].index else "MISSING" for n in dfs]
                        if len(set(v)) > 1:
                            rec = {"Emp ID": eid, "Column": col}
                            rec.update({n: str(dfs[n].at[eid, col]) if eid in dfs[n].index else "N/A" for n in dfs}); mismatches.append(rec)
                st.session_state.mismatches = mismatches; st.session_state.validation_run = True; st.rerun()
    else:
        if not st.session_state.mismatches: st.success("🎉 No Mismatches Found!")
        else: st.dataframe(pd.DataFrame(st.session_state.mismatches), use_container_width=True)
        st.button("🔄 New Audit", on_click=lambda: st.session_state.update({'validation_run': False}))

# ==========================================
# HOME PAGE
# ==========================================
if st.session_state.current_page == "Home":
    st.write("### 👋 Welcome to Settlement Portal")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("<div class='feature-card'><h3>📊 Calculator</h3>Manage CTC/Factors.</div>", unsafe_allow_html=True)
        st.button("Launch Calculator", on_click=nav_factor, use_container_width=True)
    with c2:
        st.markdown("<div class='feature-card'><h3>📂 Consolidator</h3>Build reports.</div>", unsafe_allow_html=True)
        st.button("Launch Consolidator", on_click=nav_master, use_container_width=True)
