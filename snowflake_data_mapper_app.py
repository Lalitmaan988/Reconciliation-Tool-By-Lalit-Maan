import streamlit as st
import pandas as pd
import difflib
import io

# ---------------------------------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------------------------------
st.set_page_config(
    page_title="Reconciliation Tool By Lalit Maan", 
    layout="wide", 
    page_icon="💼"
)

# Customizing the CSS for better UI and attractive footer/sidebar styling
st.markdown("""
<style>
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #f1f3f4;
        color: #5f6368;
        text-align: center;
        padding: 10px;
        font-size: 14px;
        border-top: 1px solid #e0e0e0;
        z-index: 1000;
    }
    .main-title {
        color: #1a73e8;
        font-weight: 700;
        margin-bottom: 5px;
    }
    .sub-text {
        color: #5f6368;
        font-size: 16px;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------------
# SIDEBAR / CONTACT DETAILS
# ---------------------------------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/salesforce.png", width=60)
    st.markdown("## 👨‍💻 About the Developer")
    st.info("**Built to streamline CRM data management, reduce duplicate entries, and automate complex workflows.**")
    
    st.divider()
    
    st.markdown("### 📞 Contact & Support")
    st.markdown("**Developer:** Lalit Maan")
    st.markdown("**Role:** Salesforce Admin & Data Analytics")
    st.markdown("📱 **Phone:** [+91 9001634419](tel:+919001634419)")
    st.markdown("📧 **Email:** [Lalitmaan988@gmail.com](mailto:Lalitmaan988@gmail.com)")
    st.markdown("🌐 **Portfolio:** [Data Master Jaipur](https://datamasterjaipur.blogspot.com/)")
    st.markdown("🔗 **LinkedIn:** [Lalit Maan](https://www.linkedin.com/in/lalitmaan988/)")
    st.markdown("💬 **Contact Us:** [Get in touch](https://datamasterjaipur.blogspot.com/p/contact-me.html)")
    
    st.divider()
    
    st.markdown("#### Reconciliation Tool v1.0")
    st.markdown("<small>© 2026 Lalit Maan.<br>All rights reserved.</small>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------------
# MAIN APP HEADER
# ---------------------------------------------------------------------------------
st.markdown('<h1 class="main-title">💼 Reconciliation Tool By Lalit Maan</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-text">Easily align existing database fields with new incoming data securely and efficiently.</p>', unsafe_allow_html=True)
st.divider()


# ---------------------------------------------------------------------------------
# 1. FILE UPLOAD
# ---------------------------------------------------------------------------------
st.header("Step 1 & 2: Upload Files")
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Existing Data File")
    existing_file = st.file_uploader("Upload Existing Data (CSV/Excel)", key="existing")

with col2:
    st.subheader("2️⃣ New Data File")
    new_file = st.file_uploader("Upload New Data (CSV/Excel)", key="new")

@st.cache_data
def load_data(file):
    try:
        if file.name.endswith('.csv'):
            return pd.read_csv(file)
        elif file.name.endswith('.xlsx') or file.name.endswith('.xls'):
            return pd.read_excel(file)
        else:
            st.error(f"Unsupported file format: {file.name}")
            return None
    except Exception as e:
        st.error(f"Error reading file {file.name}: {str(e)}")
        return None

if existing_file and new_file:
    df_existing = load_data(existing_file)
    df_new = load_data(new_file)
    
    if df_existing is not None and df_new is not None:
        st.success("✅ Both Files loaded successfully!")
        
        with st.expander("🔍 Preview Data", expanded=False):
            st.write("**Existing Data (First 5 Rows):**")
            st.dataframe(df_existing.head(5), use_container_width=True)
            st.write("**New Data (First 5 Rows):**")
            st.dataframe(df_new.head(5), use_container_width=True)

        # ---------------------------------------------------------------------------------
        # 2. HEADER MAPPING
        # ---------------------------------------------------------------------------------
        st.divider()
        st.header("Step 3: Header Mapping")
        st.markdown("Map headers from the **New File** to the **Existing File**. Identical names are mapped automatically. Select **Skip** if you do not want to map a column.")
        
        new_cols = list(df_new.columns)
        existing_cols = list(df_existing.columns)
        
        # Track mapped configurations in session state
        if 'col_mapping' not in st.session_state:
            st.session_state.col_mapping = {}

        with st.form("header_mapping_form"):
            st.subheader("🔗 Link Columns")
            mapping_dict = {}
            for new_col in new_cols:
                options = ["-- Skip --"] + existing_cols
                default_index = 0
                
                # Auto-matching logic
                if new_col in existing_cols:
                    default_index = options.index(new_col)
                else:
                    lower_existing = [c.lower() for c in existing_cols]
                    if new_col.lower() in lower_existing:
                        idx = lower_existing.index(new_col.lower())
                        default_index = options.index(existing_cols[idx])

                selected_col = st.selectbox(
                    f"Map New Column: '{new_col}' ➜", 
                    options=options,
                    index=default_index,
                    key=f"map_{new_col}"
                )
                mapping_dict[new_col] = selected_col
            
            mapping_submitted = st.form_submit_button("Confirm Mapping")
            if mapping_submitted:
                st.session_state.col_mapping = mapping_dict
                st.success("✅ Mapping Confirmed!")

        # ---------------------------------------------------------------------------------
        # 3. LOOKUP CONFIGURATION & PROCESSING
        # ---------------------------------------------------------------------------------
        if st.session_state.col_mapping:
            st.divider()
            st.header("Step 4: Primary Lookup Field")
            
            mapped_new_cols = {k: v for k, v in st.session_state.col_mapping.items() if v != "-- Skip --"}
            
            if not mapped_new_cols:
                st.error("No columns mapped! Please map at least one column to proceed.")
            else:
                st.markdown("Select a **Primary Key** to use for data matching (Exact & Fuzzy Lookup).")
                lookup_col_new = st.selectbox("Select Primary Lookup Column (from New File)", options=list(mapped_new_cols.keys()))
                
                if lookup_col_new:
                    lookup_col_existing = mapped_new_cols[lookup_col_new]
                    st.info(f"Using **'{lookup_col_new}'** (New) matched with **'{lookup_col_existing}'** (Existing)")
                    
                    st.markdown("---")
                    fuzzy_threshold = st.slider(
                        "Fuzzy Match Sensitivity (%)", 
                        min_value=10, 
                        max_value=100, 
                        value=85, 
                        step=5, 
                        help="Controls how strictly the text must match to be considered a fuzzy match. 100% means exact match required."
                    )
                    st.markdown("---")
                    
                    match_button = st.button("🚀 Run Data Reconciliation", type="primary", use_container_width=True)
                    
                    if match_button:
                        with st.spinner("Matching Data... Applying Exact and Fuzzy lookups..."):
                            
                            df_new_temp = df_new.copy()
                            df_existing_temp = df_existing.copy()
                            
                            df_new_temp['_match_key'] = df_new_temp[lookup_col_new].astype(str).str.strip().str.lower()
                            df_existing_temp['_match_key'] = df_existing_temp[lookup_col_existing].astype(str).str.strip().str.lower()
                            
                            existing_cols_to_add = [c for c in df_existing.columns if c != lookup_col_existing]
                            
                            # Initialize appended columns
                            for c in existing_cols_to_add:
                                df_new_temp[f"Existing_{c}"] = ""
                                
                            df_new_temp['Match_Reason'] = "No Match"
                            
                            existing_keys = df_existing_temp['_match_key'].dropna().unique().tolist()
                            unique_new_keys = df_new_temp['_match_key'].dropna().unique()
                            
                            match_dict = {} 
                            
                            # Matching Engine
                            for val in unique_new_keys:
                                if pd.isna(val) or val == "" or val == "nan":
                                    continue
                                
                                if val in existing_keys:
                                    match_dict[val] = (val, "Exact Match")
                                else:
                                    matches = difflib.get_close_matches(val, existing_keys, n=1, cutoff=(fuzzy_threshold/100.0))
                                    if matches:
                                        matched_val = matches[0]
                                        score = round(difflib.SequenceMatcher(None, val, matched_val).ratio() * 100, 2)
                                        match_dict[val] = (matched_val, f"Fuzzy Match Score: {score}%")
                                    else:
                                        match_dict[val] = (None, "No Match")
                            
                            df_existing_indexed = df_existing_temp.set_index('_match_key')
                            
                            for idx, row in df_new_temp.iterrows():
                                n_key = row['_match_key']
                                if n_key in match_dict:
                                    matched_key, reason = match_dict[n_key]
                                    df_new_temp.at[idx, 'Match_Reason'] = reason
                                    
                                    if matched_key is not None:
                                        matched_row = df_existing_indexed.loc[matched_key]
                                        if isinstance(matched_row, pd.DataFrame):
                                            matched_row = matched_row.iloc[0]
                                            
                                        for c in existing_cols_to_add:
                                            df_new_temp.at[idx, f"Existing_{c}"] = matched_row[c]
                            
                            df_new_temp = df_new_temp.drop(columns=['_match_key'])
                            st.session_state.matched_result = df_new_temp
                            st.success("✅ Reconciliation Complete!")
                            
        # ---------------------------------------------------------------------------------
        # 5. OUTPUT PREVIEW & DOWNLOAD (EXCEL)
        # ---------------------------------------------------------------------------------
        if 'matched_result' in st.session_state:
            st.divider()
            st.header("Step 5: Mapped File Output")
            result_df = st.session_state.matched_result
            
            st.subheader("📊 Match Statistics")
            match_stats = result_df['Match_Reason'].apply(lambda x: "Fuzzy Match" if "Fuzzy" in x else x).value_counts()
            
            stat_cols = st.columns(len(match_stats))
            for i, (reason, count) in enumerate(match_stats.items()):
                stat_cols[i].metric(reason, f"{count} Rows")
            
            st.markdown("### Combined Data Result")
            st.dataframe(result_df, use_container_width=True)
            
            st.markdown("---")
            
            # EXCEL DOWNLOAD LOGIC
            output = io.BytesIO()
            # We use 'openpyxl' engine as it is the default modern excel writer for pandas
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Reconciled_Data')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Reconciled Data (Excel .xlsx)",
                data=excel_data,
                file_name='Reconciled_Output.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type="primary"
            )

# FIXED FOOTER
st.markdown("""
<div class="footer">
    <strong>Reconciliation Tool v1.0</strong> | © 2026 Lalit Maan. All rights reserved.
</div>
""", unsafe_allow_html=True)
