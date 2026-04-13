import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Universal MIS Tool", page_icon="📊", layout="wide")

st.title("📊 Universal Monthly Variance Analyzer")
st.markdown("Automated comparison for multi-column Tally reports.")

uploaded_file = st.file_uploader("Upload Trial Balance", type=["csv", "xlsx"])

if uploaded_file:
    try:
        # 1. LOAD DATA
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None).fillna("")
        else:
            df_raw = pd.read_excel(uploaded_file, header=None).fillna("")

        # 2. FIND THE HEADER ROW (Search for 'Particulars')
        header_idx = None
        for i in range(len(df_raw)):
            row_values = [str(val).strip() for val in df_raw.iloc[i].values]
            if "Particulars" in row_values:
                header_idx = i
                break

        if header_idx is None:
            st.error("Could not find 'Particulars' column.")
            st.stop()

        # 3. ADVANCED MONTH HUNTER
        # We look at everything above the header row to find Month names
        month_names = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
        
        # We create a map of which columns belong to which month
        col_to_month = {}
        current_detected_month = "Base Period"
        
        # Scan rows above the header
        for r in range(max(0, header_idx-5), header_idx):
            row = df_raw.iloc[r].tolist()
            for c_idx, val in enumerate(row):
                val_str = str(val).strip().lower()
                if any(m in val_str for m in month_names):
                    # Found a month!
                    col_to_month[c_idx] = str(val).strip()

        # 4. BUILD COLUMN NAMES
        sub_headers = [str(s).strip() for s in df_raw.iloc[header_idx].tolist()]
        final_columns = []
        
        # Track the last month found to "fill" it forward across Debit/Credit/Balance
        last_m = "Base"
        for i, s in enumerate(sub_headers):
            if i in col_to_month:
                last_m = col_to_month[i]
            
            if "Particulars" in s:
                final_columns.append(f"Ledger_Name_{i}")
            else:
                # Combine Month and the sub-header (Debit/Credit/Balance)
                final_columns.append(f"{last_m}|{s}|{i}")

        # 5. PREPARE DATA
        df = df_raw.iloc[header_idx + 1:].copy()
        df.columns = final_columns
        
        # 6. GROUP BY MONTH FOR SIDEBAR
        all_detected_months = sorted(list(set([c.split('|')[0] for c in final_columns if '|' in c])))
        
        st.sidebar.header("Select Comparison")
        m1_choice = st.sidebar.selectbox("Base Month", all_detected_months, index=0)
        m2_choice = st.sidebar.selectbox("Current Month", all_detected_months, index=len(all_detected_months)-1)

        # 7. CALCULATION ENGINE
        def get_balance(month_name, data):
            relevant = [c for c in data.columns if c.startswith(month_name)]
            
            def clean(x):
                try:
                    s = str(x).replace(' Dr', '').replace(' Cr', '').replace(',', '').strip()
                    return float(s) if s else 0.0
                except: return 0.0

            # 1. Try to find a 'Balance' or 'Closing' column first
            bal_col = [c for c in relevant if 'Balance' in c or 'Closing' in c]
            if bal_col:
                return data[bal_col[0]].apply(clean)
            
            # 2. If not, calculate Net (Debit - Credit)
            dr_col = [c for c in relevant if 'Debit' in c]
            cr_col = [c for c in relevant if 'Credit' in c]
            
            dr_vals = data[dr_col[0]].apply(clean) if dr_col else 0.0
            cr_vals = data[cr_col[0]].apply(clean) if cr_col else 0.0
            return dr_vals - cr_vals

        # 8. GENERATE FINAL REPORT
        p_col = [c for c in df.columns if 'Ledger_Name' in c][0]
        report = pd.DataFrame()
        report['Particulars'] = df[p_col]
        report[m1_choice] = get_balance(m1_choice, df)
        report[m2_choice] = get_balance(m2_choice, df)
        
        report['Variance'] = report[m2_choice] - report[m1_choice]
        report['% Change'] = (report['Variance'] / report[m1_choice].replace(0, 1))

        # 9. DISPLAY
        st.subheader(f"Results: {m1_choice} vs {m2_choice}")
        st.dataframe(report.style.format({m1_choice: "{:,.2f}", m2_choice: "{:,.2f}", "Variance": "{:,.2f}", "% Change": "{:.1%}"}), use_container_width=True)

        # 10. DOWNLOAD
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            report.to_excel(writer, sheet_name='MIS_Report', index=False)
            wb, ws = writer.book, writer.sheets['MIS_Report']
            
            fmt_hdr = wb.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1})
            fmt_num = wb.add_format({'num_format': '#,##0.00', 'border': 1})
            fmt_pct = wb.add_format({'num_format': '0.0%', 'border': 1})

            ws.set_column('A:A', 40)
            ws.set_column('B:D', 18, fmt_num)
            ws.set_column('E:E', 12, fmt_pct)
            for i, col in enumerate(report.columns):
                ws.write(0, i, col, fmt_hdr)

        st.download_button("📥 Download Official MIS Report", output.getvalue(), "MIS_Analysis.xlsx")

    except Exception as e:
        st.error(f"Something went wrong: {e}")
