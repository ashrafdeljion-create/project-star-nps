import os
import re
import io
import pandas as pd
import pyreadstat
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
import streamlit as st

st.set_page_config(page_title="Project Star NPS Generator", layout="centered")

st.title("⭐ Project Star: NPS Dashboard & Streamlined Data Generator")
st.markdown("Upload your master SPSS (`.sav`) data file below, select your wave preferences, and click **Run Processing** to generate your reports.")

# --- File Uploader ---
uploaded_file = st.file_uploader("Upload Master SPSS Data File (.sav)", type=["sav"])

# --- Wave Filter Configuration ---
filter_option = st.radio("Select Wave Filter Option:", ["All Waves", "Custom Range (e.g., Wave 1 to 10)", "Specific Waves List"])

selected_waves_filter = 'ALL'

if filter_option == "Custom Range (e.g., Wave 1 to 10)":
    col1, col2 = st.columns(2)
    with col1:
        start_w = st.number_input("Start Wave Number", min_value=1, max_value=30, value=1)
    with col2:
        end_w = st.number_input("End Wave Number", min_value=1, max_value=30, value=10)
    selected_waves_filter = [f'Wave {i}' for i in range(int(start_w), int(end_w) + 1)]

elif filter_option == "Specific Waves List":
    waves_input = st.text_input("Enter waves separated by commas:", "Wave 20, Wave 21, Wave 22")
    selected_waves_filter = [w.strip() for w in waves_input.split(',')]

if st.button("🚀 Run Processing & Generate Reports", type="primary"):
    if uploaded_file is None:
        st.error("Please upload a `.sav` file first!")
    else:
        with st.spinner("Processing data and building formatted Excel dashboard..."):
            # Save uploaded file temporarily to process with pyreadstat
            temp_src_path = "temp_input.sav"
            with open(temp_src_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            dest_filename = "Overall NPS Rating per BM RM Portfolio.xlsx"
            sav_output_filename = "Project Star_NPS_Streamlined.sav"

            # 1. Load data
            df_raw, meta = pyreadstat.read_sav(temp_src_path, apply_value_formats=False)
            df_raw.columns = [str(col).strip().upper() for col in df_raw.columns]

            df_lbl, _ = pyreadstat.read_sav(temp_src_path, apply_value_formats=True)
            df_lbl.columns = [str(col).strip().upper() for col in df_lbl.columns]

            # 2. Column Mapping & Subsetting
            target_columns_mapping = {
                'UNIQUEID': ['UNIQUEID', 'ID'],
                'RUID': ['RUID'],
                'TYPE': ['TYPE'],
                'WAVE': ['WAVE'],
                'REGION': ['REGION'],
                'SUBREG': ['SUBREG'],
                'SEGMENT': ['SEGMENT'],
                'RM_NPS1': ['RM_NPS1', 'BM_NPS1', 'BMNPS01'],
                'FNB_NPS1': ['FNB_NPS1', 'FNBNPS01'],
                'PRIM_OFCR_IND': ['PRIM_OFCR_IND', 'PRIM_OFC'],
                'OFFICER_NAME': ['OFFICER_NAME', 'OFFICER_M', 'OFFICER_NAME_']
            }

            upper_to_orig = {str(c).strip().upper(): c for c in df_raw.columns}
            rename_map = {}
            for target, candidates in target_columns_mapping.items():
                for cand in candidates:
                    if cand.upper() in upper_to_orig:
                        rename_map[upper_to_orig[cand.upper()]] = target
                        break

            df_data = pd.DataFrame()
            for orig_col, target in rename_map.items():
                if target in ['TYPE', 'WAVE', 'REGION', 'SUBREG', 'SEGMENT'] and orig_col in df_lbl.columns:
                    df_data[target] = df_lbl[orig_col]
                else:
                    df_data[target] = df_raw[orig_col]

            if 'TYPE' in df_data.columns:
                df_data = df_data[df_data['TYPE'].astype(str).str.contains('Growth|1|2', case=False, na=False)].copy()

            if selected_waves_filter != 'ALL' and 'WAVE' in df_data.columns:
                df_data = df_data[df_data['WAVE'].isin(selected_waves_filter)].copy()

            # 3. Clean officer names & codes
            def clean_officer_name(name):
                if pd.isna(name):
                    return name
                s = str(name).strip()
                s = re.sub(r'\s*\(.*?\)', '', s)
                return ' '.join(s.split())

            if 'OFFICER_NAME' in df_data.columns:
                df_data['OFFICER_NAME'] = df_data['OFFICER_NAME'].apply(clean_officer_name)

            if 'OFFICER_NAME' in df_data.columns and 'PRIM_OFCR_IND' in df_data.columns:
                df_data['OFFICER_NAME'] = df_data['OFFICER_NAME'].astype(str).str.strip()
                df_data['PRIM_OFCR_IND'] = df_data['PRIM_OFCR_IND'].astype(str).str.strip()
                df_data['OFFICER_NAME2'] = df_data['OFFICER_NAME'] + df_data['PRIM_OFCR_IND']
                df_data = df_data.sort_values(by='OFFICER_NAME2').reset_index(drop=True)
                unique_names = df_data['OFFICER_NAME2'].unique()
                name_to_code = {name: idx + 1 for idx, name in enumerate(unique_names)}
                df_data['OFFICE_CODE'] = df_data['OFFICER_NAME2'].map(name_to_code)

            # Save Streamlined SAV
            pyreadstat.write_sav(df_data, sav_output_filename)

            # 4. Build Excel Dashboard
            with pd.ExcelWriter(dest_filename, engine='openpyxl') as writer:
                df_data.to_excel(writer, sheet_name='data', index=False)

            wb = openpyxl.load_workbook(dest_filename)
            ws_toc = wb.create_sheet(title='TOC', index=0)
            ws_nps = wb.create_sheet(title='NPS', index=1)
            ws_data = wb['data']
            ws_data.sheet_state = 'hidden'

            TEAL_HEADER_FILL = PatternFill(start_color="00A3AD", end_color="00A3AD", fill_type="solid")
            LIGHT_TEAL_FILL = PatternFill(start_color="D9F2F4", end_color="D9F2F4", fill_type="solid")
            ORANGE_HEADER_FILL = PatternFill(start_color="F58220", end_color="F58220", fill_type="solid")
            LIGHT_ORANGE_FILL = PatternFill(start_color="FDF3EC", end_color="FDF3EC", fill_type="solid")
            BANNER_FILL = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
            ZEBRA_FILL = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")

            WHITE_BOLD_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
            BOLD_FONT = Font(name="Calibri", size=11, bold=True)
            REGULAR_FONT = Font(name="Calibri", size=11)
            THIN_BORDER = Border(left=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD'), top=Side(style='thin', color='DDDDDD'), bottom=Side(style='thin', color='DDDDDD'))

            ws_toc.cell(row=1, column=2, value="FNB Customer Satisfaction Study 2026: NPS Rating per BM/RM Portfolio")
            ws_toc.cell(row=1, column=2).fill = TEAL_HEADER_FILL
            ws_toc.cell(row=1, column=2).font = TITLE_FONT
            ws_toc.cell(row=1, column=2).alignment = Alignment(horizontal="center", vertical="center")

            ws_toc.append(["", "Select Wave:", "All Waves"])
            ws_toc.cell(row=2, column=2).font = BOLD_FONT
            ws_toc.cell(row=2, column=2).alignment = Alignment(horizontal="right")
            ws_toc.cell(row=2, column=3).fill = LIGHT_TEAL_FILL
            ws_toc.cell(row=2, column=3).border = THIN_BORDER

            def wave_sort_key(val):
                match = re.search(r'\d+', str(val))
                return int(match.group()) if match else 0

            raw_waves = df_data['WAVE'].dropna().unique() if 'WAVE' in df_data.columns else []
            available_waves = sorted(raw_waves, key=wave_sort_key)
            wave_list_str = '"All Waves,' + ','.join([str(w) for w in available_waves]) + '"'
            wave_dv = DataValidation(type="list", formula1=wave_list_str, allow_blank=False)
            ws_toc.add_data_validation(wave_dv)
            wave_dv.add(ws_toc['C2'])

            ws_toc.append(["", "Select Type:", "All Types"])
            ws_toc.cell(row=3, column=2).font = BOLD_FONT
            ws_toc.cell(row=3, column=2).alignment = Alignment(horizontal="right")
            ws_toc.cell(row=3, column=3).fill = LIGHT_TEAL_FILL
            ws_toc.cell(row=3, column=3).border = THIN_BORDER

            available_types = sorted(df_data['TYPE'].dropna().unique()) if 'TYPE' in df_data.columns else []
            type_list_str = '"All Types,' + ','.join([str(t) for t in available_types]) + '"'
            type_dv = DataValidation(type="list", formula1=type_list_str, allow_blank=False)
            ws_toc.add_data_validation(type_dv)
            type_dv.add(ws_toc['C3'])

            ws_toc.append([])
            ws_toc.append(["", "Table of Contents", ""])
            toc_header_row = 5
            ws_toc.cell(row=toc_header_row, column=2).font = Font(name="Calibri", size=12, bold=True, color="F58220")

            toc_entries = []

            def write_block(ws, title_text, officers_subset):
                start_row = ws.max_row + 1 if ws.max_row > 1 else 1
                if ws.max_row == 1 and ws['A1'].value is None:
                    start_row = 1
                ws.append([title_text] + [""] * 13)
                ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=14)
                banner_cell = ws.cell(row=ws.max_row, column=1)
                banner_cell.fill = BANNER_FILL
                banner_cell.font = WHITE_BOLD_FONT
                banner_cell.alignment = Alignment(horizontal="center", vertical="center")
                
                ws.append([""] * 14)
                back_row = ws.max_row
                back_cell = ws.cell(row=back_row, column=1, value="Back to TOC")
                back_cell.hyperlink = f"#TOC!B{toc_header_row}"
                back_cell.font = Font(name="Calibri", size=11, color="00A3AD", underline="single")
                
                h1_row = ws.max_row + 1
                ws.append(["BM/RM Name", "Officer Code", "Sub-Region", "NPS Segments_BM", "", "", "", "Count", "NPS Segments_FNB Business", "", "", "", "Count", "Officer Code"])
                h2_row = ws.max_row + 1
                ws.append(["", "", "", "NPS Score", "Detractors", "Passives", "Promoters", "", "NPS Score", "Detractors", "Passives", "Promoters", "", ""])
                
                for c in [1, 2, 3, 8, 14]:
                    ws.cell(row=h1_row, column=c).fill = BANNER_FILL
                    ws.cell(row=h1_row, column=c).font = WHITE_BOLD_FONT
                    ws.cell(row=h1_row, column=c).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=h2_row, column=c).fill = BANNER_FILL
                    ws.cell(row=h2_row, column=c).font = WHITE_BOLD_FONT
                    ws.cell(row=h2_row, column=c).alignment = Alignment(horizontal="center", vertical="center")

                for c in range(4, 8):
                    ws.cell(row=h1_row, column=c).fill = TEAL_HEADER_FILL
                    ws.cell(row=h1_row, column=c).font = WHITE_BOLD_FONT
                    ws.cell(row=h1_row, column=c).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=h2_row, column=c).fill = LIGHT_TEAL_FILL
                    ws.cell(row=h2_row, column=c).font = BOLD_FONT
                    ws.cell(row=h2_row, column=c).alignment = Alignment(horizontal="center", vertical="center")

                for c in range(9, 14):
                    ws.cell(row=h1_row, column=c).fill = ORANGE_HEADER_FILL
                    ws.cell(row=h1_row, column=c).font = WHITE_BOLD_FONT
                    ws.cell(row=h1_row, column=c).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=h2_row, column=c).fill = LIGHT_ORANGE_FILL
                    ws.cell(row=h2_row, column=c).font = BOLD_FONT
                    ws.cell(row=h2_row, column=c).alignment = Alignment(horizontal="center", vertical="center")

                row_counter = 0
                for _, row in officers_subset.iterrows():
                    formula_row = ws.max_row + 1
                    row_counter += 1
                    rm_name = row['OFFICER_NAME']
                    prim_code = row['PRIM_OFCR_IND']
                    col_c_val = str(title_text) if title_text != "TOTAL" else str(row['SUBREG'])
                    office_code = row['OFFICE_CODE']
                    
                    bm_count_formula = f'=IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$H:$H, ">0", data!$M:$M, N{formula_row}), IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$H:$H, ">0", data!$M:$M, N{formula_row}), IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$H:$H, ">0", data!$M:$M, N{formula_row}), COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$H:$H, ">0", data!$M:$M, N{formula_row}))))'
                    fnb_count_formula = f'=IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$I:$I, ">0", data!$M:$M, N{formula_row}), IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$I:$I, ">0", data!$M:$M, N{formula_row}), IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$I:$I, ">0", data!$M:$M, N{formula_row}), COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$I:$I, ">0", data!$M:$M, N{formula_row}))))'
                    
                    bm_det_formula = f'=IFERROR(IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$H:$H, 1, data!$M:$M, N{formula_row})/H{formula_row}, IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$H:$H, 1, data!$M:$M, N{formula_row})/H{formula_row}, IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$H:$H, 1, data!$M:$M, N{formula_row})/H{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$H:$H, 1, data!$M:$M, N{formula_row})/H{formula_row}))), 0)'
                    bm_pas_formula = f'=IFERROR(IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$H:$H, 2, data!$M:$M, N{formula_row})/H{formula_row}, IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$H:$H, 2, data!$M:$M, N{formula_row})/H{formula_row}, IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$H:$H, 2, data!$M:$M, N{formula_row})/H{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$H:$H, 2, data!$M:$M, N{formula_row})/H{formula_row}))), 0)'
                    bm_pro_formula = f'=IFERROR(IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$H:$H, 3, data!$M:$M, N{formula_row})/H{formula_row}, IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$H:$H, 3, data!$M:$M, N{formula_row})/H{formula_row}, IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$H:$H, 3, data!$M:$M, N{formula_row})/H{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$H:$H, 3, data!$M:$M, N{formula_row})/H{formula_row}))), 0)'
                    
                    fnb_det_formula = f'=IFERROR(IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$I:$I, 1, data!$M:$M, N{formula_row})/M{formula_row}, IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$I:$I, 1, data!$M:$M, N{formula_row})/M{formula_row}, IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$I:$I, 1, data!$M:$M, N{formula_row})/M{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$I:$I, 1, data!$M:$M, N{formula_row})/H{formula_row}))), 0)'
                    fnb_pas_formula = f'=IFERROR(IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$I:$I, 2, data!$M:$M, N{formula_row})/M{formula_row}, IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$I:$I, 2, data!$M:$M, N{formula_row})/M{formula_row}, IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$I:$I, 2, data!$M:$M, N{formula_row})/M{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$I:$I, 2, data!$M:$M, N{formula_row})/H{formula_row}))), 0)'
                    fnb_pro_formula = f'=IFERROR(IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$I:$I, 3, data!$M:$M, N{formula_row})/M{formula_row}, IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$I:$I, 3, data!$M:$M, N{formula_row})/M{formula_row}, IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$I:$I, 3, data!$M:$M, N{formula_row})/M{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$I:$I, 3, data!$M:$M, N{formula_row})/H{formula_row}))), 0)'

                    ws.append([
                        rm_name, prim_code, col_c_val,
                        f'=IF(H{formula_row}>0, SUM(G{formula_row}-E{formula_row})*100, "")',
                        bm_det_formula, bm_pas_formula, bm_pro_formula, bm_count_formula,
                        f'=IF(M{formula_row}>0, SUM(L{formula_row}-J{formula_row})*100, "")',
                        fnb_det_formula, fnb_pas_formula, fnb_pro_formula, fnb_count_formula,
                        office_code
                    ])
                    
                    ws[f'D{formula_row}'].number_format = '0'
                    ws[f'E{formula_row}'].number_format = '0%'
                    ws[f'F{formula_row}'].number_format = '0%'
                    ws[f'G{formula_row}'].number_format = '0%'
                    ws[f'I{formula_row}'].number_format = '0'
                    ws[f'J{formula_row}'].number_format = '0%'
                    ws[f'K{formula_row}'].number_format = '0%'
                    ws[f'L{formula_row}'].number_format = '0%'

                    for c in range(1, 15):
                        cell = ws.cell(row=formula_row, column=c)
                        cell.border = THIN_BORDER
                        cell.font = REGULAR_FONT
                        if row_counter % 2 == 0:
                            cell.fill = ZEBRA_FILL

                ws.append([])
                return start_row

            unique_officers_all = df_data[['OFFICER_NAME', 'PRIM_OFCR_IND', 'SUBREG', 'OFFICE_CODE']].drop_duplicates().sort_values(by='OFFICE_CODE')
            total_start_row = write_block(ws_nps, "TOTAL", unique_officers_all)
            toc_entries.append(("Officer Name and Code by NPS Banner", total_start_row, "TOTAL"))

            if 'SUBREG' in df_data.columns:
                for subreg in sorted(df_data['SUBREG'].dropna().unique()):
                    df_subreg = df_data[df_data['SUBREG'] == subreg]
                    unique_officers_subreg = df_subreg[['OFFICER_NAME', 'PRIM_OFCR_IND', 'SUBREG', 'OFFICE_CODE']].drop_duplicates().sort_values(by='OFFICE_CODE')
                    subreg_start_row = write_block(ws_nps, str(subreg), unique_officers_subreg)
                    toc_entries.append((f"Officer Name and Code by NPS Banner {subreg}", subreg_start_row, str(subreg)))

            for label, nps_row_num, subreg_filter in toc_entries:
                row_idx = ws_toc.max_row + 1
                if subreg_filter == "TOTAL":
                    filter_formula = f'=CONCATENATE("Filter: Wave ", TOC!$C$2, ", Type ", TOC!$C$3, ", base n =", IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNT(data!$A:$A), IF(TOC!$C$2="All Waves", COUNTIF(data!$C:$C, TOC!$C$3), IF(TOC!$C$3="All Types", COUNTIF(data!$D:$D, TOC!$C$2), COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3)))))'
                else:
                    filter_formula = f'=CONCATENATE("Filter: Wave ", TOC!$C$2, ", Type ", TOC!$C$3, ", base n =", IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIF(data!$F:$F, "{subreg_filter}"), IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$F:$F, "{subreg_filter}"), IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$F:$F, "{subreg_filter}"), COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$F:$F, "{subreg_filter}")))))'

                ws_toc.cell(row=row_idx, column=1, value="NPS").font = BOLD_FONT
                ws_toc.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center")
                
                link_cell = ws_toc.cell(row=row_idx, column=2, value=label)
                link_cell.hyperlink = f"#NPS!A{nps_row_num}"
                link_cell.font = Font(name="Calibri", size=11, color="F58220", underline="single")
                
                ws_toc.cell(row=row_idx, column=3, value=filter_formula).font = REGULAR_FONT

            for ws in wb.worksheets:
                if ws.title == 'TOC':
                    ws.column_dimensions['A'].width = 3.67
                    ws.column_dimensions['B'].width = 78.22
                    ws.column_dimensions['C'].width = 43.11
                elif ws.title == 'NPS':
                    ws.column_dimensions['A'].width = 27.11
                    ws.column_dimensions['B'].width = 10.67
                    ws.column_dimensions['C'].width = 20.11
                    ws.column_dimensions['D'].width = 16.44
                    ws.column_dimensions['I'].width = 24.67
                    for col in ['E', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'N']:
                        ws.column_dimensions[col].width = 11

            wb.save(dest_filename)

        st.success("🎉 Processing complete! Download your files below:")

        # --- Download Buttons ---
        with open(dest_filename, "rb") as f:
            st.download_button(
                label="📥 Download Excel Dashboard",
                data=f,
                file_name=dest_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with open(sav_output_filename, "rb") as f:
            st.download_button(
                label="📥 Download Streamlined SPSS (.sav) File",
                data=f,
                file_name=sav_output_filename,
                mime="application/octet-stream"
            )