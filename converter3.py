import re
import xlsxwriter

def parse_input(input_text):
    cleaned_input = re.sub(r'--\s*', '', input_text.strip())

    query_match = re.search(r'(select.*?;)', cleaned_input, re.IGNORECASE | re.DOTALL)
    class_match = re.search(r'class\s*:\s*(\S+)', cleaned_input, re.IGNORECASE)
    line_no_match = re.search(r'line\s*:\s*(\d+)', cleaned_input, re.IGNORECASE)
    actions_match = re.search(r'(method|actions?)\s*:\s*(\S+)', cleaned_input, re.IGNORECASE)
    screen_no_match = re.search(r'screen\s*:\s*(\d+)', cleaned_input, re.IGNORECASE)
    data_id_match = re.search(r'data\s*id\s*:\s*(\S+)', cleaned_input, re.IGNORECASE)
    process_step_id_match = re.search(r'process\s*step\s*id\s*:\s*(\S+)', cleaned_input, re.IGNORECASE)
    case_id_match = re.search(r'case\s*id\s*:\s*(\S+)', cleaned_input, re.IGNORECASE)

    entry = {
        'Query': query_match.group(1).strip() if query_match else '',
        'Class': class_match.group(1) if class_match else '',
        'Line No': line_no_match.group(1) if line_no_match else '',
        'actions': actions_match.group(2) if actions_match else '',
        'Screen No': screen_no_match.group(1) if screen_no_match else '',
        'Data ID': data_id_match.group(1) if data_id_match else '',
        'Process Step ID': process_step_id_match.group(1) if process_step_id_match else '',
        'Case ID': case_id_match.group(1) if case_id_match else ''
    }
    return entry

def write_to_excel(parsed_entry, output_file='output.xlsx'):
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    bold_format = workbook.add_format({'bold': True})
    normal_format = workbook.add_format()

    row = 0
    col = 0

    labels = [
        'Query',
        'Class',
        'Line No',
        'actions',
        'Screen No',
        'Data ID',
        'Process Step ID',
        'Case ID'
    ]

    rich_text = []
    for label in labels:
        value = parsed_entry.get(label, '')
        # Append label in bold
        rich_text.append(bold_format)
        rich_text.append(f"{label}: ")
        # Append value in normal format or space if missing
        if value:
            rich_text.append(normal_format)
            rich_text.append(f"{value}\n")
        else:
            rich_text.append('\n')

    # Write rich text to the cell
    worksheet.write_rich_string(row, col, *rich_text)

    workbook.close()

# Sample Input Text
input_text = """
-- Line: 535 
-- class: CM_SecDepst_Refund_Leasing
-- Screen 1
-- method: cmd_ok_ActionPerformed
-- EVENT: all radio btn click
select value(AVL_SEQ_NO,0) as AVL_SEQ_NO,int_lim_no,int_lim_year,LOAN_TYPE,digits(LOAN_NO) as LOAN_NO,digits(LOAN_YEAR) as LOAN_YEAR,digits(BRN_CD) as BRN_CD,digits(AC_TYPE) as AC_TYPE,digits(CUST_NO) as CUST_NO,digits(RUN_NO) as RUN_NO,digits(CHK_DIGT) as CHK_DIGT,digits(NOM_AC_TYPE) as NOM_AC_TYPE,digits(NOM_CUST_NO) as NOM_CUST_NO,digits(NOM_RUN_NO) as NOM_RUN_NO,digits(NOM_CHK_DIGT) as NOM_CHK_DIGT,digits(LOAN_AC_TYPE) as LOAN_AC_TYPE,digits(LOAN_CUST_NO) as LOAN_CUST_NO,digits(LOAN_RUN_NO) as LOAN_RUN_NO,digits(LOAN_CHK_DIGT) as LOAN_CHK_DIGT,digits(OVR_AC_TYPE) as OVR_AC_TYPE,digits(OVR_CUST_NO) as OVR_CUST_NO,digits(OVR_RUN_NO) as OVR_RUN_NO,digits(OVR_CHK_DIGT) as OVR_CHK_DIGT,digits(CR_AC_TYPE) as CR_AC_TYPE,digits(CR_CUST_NO) as CR_CUST_NO,digits(CR_RUN_NO) as CR_RUN_NO,digits(CR_CHK_DIGT) as CR_CHK_DIGT,GRANT_DATE,MAT_DATE,LOAN_AMT,FREQ,DOC_TYPE,VALUE(DOC_NO,0) AS DOC_NO,VALUE(DOC_YEAR,0) AS DOC_YEAR,LUM_SUM_INST,VALUE(SECDEP_NO,0) AS SECDEP_NO,VALUE(SECDEP_PERC,0) AS SECDEP_PERC,VALUE(SECDEP_AMT,0) AS SECDEP_AMT,value(SALVAGE_PERC,0) as SALVAGE_PERC,value(Salvage_AMT,0) as Salvage_AMT,digits(Salvage_AC_TYPE) as Salvage_AC_TYPE,digits(Salvage_CUST_NO) as Salvage_CUST_NO,digits(Salvage_RUN_NO) as Salvage_RUN_NO,digits(Salvage_CHK_DIGT) as Salvage_CHK_DIGT, VALUE(TENURE,0) as TENURE,TENURE_UNITS,VALUE(NO_OF_INST,0) as NO_OF_INST,REPAY_FREQ,MORAT_DT,VALUE(MORAT_PRD,0) as MORAT_PRD,INST_SCH_DT,value(status,'') as status,value(accr_amt,0) as accr_amt,value(rec_amt,0) as rec_amt,digits(RT_CODE) as RT_CODE,BASE_RATE,SPREAD_RT,FLOOR_RT,CEILING_RT,MU_RATE,REMARKS ,value(charity_amt,0) as charity_amt from cm_loan_tl where  brn_cd=1024 and full='Y' and (loan_amt=rec_amt) and status='U'  and loan_type in ('LF')  and GRANT_DATE!=date('2024-10-31');
"""

# Parse input and write to Excel
parsed_entry = parse_input(input_text)
write_to_excel(parsed_entry)