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
-- class: TfinUtil
-- Method: getCustomerInfo
-- Screen : 1
-- Line: 8476
select a.ledg_bal as ledgbal, a.shadow_d as shadow_d, a.shadow_c as shadow_c, a.lim_amt as lim_amt, a.amt_blk as block, a.uncl_amt as uncl_amt, a.AC_OP_DT, a.MOTHER_NAME, a.acc_cat_cd, a.ac_title, A.IN_ACR_C, A.IN_ACR_d, VALUE(A.MAIL_ADDRESS1,' ') as mad1, VALUE(A.MAIL_ADDRESS2,' ') as mad2, A.AC_TITLE from ACCOUNT_TL A where a.cust_no = 012664 and a.ac_type = 0966 and a.run_no = 01 and a.chk_digt = 6 and a.brn_cd = 1024;
"""

# Parse input and write to Excel
parsed_entry = parse_input(input_text)
write_to_excel(parsed_entry)