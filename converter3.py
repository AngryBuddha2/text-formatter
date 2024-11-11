import re
import xlsxwriter

def parse_input(input_text):
    # Split the input text by '###'
    entries = input_text.strip().split('###')
    parsed_entries = []
    for entry in entries:
        cleaned_input = re.sub(r'--\s*', '', entry.strip())
        # Match the SQL query
        query_match = re.search(
            r'^\s*(select|update|delete|with)\s.*?;',
            cleaned_input,
            re.IGNORECASE | re.DOTALL | re.MULTILINE
        )
        # Extract different parts using regex
        class_match = re.search(r'class\s*[:\-]?\s*(\S+)', cleaned_input, re.IGNORECASE)
        line_no_match = re.search(r'line\s*[:\-]?\s*(\d+)', cleaned_input, re.IGNORECASE)
        actions_match = re.search(r'(method|actions?)\s*[:\-]?\s*(\S+)', cleaned_input, re.IGNORECASE)
        screen_no_match = re.search(r'screen\s*[:\-]?\s*(\d+)', cleaned_input, re.IGNORECASE)
        data_id_match = re.search(r'data\s*id\s*[:\-]?\s*(\S+)', cleaned_input, re.IGNORECASE)
        process_step_id_match = re.search(r'process\s*step\s*id\s*[:\-]?\s*(\S+)', cleaned_input, re.IGNORECASE)
        case_id_match = re.search(r'case\s*id\s*[:\-]?\s*(\S+)', cleaned_input, re.IGNORECASE)

        # Combine Process Step ID and Data ID
        if process_step_id_match and data_id_match:
            process_data_id = f"{process_step_id_match.group(1)}/{data_id_match.group(1)}"
        elif process_step_id_match:
            process_data_id = process_step_id_match.group(1)
        elif data_id_match:
            process_data_id = data_id_match.group(1)
        else:
            process_data_id = ''

        parsed_entry = {
            'Process/Data ID': process_data_id,
            'Class': class_match.group(1) if class_match else '',
            'Line No': line_no_match.group(1) if line_no_match else '',
            'Actions': actions_match.group(2) if actions_match else '',
            'Screen No': screen_no_match.group(1) if screen_no_match else '',
            'Data ID': data_id_match.group(1) if data_id_match else '',
            'Process Step ID': process_step_id_match.group(1) if process_step_id_match else '',
            'Case ID': case_id_match.group(1) if case_id_match else ''

        }
        parsed_entries.append(parsed_entry)
    return parsed_entries

def write_to_excel(parsed_entries, output_file='output.xlsx'):
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()
    bold_format = workbook.add_format({'bold': True})
    normal_format = workbook.add_format({'text_wrap': True})
    # Write headers
    worksheet.write(0, 0, 'Comment', bold_format)
    worksheet.write(0, 1, 'Process/Data ID', bold_format)
    worksheet.write(0, 2, 'Case ID', bold_format)
    # Labels to include in the comment (excluding 'Process/Data ID' and 'Case ID')
    comment_labels = [
            'Query',
            'Class',
            'Line No',
            'Actions',
            'Screen No',
            'Data ID',
            'Process Step ID',
            'Case ID'
        ]
    # Write data
    row = 1
    for entry in parsed_entries:
        # Prepare rich text for the comment
        rich_text = []
        for label in comment_labels:
            value = entry.get(label, '')
            rich_text.append(bold_format)
            rich_text.append(f"{label}: ")
            if value:
                rich_text.append(normal_format)
                rich_text.append(f"{value}\n")
            else:
                rich_text.append('\n')
        # Write rich text to the first column
        worksheet.write_rich_string(row, 0, *rich_text)
        # Write 'Process/Data ID' to the second column
        worksheet.write(row, 1, entry.get('Process/Data ID', ''), normal_format)
        # Write 'Case ID' to the third column
        worksheet.write(row, 2, entry.get('Case ID', ''), normal_format)
        row += 1
    workbook.close()

# Sample Input Text
input_text = """
-- Line 535 
-- class: CM_SecDepst_Refund_Leasing
-- Screen: 1
-- method: cmd_ok_ActionPerformed
-- EVENT: all radio btn click
-- Data ID: U.A.28.79.D.3
-- Process Step ID: P.A.28.79.D.3
select value(AVL_SEQ_NO,0) as AVL_SEQ_NO from cm_loan_tl where brn_cd=1024;

###

-- class: TfinUtil
-- Method: getCustomerInfo
-- Screen : 1
-- Line: 8476
-- Case ID: C123
update ACCOUNT_TL set ledg_bal = ledg_bal + 100 where cust_no = 012664;

###

-- Case ID: C456
-- Actions: deleteAccount
-- Data ID: D789
delete from ACCOUNT_TL where cust_no = 012345;
"""

# Parse input and write to Excel
parsed_entries = parse_input(input_text)
write_to_excel(parsed_entries)