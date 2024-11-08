# Python Query Optimizer

This project is a Python-based tool that parses SQL queries and related metadata from a given input text and writes the parsed information into an Excel file with formatted text. The labels are bolded for better readability.

## Features

- **Parse Input Text:** Extracts SQL queries and related metadata from a structured input text.
- **Write to Excel:** Outputs the parsed information into an Excel file with bolded labels.
- **Rich Text Formatting:** Uses `xlsxwriter` to apply rich text formatting in the Excel file.

## Dependencies

- Python 3.x
- `xlsxwriter` library

You can install the required dependencies using pip:

```bash
pip3 install xlsxwriter
```

## Usage
1. Prepare Input Text: Modify the input_text variable in converter3.py with your structured input text.

2. Run the Script: Execute the script to parse the input text and generate the Excel file.

```bash
python3 [converter3.py](http://_vscodecontentref_/0)
```
3. Output: The script will generate an output.xlsx file in the current directory with the parsed information.


## Example Input
``` bash
input_text = """
-- Line: 535 
-- class: CM_SecDepst_Refund_Leasing
-- Screen 1
-- method: cmd_ok_ActionPerformed
-- EVENT: all radio btn click
select value(AVL_SEQ_NO,0) as AVL_SEQ_NO,int_lim_no,int_lim_year,LOAN_TYPE,digits(LOAN_NO) as LOAN_NO,digits(LOAN_YEAR) as LOAN_YEAR,digits(BRN_CD) as BRN_CD,digits(AC_TYPE) as AC_TYPE,digits(CUST_NO) as CUST_NO,digits(RUN_NO) as RUN_NO,digits(CHK_DIGT) as CHK_DIGT,digits(NOM_AC_TYPE) as NOM_AC_TYPE,digits(NOM_CUST_NO) as NOM_CUST_NO,digits(NOM_RUN_NO) as NOM_RUN_NO,digits(NOM_CHK_DIGT) as NOM_CHK_DIGT,digits(LOAN_AC_TYPE) as LOAN_AC_TYPE,digits(LOAN_CUST_NO) as LOAN_CUST_NO,digits(LOAN_RUN_NO) as LOAN_RUN_NO,digits(LOAN_CHK_DIGT) as LOAN_CHK_DIGT,digits(OVR_AC_TYPE) as OVR_AC_TYPE,digits(OVR_CUST_NO) as OVR_CUST_NO,digits(OVR_RUN_NO) as OVR_RUN_NO,digits(OVR_CHK_DIGT) as OVR_CHK_DIGT,digits(CR_AC_TYPE) as CR_AC_TYPE,digits(CR_CUST_NO) as CR_CUST_NO,digits(CR_RUN_NO) as CR_RUN_NO,digits(CR_CHK_DIGT) as CR_CHK_DIGT,GRANT_DATE,MAT_DATE,LOAN_AMT,FREQ,DOC_TYPE,VALUE(DOC_NO,0) AS DOC_NO,VALUE(DOC_YEAR,0) AS DOC_YEAR,LUM_SUM_INST,VALUE(SECDEP_NO,0) AS SECDEP_NO,VALUE(SECDEP_PERC,0) AS SECDEP_PERC,VALUE(SECDEP_AMT,0) AS SECDEP_AMT,value(SALVAGE_PERC,0) as SALVAGE_PERC,value(Salvage_AMT,0) as Salvage_AMT,digits(Salvage_AC_TYPE) as Salvage_AC_TYPE,digits(Salvage_CUST_NO) as Salvage_CUST_NO,digits(Salvage_RUN_NO) as Salvage_RUN_NO,digits(Salvage_CHK_DIGT) as Salvage_CHK_DIGT, VALUE(TENURE,0) as TENURE,TENURE_UNITS,VALUE(NO_OF_INST,0) as NO_OF_INST,REPAY_FREQ,MORAT_DT,VALUE(MORAT_PRD,0) as MORAT_PRD,INST_SCH_DT,value(status,'') as status,value(accr_amt,0) as accr_amt,value(rec_amt,0) as rec_amt,digits(RT_CODE) as RT_CODE,BASE_RATE,SPREAD_RT,FLOOR_RT,CEILING_RT,MU_RATE,REMARKS ,value(charity_amt,0) as charity_amt from cm_loan_tl where  brn_cd=1024 and full='Y' and (loan_amt=rec_amt) and status='U'  and loan_type in ('LF')  and GRANT_DATE!=date('2024-10-31');
"""
```

## Future Updates
1. **Modularization**: Refactor the code to improve modularity and readability.
2. **Error Handling**: Add comprehensive error handling to manage various edge cases and input anomalies.
3. **Configuration File**: Allow input text and output file paths to be specified via a configuration file or command-line arguments.
4. **Unit Tests**: Implement unit tests to ensure the reliability and correctness of the parsing and writing functionalities.
5. **Additional Formats**: Support additional output formats such as CSV or JSON.

## Contributing
Contributions are welcome! Please fork the repository and submit a pull request for any enhancements or bug fixes.

## License
This project is licensed under the MIT License.
