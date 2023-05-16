# identify spreadsheet
# use information in columns to create value pairs
# identify directories
# find folder names A2+B2
# find file in column D2 and copy to folder named A2+B2
# find file in column G1 and copy to folder named A2+B2
# find file in column I1 and copy to folder named A2+B2
# loop

import os
import openpyxl
import pandas
import shutil

PHASE_3_XL = "//khfiler01/Searchflow/CSF/oper/Operations/Portfolios/Portfolios/Land Registry Extracts/LRE 1712 Penningtons/Script testing/Phase 3 - First Batch.xlsx"
REGISTERS = "//khfiler01/Searchflow/CSF/oper/Operations/Portfolios/Portfolios/Land Registry Extracts/LRE 1712 Penningtons/Script testing/Registers"
ADD_LEASES = "//khfiler01/Searchflow/CSF/oper/Operations/Portfolios/Portfolios/Land Registry Extracts/LRE 1712 Penningtons/Script testing/Additional Leases"
PHASE3_COMPLETE = "//khfiler01/Searchflow/CSF/oper/Operations/Portfolios/Portfolios/Land Registry Extracts/LRE 1712 Penningtons/Script testing/Phase 3 - Batch 1 Complete"

excel_output = pandas.read_excel(PHASE_3_XL, sheet_name="Sheet1")

for row in excel_output.iterrows():
    row = row[1].to_dict()
    Block_Code = row["Block Code"]
    Block_Name = row["Block Name"]
    ColumnD = row["freehold_title_number"]
    ColumnG = row["head_leasehold_title_number"]
    ColumnI = row["under_leasehold_title_number"]
    output_file = f"{PHASE3_COMPLETE}/{Block_Code} - {Block_Name}"
    registerD = f"{REGISTERS}/{ColumnD}.pdf"
    registerG = f"{REGISTERS}/{ColumnG}.pdf"
    registerI = f"{REGISTERS}/{ColumnI}.pdf"
    shutil.copyfile(registerD, output_file)
    shutil.copyfile(registerG, output_file)
    shutil.copyfile(registerI, output_file)
    input()