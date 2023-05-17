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
    if f'{ColumnD}' != "0":
        try:
            shutil.copy(f"{registerD}", output_file)
            print(f"Register {registerD} has been copied")
        except FileNotFoundError as error:
            pass
        except OSError as error:
            pass
    if f'{ColumnG}' != "0":
        try:
            shutil.copy(registerG, output_file)
            print(f"Register {registerG} has been copied")
        except FileNotFoundError as error:
            pass
        except OSError as error:
            pass
    if f'{ColumnI}' != "0":
        try:
            shutil.copy(registerI, output_file)
            print(f"Register {registerI} has been copied")
        except FileNotFoundError as error:
            pass
        except OSError as error:
            pass
