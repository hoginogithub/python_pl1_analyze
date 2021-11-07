import openpyxl

result_sheet = ""

def init():
    try:
        excel_path = r"C:\Users\hogin\Documents\prog\python\pl1\PL1解析マクロ.xlsm"
        wb = openpyxl.load_workbook(excel_path)
        global result_sheet 
        result_sheet = wb["比較結果"]
        return True
    except Exception:
        return False

def make_mcl_col():
    if result_sheet["F3"].value == "UT Case ID":
        return 


if "__main__" == __name__:
    if not (init()):
        exit

    make_mcl_col()

    for row in range(3, result_sheet.max_row + 1):
        print(result_sheet["D" + str(row)].value)