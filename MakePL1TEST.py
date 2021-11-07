import openpyxl
from openpyxl.workbook.workbook import Workbook


if "__main__" == __name__:

    excel_path = r"C:\Users\hogin\Documents\prog\python\pl1\PL1解析マクロ.xlsx"
    wb = openpyxl.load_workbook(excel_path)
    result_sheet = wb["比較結果"]

    if result_sheet["F2"].value == "UT Case ID":
        exit
    print(f'F2={result_sheet["F2"].value}')
    try:
        result_sheet.insert_cols(5)
        result_sheet["F2"].value == "UT Case ID"
    except Exception as e:
        print(f"例外でした...{e}")
        exit

    """
    for row in range(3, result_sheet.max_row + 1):
        print(result_sheet["D" + str(row)].value)
    """

    wb.save(excel_path)