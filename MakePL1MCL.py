import openpyxl
from openpyxl.styles import Alignment
from openpyxl.styles import Color

try:
    excel_path = r"C:\Users\hogin\Documents\prog\python\pl1\PL1解析マクロ.xlsm"
    wb = openpyxl.load_workbook(excel_path, read_only=False, keep_vba=True)
    result_sheet = wb["比較結果"]

    if result_sheet["F2"].value == "UT Case ID":
        pass
    else:
        result_sheet.insert_cols(6)
        result_sheet["F2"].value == "UT Case ID"
        result_sheet["F2"].alignment = Alignment("center", "center")
        
    comment_flag = False
    string_flag = False
    
    for row in range(3, result_sheet.max_row + 1):
        src = str(result_sheet["D" + str(row)].value).upper()
        while len(src) == 0:
            if comment_flag == True:
                start_comment_start = find(src,  )
    

    wb.save(excel_path)
    
except Exception as e:
    print(f"{e}")