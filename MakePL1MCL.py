import openpyxl
from openpyxl.styles import Alignment
from openpyxl.styles import Color

line_num = 0
comment_flag = False 

def search_comment_start(src):
    global line_num

    ret = src.find("/*")
    # print(f"call comment start line:{line_num} pos='/*' pos={ret}")
    if ret == -1:
        return src
    global cooment_flag
    comment_flag = True
    src = "" if ret == 1 else src[0:ret -1]

    src = src + search_comment_end(src[ret + 2:])

def search_comment_end(src):
    global cooment_flag
    ret = src.find("*/")
    print(f"call comment start line:{line_num} pos='*/' pos={ret}")
    
    if ret == -1:
        return ""

    comment_flag = False
    src = src[ret + 2:]

    src = src + search_comment_start(src)

if "__main__" == __name__:
    try:
        excel_path = r"C:\Users\hogin\Documents\prog\python\pl1\PL1解析マクロ.xlsm"
        wb = openpyxl.load_workbook(excel_path, read_only=False, keep_vba=True)
        result_sheet = wb["比較結果"]
    
        if result_sheet["F2"].value == "UT Case ID":
            pass
        else:
            result_sheet.insert_cols(6)
            result_sheet["F2"].value = "UT Case ID"
            result_sheet["F2"].alignment = Alignment("center", "center")
            
        comment_flag = False
        string_flag = False
        
        line_num = 0
        for row in result_sheet["D:D"]:
            src = "" if str(row.value) == "None" else str(row.value)
            
            line_num += 1
            src = src.upper()
            
            src = search_comment_start(src) if comment_flag == False else search_comment_end(src)

            print(f"STEP:{line_num}={src}")
            
        
    
        wb.save(excel_path)
        
    except Exception as e:
        print(f"{e}")