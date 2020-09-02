import win32com.client

def __close_excel_window(file_name):
    """terminate excel application"""
    xl = win32com.client.Dispatch("Excel.Application")
    xl.DisplayAlerts = False
    for wb in xl.workbooks:
        if file_name.lower() in wb.name.lower():
            wb.Close(True)
            break
    return 

fn = "2020-07-17 210723 NH55_SCC0011827_2H.xlsx"
__close_excel_window(fn)
