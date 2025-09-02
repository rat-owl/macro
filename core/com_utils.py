# core/com_utils.py

import os
import win32com.client
import pythoncom
from win32com.client import Dispatch


def hwp_opened():
    context = pythoncom.CreateBindCtx(0)
    running_coms = pythoncom.GetRunningObjectTable()
    monikers = running_coms.EnumRunning()

    for moniker in monikers:
        name = moniker.GetDisplayName(context, moniker)
        if name == '!HwpObject.96.1':
            obje = running_coms.GetObject(moniker)
            hwp = Dispatch(obje.QueryInterface(pythoncom.IID_IDispatch))
            return hwp


def get_excel_instance():
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except:
        excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    return excel


def is_workbook_open(excel, workbook_name):
    for workbook in excel.Workbooks:
        if workbook.Name.lower() == workbook_name.lower():
            return True
    return False


def open_workbook(excel, excel_file):
    workbook_name = os.path.basename(excel_file)

    if is_workbook_open(excel, workbook_name):
        print(f"'{workbook_name}' 워크북은 이미 열려 있습니다.")
        wb = excel.Workbooks[workbook_name]
    else:
        print(f"'{workbook_name}' 워크북을 여는 중...")
        wb = excel.Workbooks.Open(excel_file)
        try:
            wb.Activate()
            excel.Application.WindowState = -4137  # xlMaximized
        except:
            pass

    return wb
