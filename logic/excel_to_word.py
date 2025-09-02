# logic/excel_to_word.py

import time
from win32com.client import Dispatch
from core.com_utils import get_excel_instance, open_workbook


def excel_to_word(excel_file):
    excel = get_excel_instance()
    wb = open_workbook(excel, excel_file)
    wsSource = wb.Sheets("전분류군(집계)")

    word = Dispatch("Word.Application")
    word.Visible = True
    doc = word.Documents.Add()

    def write_to_word(sheet_name, a, b):
        print(f"시트 [{sheet_name}] 데이터 입력 시작 (a={a}, b={b})")
        if a == 0 and b == 0:
            return

        wsList = wb.Sheets(sheet_name)

        first_find = None
        last_row = None

        if sheet_name == "조류군집분석표":
            find_result = wsList.UsedRange.Find("J'", MatchCase=True)
        else:
            find_result = wsList.UsedRange.Find("출현종수", MatchCase=True)

        if find_result:
            first_find = find_result.Address
            while True:
                last_row = find_result.Row
                find_result = wsList.UsedRange.FindNext(find_result)
                if find_result.Address == first_find:
                    break

        if not last_row:
            raise ValueError(f"'{sheet_name}' 시트에서 '출현종수'를 찾을 수 없습니다.")

        table_range = wsList.Range(wsList.Cells(2, 1), wsList.Cells(last_row, 6 + a + b))
        table_range.Copy()
        time.sleep(1)
        selection = word.Selection
        selection.Collapse(0)
        selection.Paste()
        selection.InsertParagraphAfter()
        print(f"시트 [{sheet_name}] 데이터 입력 완료")

    def find_row_by_keyword(ws, keyword):
        used_range = ws.UsedRange
        for row in range(1, used_range.Rows.Count + 1):
            cell_value = str(ws.Cells(row, 2).Value)
            if keyword in cell_value:
                return row
        raise ValueError(f"'{keyword}' 키워드를 찾을 수 없습니다.")

    categories = {
        "포유류": "포유류목록",
        "조류": "조류목록",
        "양서파충류": "양서파충류목록",
        "육상곤충": "곤충목록",
        "조류군집": "조류군집분석표"
    }

    for keyword, sheet_name in categories.items():
        try:
            row = find_row_by_keyword(wsSource, keyword)
            a = wsSource.Cells(row, 3).Value or 0

            if sheet_name == "조류군집분석표":
                b = 0
            else:
                b = wsSource.Cells(row, 4).Value or 0

            write_to_word(sheet_name, int(a), int(b))

        except Exception as e:
            print(f"[{sheet_name}] 처리 실패: {e}")

    doc.Content.WholeStory()
    doc.Content.Copy()

    return word


def insect_to_word(excel_file):
    excel = get_excel_instance()
    wb = open_workbook(excel, excel_file)
    wsSource = wb.Sheets("전분류군(집계)")

    word = Dispatch("Word.Application")
    word.Visible = True
    doc = word.Documents.Add()

    def write_to_insect(sheet_name, a, b):
        if a == 0 and b == 0:
            print(f"[{sheet_name}] 스킵: a={a}, b={b}")
            return

        print(f"[{sheet_name}] 실행: a={a}, b={b}")
        wsList = wb.Sheets(sheet_name)
        find_result = wsList.UsedRange.Find("출현종수", MatchCase=True)
        if not find_result:
            raise ValueError(f"'{sheet_name}' 시트에서 '출현종수'를 찾을 수 없습니다.")

        first_find = find_result.Address
        last_row = find_result.Row

        while True:
            find_result = wsList.UsedRange.FindNext(find_result)
            if not find_result or find_result.Address == first_find:
                break
            last_row = find_result.Row

        def find_middle_row(wsList, col, start_row):
            row = start_row
            while wsList.Cells(row, col).Value is not None:
                row += 1
            return row - 1

        Middle_row = find_middle_row(wsList, 1, 3)
        print(f"[{sheet_name}] 첫 Middle_row: {Middle_row}")
        selection = word.Selection

        try:
            table_range = wsList.Range(wsList.Cells(2, 1), wsList.Cells(Middle_row, 6 + a + b))
            if table_range:
                table_range.Copy()
                time.sleep(0.5)
                selection.Collapse(0)
                selection.PasteSpecial(DataType=22)
                selection.InsertParagraphAfter()
                print(f"[{sheet_name}] {2}~{Middle_row} 붙여넣기 완료")
        except Exception as e:
            print(f"[{sheet_name}] 오류 발생: {e}")
            raise e

        while Middle_row < last_row:
            first_row = Middle_row + 2
            Middle_row = find_middle_row(wsList, 1, first_row)

            print(f"[{sheet_name}] 복사할 범위: {first_row}~{Middle_row}")

            try:
                table_range = wsList.Range(wsList.Cells(first_row, 1), wsList.Cells(Middle_row, 6 + a + b))
                if table_range:
                    table_range.Copy()
                    time.sleep(0.5)
                    selection.Collapse(0)
                    selection.PasteSpecial(DataType=22)
                    selection.InsertParagraphAfter()
                    print(f"[{sheet_name}] {first_row}~{Middle_row} 붙여넣기 완료")
            except Exception as e:
                print(f"[{sheet_name}] 오류 발생: {e}")
                raise e

        print(f"[{sheet_name}] 완료")

    a = wsSource.Cells(20, 3).Value or 0
    b = wsSource.Cells(20, 4).Value or 0
    write_to_insect("곤충목록", int(a), int(b))

    doc.Content.WholeStory()
    doc.Content.Copy()
    print("ok")

    return word
