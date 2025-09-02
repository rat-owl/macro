# logic/job_trigger.py

from core.file_utils import select_file
from logic.excel_to_word import excel_to_word, insect_to_word
from logic.word_to_hwp import copy_to_hwp
from logic.styles import remove_superscript, 학명스타일
from logic.hwp_formatting import 표기본, 표문단설정, 국명, 필요없는것, 빈공간지우기, 곤충기본, 곤충문단, 곤충국명
from core.global_state import get_window



def excel_to_hwp():
    try:
        excel_file = select_file(title="excel을 선택하시오", filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xls;*.xlsb")])
        excel_to_word(excel_file)
        copy_to_hwp()
        window = get_window()
        window.show_message_signal.emit("엑셀에서 한글로 변환 완료")
    except Exception as e:
        print(f"오류 발생: {str(e)}")


def insect_to_hwp():
    try:
        excel_file = select_file(title="excel을 선택하시오", filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xls;*.xlsb")])
        insect_to_word(excel_file)
        window = get_window()
        window.show_message_signal.emit("엑셀에서 한글로 변환 완료")
    except Exception as e:
        print(f"오류 발생: {str(e)}")


def 표():
    try:
        표기본()
        remove_superscript()
        표문단설정()
        학명스타일("Family", "동물상,Family")
        학명스타일("Order", "동물상,Order")
        학명스타일("Class", "동물상,Class")
        국명()
        빈공간지우기()
        필요없는것()
        win = get_window()
        win.show_message_signal.emit("한글 종목록 작업 완료")
    except Exception as e:
        print(f"오류 발생: {str(e)}")


def 곤충():
    try:
        곤충기본()
        곤충문단()
        remove_superscript()
        학명스타일("Family", "동물상,Family")
        학명스타일("Order", "동물상,Order")
        곤충국명()
        빈공간지우기()
        win = get_window()
        win.show_message_signal.emit("곤충만 종목록 작업 완료")
    except Exception as e:
        print(f"오류 발생: {str(e)}")
