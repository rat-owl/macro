# core/thread_runner.py

import threading
import pythoncom

IS_WORKING = False  # 전역 상태값이므로 추후 global_state로 이동 예정

def run_in_thread(func):
    global IS_WORKING

    def wrapper():
        global IS_WORKING
        if IS_WORKING:
            return

        try:
            IS_WORKING = True
            pythoncom.CoInitialize()
            func()
        except Exception as e:
            print(f"오류 발생: {str(e)}")
        finally:
            IS_WORKING = False
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    thread = threading.Thread(target=wrapper)
    thread.daemon = True
    thread.start()
