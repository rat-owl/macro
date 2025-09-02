import psutil
from PySide6.QtWidgets import QFileDialog
from pywinauto import Application


def select_file(title, filetypes):
    file_path, _ = QFileDialog.getOpenFileName(None, title, "", filetypes[0][1])
    return file_path


def focus_on(w_file):
    try:
        app = Application(backend="uia").connect(title_re=w_file)
        window = app.window(title_re=w_file)
        window.set_focus()
        return True
    except Exception as e:
        print(f"오류: {e}")
        return False


def is_program_running(program_name):
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if program_name.lower() in proc.info['name'].lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
    return False