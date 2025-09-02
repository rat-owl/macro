# --- main.py (수정본) ---
import os
import sys
import ctypes
import traceback

# 1) (권장) PyInstaller 환경에서 Qt 플러그인 경로 먼저 잡아주기
def _patch_qt_plugin_path():
    """
    EXE(동결) 상태면 _MEIPASS 아래의 PySide6/Qt/plugins를 우선 사용.
    개발(소스) 상태면 site-packages의 PySide6 경로를 사용.
    """
    try:
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            base = sys._MEIPASS  # PyInstaller가 임시로 푸는 경로
            qt_plugins = os.path.join(base, 'PySide6', 'Qt', 'plugins')
        else:
            # 개발 환경: 설치된 PySide6 위치에서 plugins 경로 추적
            import PySide6
            qt_plugins = os.path.join(os.path.dirname(PySide6.__file__), 'Qt', 'plugins')

        # 환경변수 + Qt 라이브러리 경로 모두 설정
        os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = qt_plugins  # 주로 사용하는 변수
        os.environ['QT_PLUGIN_PATH'] = qt_plugins               # 보험

        # Qt가 읽을 qt.conf가 있다면 dist 옆 폴더에 두고 이 줄은 생략 가능
        # QCoreApplication.libraryPaths 를 건드리려면 Qt 임포트 후에 다시 세팅
        return qt_plugins
    except Exception:
        return None

# 2) 메인 스레드에서 COM을 STA로 ‘먼저’ 고정 (PySide6/pywinauto 보다 앞에서!)
def _init_com_sta_once():
    COINIT_APARTMENTTHREADED = 0x2  # STA
    hr = ctypes.windll.ole32.CoInitializeEx(None, COINIT_APARTMENTTHREADED)
    # S_OK(0) / S_FALSE(1)이면 이미/성공 초기화. 다른 값이면 로깅만.
    # RPC_E_CHANGED_MODE(0x80010106)도 여기서 한 번 먹고 지나가면 보통 안정화됨.
    return hr

def main():
    # 먼저 경로/COM 고정
    qt_plugins = _patch_qt_plugin_path()
    hr = _init_com_sta_once()

    # 여기서부터 Qt 임포트 (COM/경로 세팅 이후로 미룬다)
    from PySide6 import QtWidgets, QtCore

    # (선택) Qt 라이브러리 경로에도 plugins 주입
    try:
        from PySide6.QtCore import QCoreApplication
        paths = QCoreApplication.libraryPaths()
        if qt_plugins and qt_plugins not in paths:
            QCoreApplication.setLibraryPaths([qt_plugins] + paths)
    except Exception:
        pass

    # 이제 우리 코드 임포트 (UI 생성 시 내부에서 pywinauto를 만지더라도 COM은 이미 STA)
    from ui.ui_main import MainWindow
    from core.global_state import set_window

    # 앱 생성
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)

    # 디버깅 정보
    print(f"Plugin path: {qt_plugins}")
    try:
        print(f"Qt version: {QtCore.__version__}")
    except Exception:
        pass
    # print(f"CoInitializeEx HR: 0x{hr & 0xffffffff:08X}")

    # 윈도우 생성/표시
    window = MainWindow()
    set_window(window)
    window.show()

    try:
        sys.exit(app.exec())
    except Exception as e:
        print(f"❌ 앱 실행 중 오류: {e}")
        traceback.print_exc()
        # COM 해제는 프로세스 종료 시 자동이지만 필요하면 아래 주석 해제
        # ctypes.windll.ole32.CoUninitialize()

if __name__ == '__main__':
    # 멀티프로세싱 쓸 때 동결 안전장치
    import multiprocessing as mp
    mp.freeze_support()

    try:
        main()
    except Exception as e:
        print("❌ 실행 중 오류 발생:", str(e))
        traceback.print_exc()
        input("엔터를 누르면 종료됩니다...")
