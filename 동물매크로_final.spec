# 동물매크로_final.spec
# 빌드:  pyinstaller --noconfirm 동물매크로_final.spec

block_cipher = None

import os
from PyInstaller.utils.hooks import collect_all, collect_submodules

# 1) PySide6 / shiboken6 전체 수집 (plugins/qwindows.dll 포함)
pyside6_datas, pyside6_bins, pyside6_hidden = collect_all('PySide6')
shib6_datas,   shib6_bins,   shib6_hidden   = collect_all('shiboken6')

# 2) 우리 패키지들 하위 모듈 긁기 (동적 import 대비)
hidden_ui    = collect_submodules('ui')       # ui.*, ui.ui_main 등
hidden_hwp   = collect_submodules('hwp')      # hwp.*
hidden_logic = collect_submodules('logic')    # logic.*

# 3) qt.conf 동봉 (★ 2튜플 형식)
extra_datas = [
    ('qt.conf', '.'),           # <-- (src, dest)
    # 필요하면 아이콘/기타 데이터도 여기서 ('icon/icon.ico','icon') 같은 식으로
]

# 4) 최종 hiddenimports (명시 + 수집)
hiddenimports = list(set(
    pyside6_hidden + shib6_hidden +
    hidden_ui + hidden_hwp + hidden_logic +
    [
        'ui.ui_main',            # ui_dialog는 ui_main 안에 있음
        'hwp.hwp_api_adapter',
        # 'PySide6.QtNetwork', 'PySide6.QtQml',  # 실제 쓰면 추가
    ]
))

# 5) Analysis
a = Analysis(
    ['main.py'],  # 엔트리
    pathex=['D:\\매크로\\리팩토링 프로젝트\\다시'],  # 프로젝트 루트 절대경로
    binaries=pyside6_bins + shib6_bins,
    datas=pyside6_datas + shib6_datas + extra_datas,  # ★ 여기 2튜플만
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 진짜 안 쓰면 용량/경고 줄이려고 배제(쓰면 지워)
        'numba', 'llvmlite',
        'scipy', 'matplotlib', 'pyarrow', 'lxml', 'IPython',
        'PySide6.scripts.deploy_lib',  # 경고 잠재우고 싶으면 배제
    ],
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz, a.scripts, a.binaries, a.zipfiles, a.datas,
    [],
    name='동물매크로',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,                 # 에러 0 확인될 때까지 콘솔 ON
    icon='icon/icon.ico',         # 아이콘 경로 맞음
)

coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas,
    strip=False, upx=True, upx_exclude=[],
    name='동물매크로_final'
)
