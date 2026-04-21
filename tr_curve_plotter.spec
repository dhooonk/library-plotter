# -*- mode: python ; coding: utf-8 -*-

import sys
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

# 프로그램 구동에 필수적인 외부 라이브러리 (Hidden imports)
# 동적 로딩으로 인해 PyInstaller가 자동으로 감지하지 못할 수 있는 패키지들을 명시합니다.
hidden_imports = [
    'pandas',
    'openpyxl',
    'matplotlib',
    'scipy',
    'sklearn',
    'sklearn.metrics',
]

# 만약 패키지 내부의 데이터 파일이 필요하다면 아래와 같이 수집할 수 있습니다.
datas = []
# 추가 예시: datas += collect_data_files('matplotlib')

a = Analysis(
    ['main.py'],                 # 메인 실행 파일
    pathex=['.'],                # core/, utils/ 패키지 탐색을 위해 프로젝트 루트 포함
    binaries=[],                 # 포함할 동적 라이브러리(dll, so 등)
    datas=datas,                 # 포함할 일반 데이터 파일(이미지, 텍스트 등)
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],                 # 빌드에서 제외할 불필요한 모듈
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 하나의 폴더(디렉토리) 형태로 출력물을 만들지,
# 하나의 단일 실행 파일(.exe)로 만들지에 따라 아래 설정이 달라집니다.
# -------------------------------------------------------------
# [옵션 1] 단일 파일 형태 (Onefile) 로 빌드하는 경우
# (아래 설정 사용시 배포는 쉽지만, 로딩 속도가 다소 느릴 수 있습니다.)
# -------------------------------------------------------------
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='tr_curve_plotter',  # 생성될 실행 파일 이름 (.exe)
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,                   # UPX 압축 사용 여부 (용량 최적화)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,              # 콘솔 창(CMD)을 표시할지 여부 (GUI 앱이므로 False)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='app_icon.ico',      # (선택) 아이콘 파일 경로 지정
)

# -------------------------------------------------------------
# [옵션 2] 폴더 형태 (Onedir) 로 빌드하는 경우
# 단일 파일(EXE) 빌드가 목적이므로 아래 부분은 주석 처리합니다.
# 만약 폴더 형태로 빌드하고 싶다면 위의 EXE에서 a.binaries~a.datas를 빼고,
# 아래 COLLECT 블록의 주석을 해제하세요.
# -------------------------------------------------------------
# coll = COLLECT(
#     exe,
#     a.binaries,
#     a.zipfiles,
#     a.datas,
#     strip=False,
#     upx=True,
#     upx_exclude=[],
#     name='tr_curve_plotter',
# )
