# PyInstaller 스펙 파일. Windows에서 다음처럼 빌드한다:
#   pyinstaller musago_report.spec
# scripts/build_exe.bat 을 실행하면 이 스펙으로 자동 빌드된다.

block_cipher = None

a = Analysis(
    ['scripts/gui_app.py'],
    pathex=['scripts'],
    binaries=[],
    datas=[],
    hiddenimports=['calc_report_v7', 'build_report_v7', 'write_report_v7'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='무사고전환_보고서_생성기',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
