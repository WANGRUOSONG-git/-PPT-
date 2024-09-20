# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('/Users/russell/Documents/PYTHON/PPT/resources/user_manual.pdf', 'resources'), ('/Users/russell/Documents/PYTHON/PPT/resources/example_PPT.zip', 'resources')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='麦子PPT翻译',
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
    icon=['/Users/russell/Documents/PYTHON/PPT/icon.ico'],
)
app = BUNDLE(
    exe,
    name='麦子PPT翻译.app',
    icon='/Users/russell/Documents/PYTHON/PPT/icon.ico',
    bundle_identifier=None,
)
