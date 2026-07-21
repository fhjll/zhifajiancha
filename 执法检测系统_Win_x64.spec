# -*- mode: python ; coding: utf-8 -*-
import sys
sys.setrecursionlimit(5000)

from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all

datas = [('事实认定书模板.docx', '.'), ('使用说明.md', '.'), ('column_mapping.json', '.')]
binaries = []
hiddenimports = ['pandas', 'openpyxl', 'openpyxl.cell._writer', 'openpyxl.reader.excel', 'openpyxl.workbook', 'openpyxl.writer.excel', 'docx', 'lxml', 'lxml.html', 'customtkinter', 'PIL', 'PIL._tkinter_finder', 'openai', 'httpx', 'httpx._transports.default', 'httpcore', 'sniffio', 'ctypes', 'ctypes.util', 'numpy', 'numpy.core._methods', 'numpy.lib.format', 'numpy.random', 'numpy.random.common', 'numpy.random.bounded_integers', 'preprocess', 'generate_report', 'process_zero_balance', 're', 'threading', 'concurrent', 'concurrent.futures', 'io', 'tkinter', 'tkinter.filedialog', 'tkinter.messagebox', 'xml', 'xml.etree', 'xml.etree.ElementTree']
hiddenimports += collect_submodules('openpyxl')
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pandas')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


block_cipher = None


a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'PyQt5', 'notebook', 'IPython', 'jupyter', 'bokeh', 'plotly', 'test', 'unittest'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='执法检测系统',
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
