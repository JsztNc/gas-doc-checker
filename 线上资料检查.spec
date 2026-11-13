# -*- mode: python ; coding: utf-8 -*-
# V1.4.2 起：多文件版（onedir）。
#   - exe 只含界面与基础设施；checker/（审核规则）不编进 exe，
#     作为外挂 .py 随包分发，改规则不用重新打包。
#   - 过渡用单文件版用 desktop? 不 —— 单文件过渡版用本 spec 的 onefile 分支
#     打包：python -m PyInstaller -f? 见 README/发版脚本。
#     实现：设置环境变量 ONEFILE=1 时走单文件分支（checker 打进 exe，供老用户自更新）。


a = Analysis(
    ['ui.py'],
    pathex=[],
    binaries=[],
    datas=[('version.txt', '.'), ('feedback_token.txt', '.'), ('使用说明.txt', '.')],
    hiddenimports=['clr', 'fitz', 'docx'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['checker'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

import os
onefile = os.environ.get("ONEFILE") == "1"

if onefile:
    # 过渡版：checker 仍打进 exe（老用户的更新器只认 exe 资产）
    a = Analysis(
        ['ui.py'],
        pathex=[],
        binaries=[],
        datas=[('version.txt', '.'), ('feedback_token.txt', '.'), ('使用说明.txt', '.')],
        hiddenimports=['clr', 'fitz', 'docx', 'checker.rules', 'checker.docx_reader',
                       'checker.finder', 'checker.fixer', 'checker.pdf_reader'],
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
        name='线上资料检查',
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
else:
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name='线上资料检查',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name='线上资料检查',
    )
