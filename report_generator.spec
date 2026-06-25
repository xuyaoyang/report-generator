# -*- mode: python ; coding: utf-8 -*-
datas = [
    ('config/material_categories.json', 'config'),
    ('products', 'products'),
    ('ui/resources', 'ui/resources'),
    ('image_lib', 'image_lib'),
]

hiddenimports = []


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pytest',
        'unittest',
        'tkinter',
        'torch',
        'torchvision',
        'torchaudio',
        'transformers',
        'datasets',
        'pandas',
        'scipy',
        'matplotlib',
        'pyarrow',
        'cv2',
        'av',
        'bitsandbytes',
        'onnxruntime',
    ],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='融海报告生成',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
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
    name='融海报告生成',
)
