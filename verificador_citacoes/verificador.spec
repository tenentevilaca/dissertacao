# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec para VerificadorCitacoes.exe

block_cipher = None

a = Analysis(
    ['iniciar.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('static',  'static'),   # interface web
        ('modules', 'modules'),  # módulos Python
    ],
    hiddenimports=[
        # uvicorn
        'uvicorn',
        'uvicorn.logging',
        'uvicorn.loops',
        'uvicorn.loops.auto',
        'uvicorn.protocols',
        'uvicorn.protocols.http',
        'uvicorn.protocols.http.auto',
        'uvicorn.protocols.websockets',
        'uvicorn.protocols.websockets.auto',
        'uvicorn.lifespan',
        'uvicorn.lifespan.on',
        # fastapi / starlette
        'fastapi',
        'starlette',
        'starlette.routing',
        'starlette.middleware',
        'starlette.responses',
        'starlette.requests',
        # async
        'anyio',
        'anyio._backends._asyncio',
        'anyio._backends._trio',
        # anthropic / httpx
        'anthropic',
        'httpx',
        'httpcore',
        # docx
        'docx',
        'lxml',
        'lxml.etree',
        # fitz (PyMuPDF)
        'fitz',
        # outros
        'multipart',
        'email.mime',
        'email.mime.multipart',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'numpy', 'pandas', 'scipy'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zlib_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='VerificadorCitacoes',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,       # janela de terminal visível (mostra progresso)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
