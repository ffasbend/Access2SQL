# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = [('build/mdb_staging/bin/mdb-export', 'mdbtools/bin'), ('build/mdb_staging/bin/mdb-prop', 'mdbtools/bin'), ('build/mdb_staging/bin/mdb-queries', 'mdbtools/bin'), ('build/mdb_staging/bin/mdb-schema', 'mdbtools/bin'), ('build/mdb_staging/bin/mdb-sql', 'mdbtools/bin'), ('build/mdb_staging/bin/mdb-tables', 'mdbtools/bin'), ('build/mdb_staging/lib/libglib-2.0.0.dylib', 'mdbtools/lib'), ('build/mdb_staging/lib/libintl.8.dylib', 'mdbtools/lib'), ('build/mdb_staging/lib/libmdb.3.dylib', 'mdbtools/lib'), ('build/mdb_staging/lib/libmdbsql.3.dylib', 'mdbtools/lib'), ('build/mdb_staging/lib/libpcre2-8.0.dylib', 'mdbtools/lib'), ('build/mdb_staging/lib/libreadline.8.dylib', 'mdbtools/lib')]
hiddenimports = []
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['access2sql_gui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    [],
    exclude_binaries=True,
    name='Access2SQL',
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
    name='Access2SQL',
)
app = BUNDLE(
    coll,
    name='Access2SQL.app',
    icon=None,
    bundle_identifier=None,
)
