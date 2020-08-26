# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['WEM_Mapping.py'],
             pathex=['C:\\Users\\Accounting\\PycharmProjects\\WEM_Mapping'],
             binaries=[],
             datas=[],
             hiddenimports=['PySimpleGUI', 'numpy', 'pandas', 're', 'time', 'selenium', 'os', 'shutil', 'glob', 'playsound', 'rpy2'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='WEM_Mapping',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False , icon='WEM_M_Icon.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='WEM_Mapping')
