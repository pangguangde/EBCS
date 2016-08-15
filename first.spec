# -*- mode: python -*-

block_cipher = None


a = Analysis(['first.py'],
             pathex=['/Users/pangguangde/Desktop/\xe6\xa2\x81\xe4\xb8\xbd\xe9\x9c\x9e\xe5\xaf\xb9\xe8\xb4\xa6/\xe7\x95\x8c\xe9\x9d\xa2\xe7\x89\x88'],
             binaries=None,
             datas=None,
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='first',
          debug=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='first')
