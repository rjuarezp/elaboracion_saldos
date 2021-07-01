from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': [], 'excludes': []}

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('main.py', base=base, targetName = 'Saldos')
]

setup(name='Saldos',
      version = '1.0',
      description = 'Saldos para Marien',
      options = {'build_exe': build_options},
      executables = executables)
