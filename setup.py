from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': [], 'excludes': []}

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('main.py', base=base)
]

setup(name='elaboracion_saldos',
      version = '0.2',
      description = 'Elaboracion de saldos',
      options = {'build_exe': build_options},
      executables = executables)
