from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': [], 'excludes': []}

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('main.py', base=base, targetName = 'elaboracion_saldos_v0.2')
]

setup(name='Elaboracion_de_saldos',
      version = '0.2',
      description = 'Elaboracion de Saldos para Ferticampo',
      options = {'build_exe': build_options},
      executables = executables)
