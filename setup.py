from cx_Freeze import setup, Executable
import os
import sys
# Dependencies are automatically detected, but it might need
# fine tuning.
os.environ['TCL_LIBRARY'] = r'C:\Python\Python36\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Python\Python36\tcl\tk8.6'

Include_Modules = [
    "numpy.core._methods", "numpy.lib.format"
]

buildOptions = {"includes": Include_Modules, "include_files": ["tcl86t.dll", "tk86t.dll", "proteinsimple_logo_bt.ico"]}

base = 'Win32GUI' if sys.platform == 'win32' else None

executables = [
    Executable('formatCSV.py', base=base)
]

setup(name='formatCSV',
      version='1.0',
      description='Description',
      options=dict(build_exe=buildOptions),
      executables=executables)
