#setup.py
import sys, os
from cx_Freeze import setup, Executable
import os.path
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

__version__ = "8.1"

packages = ["pandas","numpy","xlrd","openpyxl","email","smtplib","tkinter"]

setup(
    name = "PROKARD",
    version=__version__,
    description = "Created By Students of UVPCE",
    options = {"build_exe": {
    'include_files':[
        os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
        os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
     ],'packages': packages,'include_msvcr': True
}},
executables = [Executable("PROKARD.py",base="Win32GUI")]
)
