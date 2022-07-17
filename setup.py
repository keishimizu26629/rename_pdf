import sys
from cx_Freeze import setup, Executable
import os

base = None

if sys.platform == 'win32':
    base = 'Win32GUI'

Icon = os.getcwd() + '\images\icon.ico'

includes = [
    "os",
    "glob",
    "re",
    "math",
    "copy",
    "pdfminer.pdfinterp",
    "pdfminer.converter",
    "pdfminer.pdfpage",
    "pdfminer.layout",
    "csv",
    "datetime",
    "dateutil.relativedelta",
    "io",
    "PyPDF2",
    "reportlab.pdfgen",
    "reportlab.pdfbase",
    "reportlab.pdfbase.cidfonts",
    "reportlab.pdfbase.ttfonts",
    "tkinter",
]

excludes = [
    "numpy",
    "pandas",
    "lxml",
    "locket",
    "PyQt5",
    "matplotlib",
    "babel",
    "bcrypt",
    "brotli",
    "debugpy",
    "jedi",
    "markupsafe",
    "nacl",
    "PIL",
    "notebook",
    "pygments",
    "sphinx",
    "tornado",
    "win32com",
    "wx",
    "zmq",
    # "tkinter",
]


# sample.pyをexe化します
exe = Executable(script = "renamePdf2.py",base= base,icon=Icon)


setup(name = 'renamePdf',
    options = {
        "build_exe": {
            "includes": includes,
            "excludes": excludes,
        }
    },
    version = '0.4',
    description = 'converter',
    executables = [exe])
