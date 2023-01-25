import sys
from cx_Freeze import setup, Executable

base = None

if sys.platform == 'win32':
    base = 'Win32GUI'


includes = [
    "os",
    "glob",
    "re",
    "math",
    "copy",
    "pdfminer",
    "csv",
    "datetime",
    "dateutil",
    "io",
    "PyPDF2",
    "reportlab",
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
exe = Executable(script = "renamePdf.py", base= base)


setup(name = 'renamePdf',
    options = {
        "build_exe": {
            "includes": includes,
            "excludes": excludes,
        }
    },
    version = '0.2',
    description = 'converter',
    executables = [exe])
