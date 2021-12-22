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
    "tkinter",
    "win32com",
    "wx",
    "zmq",
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
    version = '0.1',
    description = 'converter',
    executables = [exe])
