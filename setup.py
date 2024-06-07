import sys
from cx_Freeze import setup, Executable

base = None

if sys.platform == 'win32':
    base = 'Win32GUI'

script = "renamePdf.py"
icon = "icons/icon.ico"


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
    "PIL",
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
executables = [
    Executable(
        script=script,
        base=base,
        icon=icon,  # アイコンファイルを指定します
        copyright="(C) 2022 Keisuke Shimizu"
    )
]

build_exe_options = {
    "includes": includes,  # 必要なパッケージをここに追加
    "excludes": excludes,  # 除外するパッケージをここに追加
    "include_files": [
        icon,
        "resource.res"
    ],
}

setup(name = 'renamePdf',
    version = '3.0',
    description = 'A dedicated application for processing PDFs generated when ordering system baths through the PU Order System. This application reads site information from the text within PDFs and renames the files based on this information. It also calculates the deadline for finalizing the shipping date based on the shipping information and records it within the PDF. Additionally, it performs other PDF manipulations to streamline workflow processes.',
    options={"build_exe": build_exe_options},
    executables = executables
)
