python -m venv .\env
pip3 install Pyinstaller
.\env\Scripts\activate
pyinstaller -F .\groups2sheets.py -i .\icons8-split-40.ico -n groups2sheets-header.exe --clean -c --hidden-import openpyxl --hidden-import os --hidden-import functools --hidden-import time --hidden-import configparser --hidden-import chardet
