pyinstaller show_filesize.pyw --onefile --noconsole

nuitka --mingw64 --follow-imports --onefile --enable-plugin=tk-inter --disable-console .\show_filesize.pyw

py -m pip install chardet
py -m pip install openpyxl
py -m pip install python-docx
py -m pip install pdfminer.six


py -m pip freeze > requirements.txt
py -m pip install -r requirements.txt
