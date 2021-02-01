python -m venv env
env\Scripts\activate.bat
pip3 install pyinstaller
pip3 install docx2pdf
copy hook-docx2pdf.py env\Lib\site-packages\PyInstaller\hooks