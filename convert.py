from docx2pdf import convert
from win32com import client
import pythoncom
import shutil
import os
from ui import fatal, label
from temp import temp, clean


def task(file):
    if len(file) < 2:
        fatal("No se encuentra el archivo")
    file = file[1]

    if not file.endswith('.basa'):
        fatal("No se encuentra el archivo")

    try:
        directory = file[:file.rindex('.basa')]
    except Exception as e:
        fatal("No se puede abrir el archivo:", e)

    try:
        i = directory.find(".xlsx")
        j = directory.find(".docx")
        k = directory.rfind("\\")
        label("Generando PDF...")
        if i > -1:
            outFile = directory[k+1:i] + directory[i+5:] + ".pdf"
            outFile = os.path.join(temp, outFile)
            xlsx(file, outFile)
        elif j > -1:
            inFile = directory[k+1:j] + directory[j+5:] + ".docx"
            inFile = os.path.join(temp, inFile)
            outFile = directory[k+1:j] + directory[j+5:] + ".pdf"
            outFile = os.path.join(temp, outFile)
            shutil.copy2(file, inFile)
            docx(inFile, outFile)
        else:
            fatal("No se encuentra el archivo")
        os.startfile(outFile, 'open')
        clean()
        import update
        update.check()
        os._exit(1)
    except Exception as e:
        fatal("No se puede abrir el archivo:", e)


def docx(inFile, outFile):
    pythoncom.CoInitialize()
    convert(inFile, output_path=outFile)


def xlsx(inFile, outFile):
    pythoncom.CoInitialize()
    excel = client.Dispatch("Excel.Application")

    wb = excel.Workbooks.Open(inFile)
    ws = wb.Worksheets[0]

    # Fix title and visibility
    ws.Visible = 1
    if ws.PageSetup.PrintTitleRows == "":
        ws.PageSetup.PrintTitleRows = "$4:$4"
    wb.Saved = True

    # Convert to PDF
    ws.ExportAsFixedFormat(0, outFile)
    wb.Close()
