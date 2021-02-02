from src.ui import fatal, label
from src.temp import temp, clean
import src.update as update


def task(file):
    try:
        update.check()
    except Exception as e:
        print(e)
        pass

    label("Abriendo archivo...")

    import os

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
            outFile = directory[k+1:j] + directory[j+5:] + ".pdf"
            outFile = os.path.join(temp, outFile)
            docx(file, outFile)
        else:
            fatal("No se encuentra el archivo")
        os.startfile(outFile, 'open')
        clean()
        os._exit(1)
    except Exception as e:
        fatal("No se puede abrir el archivo:", e)


def docx(inFile, outFile):
    from win32com import client
    import pythoncom

    pythoncom.CoInitialize()

    word = client.Dispatch("Word.Application")

    doc = word.Documents.Open(inFile)

    doc.SaveAs(outFile, FileFormat=17)
    doc.Close()


def xlsx(inFile, outFile):
    from win32com import client
    import pythoncom

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
