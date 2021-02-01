
import tkinter as tk
from tkinter.ttk import Progressbar
import threading
import sys
import shutil
import os
import traceback
from tkinter import messagebox
from docx2pdf import convert
from win32com import client
import pythoncom
import tempfile
import time

done = False
title = "basa 3.0 | Conversor PDF"
temp = ""


def task(file):
    global done, title, temp

    if not file.endswith('.basa'):
        fatal("No se encuentra el archivo")

    try:
        directory = file[:file.rindex('.basa')]
    except Exception as e:
        fatal("No se puede abrir el archivo:", e)

    temp = os.path.join(tempfile.gettempdir(), "basa3pdf")
    if not os.path.exists(temp):
        os.makedirs(temp)

    try:
        i = directory.find(".xlsx")
        j = directory.find(".docx")
        k = directory.rfind("\\")
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
        os._exit(1)
    except Exception as e:
        fatal("No se puede abrir el archivo:", e)


def docx(inFile, outFile):
    pythoncom.CoInitialize()
    convert(inFile, output_path=outFile)


def xlsx(inFile, outFile):
    pythoncom.CoInitialize()
    excel = client.Dispatch("Excel.Application")

    try:
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
        excel.Quit()
    except Exception as e:
        raise e  # Let the main thread handle the Exception
    finally:
        excel.Quit()


def fatal(msg, e=None):
    global done
    done = True
    if e != None:
        traceback.print_exc()
        msg += "\n\n" + str(e)
    messagebox.showerror(message=msg, title=title)
    os._exit(1)


def clean():
    now = time.time()
    for f in os.listdir(temp):
        f = os.path.join(temp, f)
        if os.stat(f).st_mtime < now - 2*60:
            try:
                os.remove(f)
            except:
                pass


root = tk.Tk()
root.title(title)

path = ""
if getattr(sys, 'frozen', False):
    path = os.path.dirname(sys.executable)
elif __file__:
    path = os.path.dirname(__file__)
icon = os.path.join(path, 'app.ico')

root.iconbitmap(icon)
root.geometry("250x100")


tpre = tk.Label(root, text='')
tpre.pack()

t = tk.Label(root, text='Abriendo PDF...', font=('helvetica', 12, 'bold'))
t.pack()

p = Progressbar(root, length=200,
                mode="indeterminate", takefocus=True, maximum=100)
p.pack()

tpost = tk.Label(root, text='')
tpost.pack()


def update():
    p.step()
    if not done:
        root.after(10, update)


if len(sys.argv) < 2:
    messagebox.showerror(
        message="No se encuentra el archivo", title=title)
    os._exit(1)

root.after(10, update)
thread = threading.Thread(target=task, args=(sys.argv[1], ))
thread.start()
root.mainloop()
