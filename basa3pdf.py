
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

done = False
title = "Gestión | Conversor PDF"


def task(file):
    global done, title

    pythoncom.CoInitialize()
    excel = client.Dispatch("Excel.Application")

    if not file.endswith('.basa'):
        dialog("No se encuentra el archivo")

    try:
        directory = file[:file.rindex('.basa')]
    except Exception as e:
        dialog("No se puede abrir el archivo:", e)

    try:
        i = directory.find(".xlsx")
        j = directory.find(".docx")
        if i > -1:
            pdf = directory[:i] + directory[i+5:] + ".pdf"
            clean(pdf)
            wb = excel.Workbooks.Open(file)
            ws = wb.Worksheets[0]
            ws.Visible = 1
            ws.ExportAsFixedFormat(0, pdf)
            wb.Close()
            excel.Quit()
        elif j > -1:
            docx = directory[:j] + directory[j+5:] + ".docx"
            pdf = directory[:j] + directory[j+5:] + ".pdf"
            clean(pdf)
            shutil.copy2(file, docx)
            convert(docx, output_path=pdf)
            os.remove(docx)
        else:
            dialog("No se encuentra el archivo")
        os.startfile(pdf, 'open')
        os._exit(1)
    except Exception as e:
        excel.Quit()
        dialog("No se puede abrir el archivo:", e)


def dialog(msg, e=None):
    done = True
    if e != None:
        traceback.print_exc()
        msg += "\n\n" + str(e)
    messagebox.showerror(message=msg, title=title)
    os._exit(1)


def clean(pdf):
    try:
        os.remove(pdf)
        return
    except:
        pass
    try:
        stat = os.stat(file + ".pdf")
        if stat:
            dialog("¡Hay un archivo con el mismo nombre!")
    except:
        pass


root = tk.Tk()
root.title(title)

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
