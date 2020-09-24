
import tkinter as tk
from tkinter.ttk import Progressbar
import threading
import sys
import shutil
import os
import traceback
from tkinter import messagebox
from docx2pdf import convert
import pythoncom

done = False
title = "Gestión | Conversor PDF"


def task(file):
    global done, title

    pythoncom.CoInitialize()

    if not file.endswith('.basa'):
        done = True
        messagebox.showerror(
            message="No se encuentra el archivo", title=title)
        os._exit(1)

    try:
        directory = file[:file.rindex('.basa')]
        docx = directory + ".docx"
        pdf = directory + ".pdf"

        shutil.copy2(file, docx)
        convert(docx, output_path=pdf)
        os.remove(docx)
        os.startfile(pdf, 'open')
        os._exit(1)
    except Exception as e:
        done = True
        traceback.print_exc()
        messagebox.showerror(
            message="No se puede abrir el archivo:\n\n" + str(e), title=title)
        os._exit(1)


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
