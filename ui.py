import tkinter as tk
from tkinter import messagebox
from tkinter.ttk import Progressbar
import os
import sys
import traceback

title = "basa 3.0 | Conversor PDF"
spin = True

# Start window and set icon
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

_label = None


def paint():
    global _label

    # Top space
    tpre = tk.Label(root, text='')
    tpre.pack()

    # Main text
    t = tk.Label(root, text="Cargando...", font=('helvetica', 12, 'bold'))
    t.pack()
    _label = t

    # Spinner
    p = Progressbar(root, length=200,
                    mode="indeterminate", takefocus=True, maximum=100)
    p.pack()

    # Bottom space
    tpost = tk.Label(root, text='')
    tpost.pack()

    def spinner():
        global spin
        if spin:
            p.step()
        root.after(10, spinner)

    root.after(10, spinner)


def label(msg):
    _label.config(text=msg)


def loop():
    root.mainloop()


def fatal(msg, e=None):
    global spin
    spin = False
    if e != None:
        traceback.print_exc()
        msg += "\n\n" + str(e)
    messagebox.showerror(message=msg, title=title)
    os._exit(1)


def confirm(title, msg):
    global spin
    spin = False
    answer = messagebox.askquestion(title, msg, icon='question')
    spin = True
    return answer
