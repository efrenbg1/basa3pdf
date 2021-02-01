import os
import tempfile
import time
from ui import fatal, label

try:
    temp = os.path.join(tempfile.gettempdir(), "basa3pdf")
    if not os.path.exists(temp):
        os.makedirs(temp)
except Exception as e:
    fatal("No se puede abrir el directorio temporal:", e=e)


def clean():
    label("Limpiando archivos viejos...")
    now = time.time()
    for f in os.listdir(temp):
        f = os.path.join(temp, f)
        if os.stat(f).st_mtime < now - 2*60:
            try:
                os.remove(f)
            except:
                pass
