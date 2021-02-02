import os
import tempfile
from src.ui import fatal, label

tempHalfLife = 24*60*60

try:
    temp = os.path.join(tempfile.gettempdir(), "basa3pdf")
    if not os.path.exists(temp):
        os.makedirs(temp)
except Exception as e:
    fatal("No se puede abrir el directorio temporal:", e=e)


def clean():
    label("Limpiando archivos viejos...")
    import time

    now = time.time()
    for f in os.listdir(temp):
        f = os.path.join(temp, f)
        if os.stat(f).st_mtime < now - tempHalfLife:
            try:
                os.remove(f)
            except:
                pass
