import requests
import wget
import subprocess
import os
from temp import temp
from ui import label, confirm, spin

version = 0.1


def check():
    global version
    r = requests.get(
        "https://api.github.com/repos/efrenbg1/basa3pdf/releases/latest")
    r = r.json()
    if float(r["tag_name"]) > version:
        label("Descargando actualización...")
        url = r["assets"][0]["browser_download_url"]
        out = os.path.join(temp, r["tag_name"] + ".exe")
        print(out)
        wget.download(url, out=out)


def install():
    label("Buscando actualizaciones...")
    for f in os.listdir(temp):
        if not f.endswith(".exe"):
            continue
        newv = -1
        try:
            newv = float(f[:-4])
        except:
            continue

        if newv <= version:
            continue

        answer = confirm("Actualizar basa3pdf",
                         "Hay una nueva versión disponible. ¿Instalar ahora?")
        if answer == "yes":
            exe = os.path.join(temp, f)
            subprocess.Popen([exe])
            os._exit(1)
