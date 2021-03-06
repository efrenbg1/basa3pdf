from src.temp import temp
from src.ui import label, confirm

checkVersion = 0.3
checkURL = "https://api.github.com/repos/efrenbg1/basa3pdf/releases/latest"
checkInterval = 24*60*60


def check():
    global version

    if previousCheck():
        return

    label("Buscando actualizaciones...")

    import requests
    r = requests.get(checkURL).json()

    if float(r["tag_name"]) <= checkVersion:
        return

    answer = confirm("Actualizar basa3pdf",
                     "Hay una nueva versión disponible. ¿Instalar ahora?")

    if answer != "yes":
        return

    install(r["assets"][0]["browser_download_url"])


def previousCheck():
    from os import path, stat, utime
    from time import time

    log = path.join(temp, ".update")

    if not path.exists(log):
        open(log, 'a').close()

    lastcheck = stat(log).st_mtime
    if lastcheck > time() - checkInterval:
        return True

    utime(log)
    return False


def install(url):
    import os
    import wget

    label("Descargando actualización...")
    out = os.path.join(temp, "basa3pdf.exe")
    wget.download(url, out=out)

    import subprocess
    label("Instalando actualización...")
    subprocess.Popen([out])
    os._exit(1)
