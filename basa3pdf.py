
import ui
import sys
import threading
from convert import task
import update

ui.paint()

update.install()

ui.label("Abriendo archivo...")

thread = threading.Thread(target=task, args=(sys.argv, ))
thread.start()


ui.loop()
