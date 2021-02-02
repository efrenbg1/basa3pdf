
import ui
from sys import argv
from threading import Thread
from convert import task

ui.paint()

thread = Thread(target=task, args=(argv, )).start()

ui.loop()
