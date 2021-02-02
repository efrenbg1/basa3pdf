from sys import argv
from threading import Thread
from src import ui, convert

ui.paint()

thread = Thread(target=convert.task, args=(argv, )).start()

ui.loop()
