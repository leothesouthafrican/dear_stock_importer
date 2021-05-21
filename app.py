from functions import *

import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class ExampleHandler(FileSystemEventHandler):
    def on_created(self, event): # when file is created
        main()
        print("It worked")

observer = Observer()
event_handler = ExampleHandler() # create event handler
# set observer to use created handler in directory
observer.schedule(event_handler, path='input_folder/')
observer.start()

# sleep until keyboard interrupt, then stop + rejoin the observer
try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()

observer.join()
