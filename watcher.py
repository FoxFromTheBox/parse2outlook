import sys
import time
import logging, os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

def rename():
   
    folder = "input"
    for count, filename in enumerate(os.listdir(folder)):
        dst = f"input.xlsx"
        src =f"{folder}/{filename}"  # foldername/filename, if .py file is outside folder
        dst =f"{folder}/{dst}"
         
        # rename() function will
        # rename all the files
        os.rename(src, dst)

class Watcher:

    DirectoryToWatch = r'C:\Users\RakitinIS\Documents\parser\input'

    def __init__(self):
        self.observer = Observer()
    def run(self):
        event_handler = Handler()
        self.observer.schedule(event_handler, self.DirectoryToWatch, recursive=True)
        self.observer.start()
        try:
            while True:
                time.sleep(5)
        except:
            self.observer.stop()
            print ("Error")
            self.observer.join()

class Handler(FileSystemEventHandler):

    @staticmethod
    def on_any_event(event):
        if event.is_directory:
            return None

        elif event.event_type == 'modified':
            # Take any action here when a file is first created.
            rename()
            os.system('python parserV2.py')
            sys.exit
            time.sleep(5)

if __name__ == '__main__':
    w = Watcher()
    w.run()