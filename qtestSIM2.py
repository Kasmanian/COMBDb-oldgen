
from queue import Queue
from threading import Thread
import time

from view import View
from model import Model

from view import *
from model import Model
from model import *

import sys
from PyQt5.QtWidgets import QApplication
  
app = QApplication(sys.argv)

def main():
    view = View(Model())
    handleThread()

class handleThread():
    x = 0
    def __init__(self, view):
        self.view = view

    # A thread that produces data
    def producer(out_q):
        #time.sleep(1)
        #print(isinstance(view.widget.currentWidget(), AdminLoginScreen))
        #self.form = view.AdminLoginScreen(self, QMainWindow)
        #self.form.user.setText("MAIK IS THE BEST")
        while True:
        # Produce some data
            for x in range(10):
                time.sleep(0.5)
                out_q.put(time.time())
                x += 1
            break
            
    # A thread that consumes data
    def consumer(in_q):
        while True:
            # Get some data
            data = in_q.get()
            # Process the data
            print(data)
            
    # Create the shared queue and launch both threads
    q = Queue()
    t1 = Thread(target = consumer, args =(q, ))
    t2 = Thread(target = producer, args =(q, ))
    t1.start()
    t2.start()

if __name__=="__main__":
    main()