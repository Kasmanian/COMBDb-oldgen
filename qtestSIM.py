import sys
import unittest
from PyQt5.QtWidgets import QApplication
from PyQt5.QtTest import QTest
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMainWindow

import threading
from queue import Queue

from view import *
from model import Model
from model import *
import view

app = QApplication(sys.argv)

class AdminLoginTest(unittest.TestCase):
    '''Test the Admin Login GUI'''
    def setUp(self):
        '''Create the GUI'''
        self.form = view.AdminLoginScreen(self, QMainWindow)
        #self.model = Model()
        #self.model.connect()
        #self.techLogin = self.model.techLogin

    def setFormToZero(self):
        '''Clear all fields'''
        self.form.user.setText("")
        self.form.pswd.setText("")
    
    def test_defaults(self):
        '''Test the GUI in its default state'''
        #Default state of Admin Login UI is all empty
        self.assertEqual(self.form.user.text(), "")
        self.assertEqual(self.form.pswd.text(), "")

        #Attempt to login with blank entries for username and password -  throws errorMessage
        loginButton = self.form.login
        QTest.mouseClick(loginButton, Qt.LeftButton)
        self.assertEqual(self.form.errorMessage.text(), "Please input all fields")

class thread1(threading.Thread):
    def __init__(self, thread_name, thread_ID):
        threading.Thread.__init__(self)
        self.thread_name = thread_name
        self.thread_ID = thread_ID
 
        # helper function to execute the threads
    def run(self):
        print(str(self.thread_name) +" "+ str(self.thread_ID));
        self.view = View(Model())

class thread2(threading.Thread):
    def __init__(self, thread_name, thread_ID):
        threading.Thread.__init__(self)
        self.thread_name = thread_name
        self.thread_ID = thread_ID
 
        # helper function to execute the threads
    def run(self):
        print(str(self.thread_name) +" "+ str(self.thread_ID));
        unittest.main()

if __name__ == "__main__":
    q = Queue()
    thread1 = thread1("GUIThread", 1000)
    thread2 = thread2("UnitTestThread", 2000)
    thread1.start()
    thread2.start()
    #unittest.main()