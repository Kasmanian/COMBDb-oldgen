import sys
import unittest
from PyQt5.QtWidgets import QApplication
from PyQt5.QtTest import QTest
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMainWindow

from view import *
from model import Model
from model import *

class AdminLoginTest(unittest.TestCase):
    '''Test the Admin Login GUI'''
    def setUp(self):
        '''Create the GUI'''
        self.view = View(Model(), True)

    def test_startup(self):
        self.assertTrue(isinstance(self.view.widget.currentWidget(), AdminLoginScreen))

if __name__ == "__main__":
    unittest.main()