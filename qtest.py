import sys
import unittest
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtTest import QTest
from PyQt5.QtCore import Qt

import view
from model import *

app = QApplication(sys.argv)

class AdminLoginTest(unittest.TestCase):
    '''Test the Admin Login GUI'''
    def setUp(self):
        '''Create the GUI'''
        self.form = view.AdminLoginScreen(self, QMainWindow)
        self.model = Model()
        self.model.connect()
        self.techLogin = self.model.techLogin
        #self.view = View(Model())

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

    def test_usernameInput(self):
        '''Test the GUI with input in Username only'''
        self.setFormToZero()
        self.form.user.setText("Reisdorf")
        self.assertEqual(self.form.pswd.text(), "")

        #Attempt to login with username input, but blank password - throws errorMessage
        loginButton = self.form.login
        QTest.mouseClick(loginButton, Qt.LeftButton)
        self.assertEqual(self.form.errorMessage.text(), "Please input all fields")

    def test_passwordInput(self):
        '''Test the GUI with input in Password only'''
        self.setFormToZero()
        self.form.pswd.setText("Password")
        self.assertEqual(self.form.user.text(), "")

        #Attempt to login with password input, but blank username - throws errorMessage
        loginButton = self.form.login
        QTest.mouseClick(loginButton, Qt.LeftButton)
        self.assertEqual(self.form.errorMessage.text(), "Please input all fields")
    
    def test_incorrectUsernameInput(self):
        '''Test the GUI with wrong username only'''
        self.setFormToZero()
        self.form.user.setText("NotAUser")
        self.form.pswd.setText("Password")

        #Attempt to login with username and password, but the username is wrong - throws errorMessage
        self.assertEqual(self.model.techLogin(self.form.user.text(), self.form.pswd.text()), False)

    def test_incorrectPasswordInput(self):
        '''Test the GUI with wrong password only'''
        self.setFormToZero()
        self.form.user.setText("Reisdorf")
        self.form.pswd.setText("WrongPassword")

        #Attempt to login with username and password, but the password is wrong - throws errorMessage
        self.assertEqual(self.model.techLogin(self.form.user.text(), self.form.pswd.text()), False)
    """
    def test_correctLogin(self):
        '''Test the GUI with correct login'''
        self.setFormToZero()
        self.form.user.setText("Reisdorf")
        self.form.pswd.setText("Password")

        #Attempt to login with correct username and password - will change screen to Admin Home Screen
        loginButton = self.form.login
        QTest.mouseClick(loginButton, Qt.LeftButton)
        self.assertEqual(AdminHomeScreen.isActiveWindow, True)
    """
if __name__ == "__main__":
    unittest.main()