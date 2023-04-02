import sys
import unittest
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtTest import QTest
from PyQt5.QtCore import Qt

import sys
sys.path.insert(0, r'C:\Users\Hoboburger\Desktop\COMBDb\src')
 
from QView import QView
from QModel import QModel
from Utility.QAdminLogin import QAdminLogin
from Utility.QAdminHome import QAdminHome

app = QApplication(sys.argv)

class AdminLoginTest(unittest.TestCase):
    '''Test the Admin Login GUI'''
    def setUp(self):
        '''Create the GUI'''
        self.model = QModel()
        self.model.connect()
        self.form = QAdminLogin(self.model, self)
        self.techLogin = self.model.techLogin

    def setFormToZero(self):
        '''Clear all fields'''
        self.form.user.setText("")
        self.form.pswd.setText("")

    def test_defaults(self):
        '''Test the GUI in its default state'''
        #Default state of Admin Login UI is all empty
        self.assertTrue("Welcome to COMBDb" in self.form.title.text())
        self.assertTrue("Admin Login" in self.form.description.text())
        self.assertTrue("Username:" in self.form.usernameTag.text())
        self.assertEqual(self.form.user.text(), "")
        self.assertTrue("Password:" in self.form.passwordTag.text())
        self.assertEqual(self.form.pswd.text(), "")
        self.assertEqual(self.form.login.text(), " Login")

        #Attempt to login with blank entries for username and password -  throws errorMessage
        QTest.mouseClick(self.form.login, Qt.LeftButton)
        self.assertEqual(self.form.errorMessage.text(), "Please input all fields")

    def test_usernameInput(self):
        '''Test the GUI with input in Username only'''
        #Start with blank form
        self.setFormToZero()

        self.form.user.setText("Reisdorf")
        self.assertEqual(self.form.pswd.text(), "")

        #Attempt to login with username input, but blank password - throws errorMessage
        QTest.mouseClick(self.form.login, Qt.LeftButton)
        self.assertEqual(self.form.errorMessage.text(), "Please input all fields")

    def test_passwordInput(self):
        '''Test the GUI with input in Password only'''
        #Start with blank form
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

    def test_correctLogin(self):
        '''Test the GUI with correct login'''
        self.setFormToZero()
        self.form.user.setText("Reisdorf")
        self.form.pswd.setText("Password")

        #Attempt to login with valid username and password - returns true
        self.assertEqual(self.model.techLogin(self.form.user.text(), self.form.pswd.text()), True)
        #loginButton = self.form.login
        #QTest.mouseClick(loginButton, Qt.LeftButton)
        #self.assertEqual(QAdminHome.isActiveWindow, True)

if __name__ == "__main__":
    unittest.main()