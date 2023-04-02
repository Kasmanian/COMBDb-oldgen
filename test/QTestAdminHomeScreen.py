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
from Utility.QPrefixGraph import QPrefixGraph

app = QApplication(sys.argv)

class AdminHomeTest(unittest.TestCase):
    '''Test the Admin Login GUI'''
    def setUp(self):
        '''Create the GUI'''
        self.model = QModel()
        self.model.connect()
        self.form = QAdminHome(self.model, self)
        self.techLogin = self.model.techLogin

    def test_defaults(self):
        '''Test the GUI in its default state'''
        #Default state of Admin Home UI
        self.assertTrue("Select your option:" in self.form.header.text())
        self.assertEqual("Culture Order Forms", self.form.cultureOrder.text())
        self.assertEqual("Result Entry", self.form.resultEntry.text())
        self.assertEqual("QA Report", self.form.qaReport.text())
        self.assertEqual(" Settings", self.form.settings.text())
        self.assertEqual(" Logout", self.form.logout.text())

if __name__ == "__main__":
    unittest.main()