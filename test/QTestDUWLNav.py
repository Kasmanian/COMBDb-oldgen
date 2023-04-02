import sys
import unittest
from PyQt5.QtWidgets import QApplication

import sys

sys.path.insert(0, r'C:\Users\Hoboburger\Desktop\COMBDb\src')
 
from Orders.QDUWLNav import QDUWLNav

app = QApplication(sys.argv)

class DUWLNavTest(unittest.TestCase):
    '''Test the DUWL Nav GUI'''
    def setUp(self):
        '''Create the GUI'''
        self.form = QDUWLNav(self)

    def test_defaults(self):
        '''Test the GUI in its default state'''
        #Default state of DUWL Nav ui
        self.assertTrue("DUWL Cultures" in self.form.header.text())
        #Order Culture button
        self.assertEqual("Order Culture", self.form.orderCulture.text())
        #Receiving Culture button
        self.assertEqual("Receiving Culture", self.form.receivingCulture.text())
        #Back button
        self.assertEqual(" Back", self.form.back.text())

if __name__ == "__main__":
    unittest.main()