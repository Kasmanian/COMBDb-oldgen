import sys
import unittest
from PyQt5.QtWidgets import QApplication
from __TestConstants import FILEPATH

import sys

sys.path.insert(0, FILEPATH)
 
from Orders.QDUWLNav import QDUWLNav
from QModel import QModel

app = QApplication(sys.argv)

class DUWLNavTest(unittest.TestCase):
    '''Test the DUWL Nav GUI'''
    def setUp(self):
        '''Create the GUI'''
        self.model = QModel()
        self.model.connect()
        self.form = QDUWLNav(self.model, self)
        self.techLogin = self.model.techLogin

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