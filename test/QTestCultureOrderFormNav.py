import sys
import unittest
from PyQt5.QtWidgets import QApplication
from __TestConstants import FILEPATH

sys.path.insert(0, FILEPATH)
 
from Orders.QOrderNav import QOrderNav
from QModel import QModel

app = QApplication(sys.argv)

class CultureOrderFormNavTest(unittest.TestCase):
    '''Test the Culture Order Nav GUI'''
    def setUp(self):
        '''Create the GUI'''
        self.model = QModel()
        self.model.connect()
        self.form = QOrderNav(self.model, self)
        self.techLogin = self.model.techLogin

    def test_defaults(self):
        '''Test the GUI in its default state'''
        #Default state of Admin Home UI
        self.assertTrue("Culture Order Form" in self.form.header.text())
        #Culture button and text underneath
        self.assertEqual("Culture", self.form.culture.text())
        self.assertTrue("Candida Cultures" in self.form.candidaCultures.text())
        self.assertTrue("Caries Activity Testing" in self.form.cariesActivityTesting.text())
        self.assertTrue("Endo Cultures" in self.form.endoCultures.text())
        self.assertTrue("Perio Cultures" in self.form.perioCultures.text())
        self.assertTrue("Research Cultures" in self.form.researchCultures.text())
        self.assertTrue("Surgical Cultures" in self.form.surgicalCultures.text())
        self.assertTrue("Tongue Cultures" in self.form.tongueCultures.text())
        #DUWL button and text underneath
        self.assertEqual("DUWL", self.form.duwl.text())
        self.assertTrue("DUWL Testing" in self.form.duwlTesting.text())
        #Back button
        self.assertEqual(" Back", self.form.back.text())

if __name__ == "__main__":
    unittest.main()