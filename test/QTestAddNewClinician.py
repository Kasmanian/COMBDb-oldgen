import sys
import unittest
from PyQt5.QtWidgets import QApplication
from __TestConstants import FILEPATH
from PyQt5.QtTest import QTest 
from PyQt5.QtCore import Qt
from QView import QView

sys.path.insert(0, FILEPATH)

from QModel import QModel
from Utility.QClinician import QClinician

app = QApplication(sys.argv)

class AddNewClinicianTest(unittest.TestCase):
    '''Test the Add New Clinican HUI'''
    def setUp(self):
        '''Creat the GUI'''
        self.model = QModel
        self.model.connect
        self.form = QClinician(self.model, self)
        self.techLogin = self.model.techLogin

    def setFormToZero(self):
        '''Clear all fields'''
        QTest.mouseClick(self.form.clear, Qt.LeftButton)
        
    def test_defaults(self):
        '''Test the GUI in its default state'''
        #Default state of Add New Clinician UI
        names=QView.name
        self.forms.names=names
        self.assertTrue("Edit Clinician:", self.form.clinDropLabel.text())
       # self.assertEqual("", self.form.clinDrop.currentText())
        self.assertTrue("Title:", self.form.titleLabel.text())
        self.assertEqual("", self.form.title.text())
        self.assertTrue("*First Name:", self.form.fNameLabel.text())
        self.assertEqual("", self.form.fName.text())
        self.assertTrue("*Last Name:", self.form.lNameLabel.text())
        self.assertEqual("", self.form.lName.text())
        self.assertTrue("*Address 1:", self.form.address1Label.text())
        self.assertEqual("", self.form.address1.text())
        self.assertTrue("Address 2:", self.form.address2Label.text())
        self.assertEqual("", self.form.address2.text())
        self.assertTrue("*City:", self.form.cityLabel.text())
        self.assertEqual("", self.form.city.text())
        self.assertTrue("*State:", self.form.stateLabel.text())
        self.assertEqual("", self.form.state.text())
        self.assertTrue("*Zip Code:", self.form.zipLabel.text())
        self.assertEqual("", self.form.zip.text())
        self.assertTrue("Phone:", self.form.phoneLabel.text())
        self.assertEqual("", self.form.phone.text())
        self.assertTrue("Fax Number:", self.form.faxLabel.text())
        self.assertEqual("", self.form.fax.text())
        self.assertTrue("E-mail:", self.form.emailLabel.text())
        self.assertEqual("", self.form.email.text())
        self.assertTrue("Enrollment Date:", self.form.enrollDateLabel.text())
        self.assertEqual("", self.form.enrollDate.text())
        self.assertTrue("Practice Name:", self.form.designationLabel.text())
        self.assertEqual("", self.form.designation.text())
        self.assertTrue("Comment:", self.form.cTextLabelLabel.text())
        self.assertEqual("", self.form.cText.text())