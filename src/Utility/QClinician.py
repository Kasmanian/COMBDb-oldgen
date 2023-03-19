from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtCore import QDate, QTimer
from PyQt5.QtGui import QIcon

from Utility.QAdminLogin import QAdminLogin

class QClinician(QMainWindow):
    def __init__(self, model, view, dropdown):
        super(QClinician, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        self.dropdown = dropdown
        loadUi("UI Screens/COMBdb_Add_New_Clinician.ui", self)
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)
        self.clear.clicked.connect(self.handleClearPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.clinDrop.currentIndexChanged.connect(self.selectedClinician)
        self.enrollDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))

    def selectedClinician(self):
        if self.clinDrop.currentText() == "":
            return
        else:
            clinician = self.model.findClinicianFull(self.view.entries[self.clinDrop.currentText()]['db'])
            self.title.setCurrentIndex(self.title.findText(clinician[0]))
            self.fName.setText(clinician[1])
            self.lName.setText(clinician[2])
            self.address1.setText(clinician[6])
            self.address2.setText(clinician[7])
            self.city.setText(clinician[8])
            self.state.setCurrentIndex(self.state.findText(clinician[9]))
            self.zip.setText(clinician[10])
            self.phone.setText(clinician[3])
            self.fax.setText(clinician[4])
            self.email.setText(clinician[11])
            self.enrollDate.setDate(self.view.dtToQDate(clinician[12]))
            self.designation.setText(clinician[5])
            self.cText.setText(clinician[13])

    def handleSavePressed(self): #Incorporate validation to make sure clinician is actually added to DB
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        title, first, last = "", "", ""
        if self.fName.text() and self.lName.text() and self.address1.text() and self.city.text() and self.state.currentText() and self.zip.text():
            #self.sample = self.model.findClinicianFull(self.view.entries[self.clinDrop.currentText()]['db'])
            if self.clinDrop.currentText() == "": #and self.sample is None:
                self.model.addClinician(
                    self.title.currentText(),
                    self.fName.text(),
                    self.lName.text(),
                    self.designation.text(),
                    self.phone.text(),
                    self.fax.text(),
                    self.email.text(),
                    self.address1.text(),
                    self.address2.text(),
                    self.city.text(),
                    self.state.currentText(),
                    self.zip.text(),
                    None,
                    None,
                    self.cText.toPlainText()
                )
                title = self.title.currentText()
                first = self.fName.text()
                last = self.lName.text()
                self.view.setClinicianList()
                self.clinDrop.clear()
                self.clinDrop.addItem("")
                self.clinDrop.addItems(self.view.names)
                self.handleClearPressed()
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                self.errorMessage.setText("New clinician added: " + title + " " + first + " " + last)
                self.view.auditor(QAdminLogin.currentTech, "Create", title + '' + first + ' ' + last, 'Clinician')
            else:
                self.model.updateClinician(
                    self.view.entries[self.clinDrop.currentText()]['db'],
                    self.title.currentText(),
                    self.fName.text(),
                    self.lName.text(),
                    self.designation.text(),
                    self.phone.text(),
                    self.fax.text(),
                    self.email.text(),
                    self.address1.text(),
                    self.address2.text(),
                    self.city.text(),
                    self.state.currentText(),
                    self.zip.text(),
                    None,
                    self.cText.toPlainText()
                )
                title = self.title.currentText()
                first = self.fName.text()
                last = self.lName.text()
                self.view.setClinicianList()
                self.clinDrop.clear()
                self.clinDrop.addItem("")
                self.clinDrop.addItems(self.view.names)
                self.handleClearPressed()
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                self.errorMessage.setText("Updated Existing Clinician: " + title + " " + first + " " + last)
                self.view.auditor(QAdminLogin.currentTech, "Update", title + '' + first + ' ' + last, 'Clinician')
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("* Denotes Required Fields")

    #@throwsViewableException
    def handleBackPressed(self):
        self.close()

    #@throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        self.close()

    #@throwsViewableException
    def handleClearPressed(self):
        self.title.setCurrentIndex(0)
        self.fName.clear()
        self.lName.clear()
        self.address1.clear()
        self.address2.clear()
        self.city.clear()
        self.state.setCurrentIndex(0)
        self.zip.clear()
        self.phone.clear()
        self.fax.clear()
        self.email.clear()
        self.enrollDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.designation.clear()
        self.cText.clear()
        self.errorMessage.clear()
        self.clinDrop.setCurrentIndex(0)

    #@throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")