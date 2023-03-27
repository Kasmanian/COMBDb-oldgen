from PyQt5.uic import loadUi
from pathlib import Path
from PyQt5.QtWidgets import QMainWindow
from mailmerge import MailMerge
from PyQt5.QtCore import QDate, QTimer, QThread
from PyQt5.QtGui import QIcon

from Utility.QAdminLogin import QAdminLogin

class QCultureOrder(QMainWindow):
    def __init__(self, model, view):
        super(QCultureOrder, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_Culture_Order_Form.ui", self)
        self.find.setIcon(QIcon('Icon/searchIcon.png'))
        self.find2.setIcon(QIcon('Icon/filterIcon.png'))
        self.addClinician.setIcon(QIcon('Icon/addClinicianIcon.png'))
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.print.setIcon(QIcon('Icon/printIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)
        self.addClinician.clicked.connect(self.handleAddNewClinicianPressed)
        self.find.clicked.connect(self.handleSearchPressed)
        self.find2.clicked.connect(self.handleAdvancedSearchPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        #self.print.clicked.connect(self.handlePrintPressed)
        self.print.clicked.connect(self.threader)
        self.clear.clicked.connect(self.handleClearPressed)
        #self.print.setEnabled(False)
        self.colDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.recDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))   
        #print(str(QDate.currentDate()) + " " + str(QTime.currentTime()))  
        self.rejectedCheckBox.clicked.connect(self.handleRejectedPressed)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.msg = "" 

    #@throwsViewableException
    def threader(self):
        self.thread = QThread()
        if self.handleSavePressed():
            self.thread.started.connect(self.handlePrintPressed)
            self.thread.start()
            self.thread.exit()

    #@throwsViewableException
    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    #@throwsViewableException
    def handleRejectedPressed(self):
        if self.rejectedCheckBox.isChecked():
            self.rejectedMessage.setStyleSheet("background-color: rgb(255, 255, 255); border-style: solid; border-width: 1px")
            self.rejectedMessage.setPlaceholderText("Reason?")
            self.rejectedMessage.setEnabled(True)
            self.rejectedMessage.setText(self.msg)
        else:
            self.rejectedMessage.setStyleSheet("background-color: rgb(123, 175, 212); border-style: solid; border-width: 0px")
            self.rejectedMessage.setPlaceholderText("")
            self.rejectedMessage.setEnabled(False)
            self.rejectedMessage.clear()

    #@throwsViewableException
    def handleSearchPressed(self, data):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        self.rejectedCheckBox.setEnabled(True)
        self.saID.setEnabled(False)
        self.type.setEnabled(False)
        if data == False:
            if not self.saID.text().isdigit():
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID may only contain numbers")
                return
            #self.sample = self.advancedSearch.queryData
            self.sample = self.model.findSample('Cultures', int(self.saID.text()), '[SampleID], [ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes], [Rejection Date], [Rejection Reason]')
            if self.sample is None:
                self.sample = self.model.findSample('CATs', int(self.saID.text()), '[SampleID], [ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes], [Rejection Date], [Rejection Reason]')
                if self.sample is None or len(self.saID.text()) != 6:
                    self.handleClearPressed()
                    self.saID.setText('xxxxxx')
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                    self.errorMessage.setText("Sample ID not found")
        else:
            self.sample = data
            self.saID.setText(str(self.sample[0]))
            data = None
        if self.sample is not None:
            if self.sample[11] != None:
                self.rejectionError.setText("(REJECTED)")
                self.rejectedCheckBox.setChecked(True)
                self.handleRejectedPressed()
            self.chID.setText(self.sample[1])
            clinician = self.model.findClinician(self.sample[2])
            clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
            self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list']+1)
            self.fName.setText(self.sample[3])
            self.lName.setText(self.sample[4])
            self.type.setCurrentIndex(self.type.findText(self.sample[5]))
            self.colDate.setDate(self.view.dtToQDate(self.sample[6]))
            self.recDate.setDate(self.view.dtToQDate(self.sample[7]))
            self.cText.setText(self.sample[8])
            self.nText.setText(self.sample[9])
            self.rejectedMessage.setText(self.sample[11])
            self.msg = self.sample[11]
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
            self.errorMessage.setText("Found previous order: " + self.saID.text())

    #@throwsViewableException
    def handleAdvancedSearchPressed(self):
        self.view.showAdvancedSearchScreen(self, "cultureOrder")

    #@throwsViewableException
    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    #@throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
    
    #@throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if self.fName.text() and self.lName.text() and self.type.currentText() and self.clinDrop.currentText() != "":
            if (self.rejectedCheckBox.isChecked() and self.rejectedMessage.text() != "") or not self.rejectedCheckBox.isChecked():
                if self.saID.text() == "":
                    self.saID.setText("0")
                self.sample = self.model.findSample('Cultures', int(self.saID.text()), '[ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes], [Rejection Date]')
                if self.sample is None:
                    self.sample = self.model.findSample('CATs', int(self.saID.text()), '[ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes], [Rejection Date]')
                    if self.sample is None:
                        table = 'CATs' if self.type.currentText()=='Caries' else 'Cultures'
                        #Create a new db entry - either culture or CAT
                        saID = self.view.model.addPatientOrder(
                            table,
                            self.chID.text(),
                            self.view.entries[self.clinDrop.currentText()]['db'],
                            self.fName.text(),
                            self.lName.text(),
                            self.colDate.date(),
                            self.recDate.date(),
                            self.type.currentText(),
                            self.model.getCurrUser(),
                            self.cText.toPlainText(),
                            self.nText.toPlainText(),
                        )
                        if saID:
                            self.saID.setText(str(saID))
                            #self.save.setEnabled(False)
                            self.print.setEnabled(True)
                            self.saID.setEnabled(False)
                            self.type.setEnabled(False)
                            self.rejectedCheckBox.setEnabled(True)
                            self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear()
                            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                            self.errorMessage.setText("Successfully saved order: " + str(self.saID.text()))
                            self.view.auditor(self.model.getCurrUser(), "Create", self.saID.text(), self.type.currentText() + '_Order')
                            return True
                    else: #Update existing CAT Order
                        if not self.saID.isEnabled() and not self.type.isEnabled():
                            if self.rejectedCheckBox.isChecked() and self.sample[9] is None:
                                rejDate = QDate.currentDate()
                            elif self.rejectedCheckBox.isChecked() and self.sample[9] is not None:
                                rejDate = self.view.dtToQDate(self.sample[9])
                            else:
                                rejDate = None
                            self.model.updateCultureOrder(
                                "CATs",
                                int(self.saID.text()),
                                self.chID.text(),
                                self.view.entries[self.clinDrop.currentText()]['db'],
                                self.fName.text(),
                                self.lName.text(),
                                self.colDate.date(),
                                self.recDate.date(),
                                self.type.currentText(),
                                self.model.getCurrUser(),
                                self.cText.toPlainText(),
                                self.nText.toPlainText(),
                                rejDate,
                                self.rejectedMessage.text() if self.rejectedCheckBox.isChecked() else None
                            )
                            #self.view.showConfirmationScreen("Are you sure you want to update an existing culture order?")
                            self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear()
                            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                            self.errorMessage.setText("Existing CAT Order Updated: " + str(self.saID.text())) 
                            self.view.auditor(self.model.getCurrUser(), "Update", self.saID.text(), self.type.currentText() + '_Order')
                            return True 
                        else: 
                            self.handleClearPressed()
                            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                            self.errorMessage.setText("Please search order to edit it")
                else: #Update existing Culture Order
                    if not self.saID.isEnabled() and not self.type.isEnabled():
                        if self.rejectedCheckBox.isChecked() and self.sample[9] is None:
                                rejDate = QDate.currentDate()
                        elif self.rejectedCheckBox.isChecked() and self.sample[9] is not None:
                            rejDate = self.view.dtToQDate(self.sample[9])
                        else:
                            rejDate = None
                        self.model.updateCultureOrder(
                            "Cultures",
                            int(self.saID.text()),
                            self.chID.text(),
                            self.view.entries[self.clinDrop.currentText()]['db'],
                            self.fName.text(),
                            self.lName.text(),
                            self.colDate.date(),
                            self.recDate.date(),
                            self.type.currentText(),
                            self.model.getCurrUser(),
                            self.cText.toPlainText(),
                            self.nText.toPlainText(),
                            rejDate,
                            self.rejectedMessage.text() if self.rejectedCheckBox.isChecked() else None
                        )
                        #self.view.showConfirmationScreen("Are you sure you want to update an existing culture order?")
                        self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear() 
                        self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                        self.errorMessage.setText("Existing Culture Order Updated: " + str(self.saID.text()))
                        self.view.auditor(self.model.getCurrUser(), "Update", self.saID.text(), self.type.currentText() + '_Order')
                        return True
                    else: 
                        self.handleClearPressed()
                        self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                        self.errorMessage.setText("Please search order to edit it")
                        return False
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Please enter reason for rejection")
                return False
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("* Denotes Required Fields")
            return False
        
    #@throwsViewableException
    def handlePrintPressed(self): 
        if self.type.currentText()!='Caries':
            template = str(Path().resolve())+r'\templates\culture_worksheet_template4.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            clinician=self.clinDrop.currentText().split(', ')
            document.merge(
                saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:]}',
                received=self.recDate.date().toString(),
                type=self.type.currentText(),
                chartID=self.chID.text(),
                clinicianName = clinician[1] + " " + clinician[0],
                patientName=f'{self.lName.text()}, {self.fName.text()}',
                comments=self.cText.toPlainText(),
                notes=self.nText.toPlainText(),
                techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
            )
            document.write(dst)
            self.view.convertAndPrint(dst)
        else:
            template = str(Path().resolve())+r'\templates\cat_worksheet_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            clinician=self.clinDrop.currentText().split(', ')
            document.merge(
                saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:]}',
                received=self.recDate.date().toString(),
                chartID=self.chID.text(),
                clinicianName = clinician[1] + " " + clinician[0],
                patientName=f'{self.lName.text()}, {self.fName.text()}',
                techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
            )
            document.write(dst)
            self.view.convertAndPrint(dst)
        return

    #@throwsViewableException
    def handleClearPressed(self):
        self.fName.clear()
        self.lName.clear()
        self.colDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.recDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.saID.clear()
        self.chID.clear()
        self.cText.clear()
        self.nText.clear()
        self.clinDrop.setCurrentIndex(0)
        self.type.setCurrentIndex(0)
        self.save.setEnabled(True)
        #self.print.setEnabled(False)
        self.clear.setEnabled(True)
        self.errorMessage.setText("")
        self.tabWidget.setCurrentIndex(0)
        self.saID.setEnabled(True)
        self.type.setEnabled(True)
        self.rejectedCheckBox.setCheckState(False)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.rejectionError.clear()
        self.msg = ""
        self.handleRejectedPressed()

    #@throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")
