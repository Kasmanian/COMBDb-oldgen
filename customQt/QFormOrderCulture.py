from PyQt5.QtWidgets import QMainWindow, QDate
from PyQt5.uic import loadUi
from pathlib import Path
from mailmerge import MailMerge

class QOrderCulture(QMainWindow):
    def __init__(self, app):
        super(QOrderCulture, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Culture_Order_Form.ui', self)
        self.clinicianDropDown.clear()
        self.clinicianDropDown.addItems(self.view.names)
        self.addClinician.clicked.connect(self.handleAddNewClinicianPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.print.setEnabled(False)
        self.collectionDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.receivedDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))

        self.app = app

    def handleAddNewClinicianPressed(self):
        self.app.showAddClinicianScreen(self.clinicianDropDown)

    def handleBackPressed(self):
        self.app.showCultureOrderNav()

    def handleReturnToMainMenuPressed(self):
        self.app.showAdminHomeScreen()
    
    def handleSavePressed(self):
        try:
            table = 'CATs' if self.cultureTypeDropDown.currentText()=='Caries' else 'Cultures'
            fields = ('[ChartID]', '[Clinician]', '[First]', '[Last]', '[Collected]', '[Received', '[Comments]', '[Notes]')
            sampleID = self.app.db.sample()
            if self.app.db.insert(
                table,
                fields,
                self.chartNum.text(),
                self.app.entries[self.clinicianDropDown.currentText()]['db'],
                self.firstName.text(),
                self.lastName.text(),
                self.collectionDate.date(),
                self.receivedDate.date(),
                self.comment.toPlainText()
            ):
                self.sampleID.setText(str(sampleID))
                self.save.setEnabled(False)
                self.print.setEnabled(True)
        except Exception as e:
            self.view.showErrorScreen(e)
    
    def handlePrintPressed(self):
        try:
            if self.cultureTypeDropDown.currentText()!='Caries':
                print(f'clinician: {self.clinicianDropDown.currentText()}')
                template = str(Path().resolve())+r'\COMBDb\templates\culture_worksheet_template.docx'
                dst = self.view.tempify(template)
                document = MailMerge(template)
                document.merge(
                    sampleID=f'{self.sampleID.text()[0:2]}-{self.sampleID.text()[2:]}',
                    received=self.receivedDate.date().toString(),
                    chartID=self.chartNum.text(),
                    clinicianName=self.clinicianDropDown.currentText(),
                    patientName=f'{self.lastName.text()}, {self.firstName.text()}',
                    comments=self.comment.toPlainText()
                )
                document.write(dst)
                try:
                    self.view.convertAndPrint(dst)
                except Exception as e:
                    self.view.showErrorScreen(e)
            else:
                template = str(Path().resolve())+r'\COMBDb\templates\cat_worksheet_template.docx'
                dst = self.view.tempify(template)
                document = MailMerge(template)
                document.merge(
                    sampleID=f'{self.sampleID.text()[0:2]}-{self.sampleID.text()[2:]}',
                    received=self.receivedDate.date().toString(),
                    chartID=self.chartNum.text(),
                    clinicianName=self.clinicianDropDown.currentText(),
                    patientName=f'{self.lastName.text()}, {self.firstName.text()}',
                )
                document.write(dst)
                try:
                    self.view.convertAndPrint(dst)
                except Exception as e:
                    self.view.showErrorScreen(e)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearPressed(self):
        try:
            self.firstName.clear()
            self.lastName.clear()
            self.collectionDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
            self.receivedDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
            self.sampleID.setText('xxxxxx')
            self.chartNum.clear()
            self.comment.clear()
            self.save.setEnabled(True)
            self.print.setEnabled(False)
            self.clear.setEnabled(True)
        except Exception as e:
            self.view.showErrorScreen(e)