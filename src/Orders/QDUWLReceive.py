from PyQt5.uic import loadUi
from pathlib import Path
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem
from mailmerge import MailMerge
from PyQt5.QtCore import QDate, QTimer
from PyQt5.QtGui import QIcon

from Utility.QAdminLogin import QAdminLogin


class QDUWLReceive(QMainWindow):
    def __init__(self, model, view):
        super(QDUWLReceive, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_DUWL_Receive_Form.ui", self)
        self.find.setIcon(QIcon('Icon/searchIcon.png'))
        self.find2.setIcon(QIcon('Icon/filterIcon.png'))
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.print.setIcon(QIcon('Icon/printIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.clearAll.setIcon(QIcon('Icon/clearAllIcon.png'))
        self.remove.setIcon(QIcon('Icon/removeIcon.png'))
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)
        self.currentKit = 1
        self.kitList = []
        self.printList = {}
        self.save.setEnabled(False)
        self.print.setEnabled(False)
        self.recDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.colDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.find.clicked.connect(self.handleSearchPressed)
        self.find2.clicked.connect(self.handleAdvancedSearchPressed)
        self.clearAll.clicked.connect(self.handleClearAllPressed)
        self.remove.clicked.connect(self.handleRemovePressed)
        self.kitTWid.setColumnCount(1)
        self.kitTWid.setHorizontalHeaderLabels(['Sample ID'])
        header = self.kitTWid.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.kitTWid.itemClicked.connect(self.activateRemove)
        self.print.setEnabled(False)
        self.remove.setEnabled(False)
        self.rejectedCheckBox.clicked.connect(self.handleRejectedPressed)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.msg = "" 

    #@throwsViewableException
    def handleAdvancedSearchPressed(self):
        self.view.showAdvancedSearchScreen(self, "duwlReceive")

    #@throwsViewableException
    def activateRemove(self):
        self.remove.setEnabled(True)

    #@throwsViewableException
    def handleBackPressed(self):
        self.view.showDUWLNav()

    #@throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    #@throwsViewableException
    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

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
        saID = int(self.saID.text())
        saIDCheck = str(saID)[0:2]+ "-" +str(saID)[2:]
        kitListValues = [value for elem in self.kitList for value in elem.values()]
        if saIDCheck not in kitListValues:
            if data == False:
                if not self.saID.text().isdigit():
                    self.saID.setText('xxxxxx')
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                    self.errorMessage.setText("Sample ID may only contain numbers")
                    return
                self.sample = self.model.findSample('Waterlines', int(self.saID.text()), '[Clinician], [Comments], [Notes], [OperatoryID], [Product], [Procedure], [Collected], [Received], [Rejection Date], [Rejection Reason]')
                if self.sample is None or len(self.saID.text()) != 6:
                    self.saID.setText('xxxxxx')
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                    self.errorMessage.setText("Sample ID not found")
            else:
                self.sampleID = data
                data = None
                self.saID.setText(str(self.sampleID[10]))
            if self.sample is not None:
                if self.sample[9] != None:
                    self.rejectionError.setText("(REJECTED)")
                    self.rejectedCheckBox.setChecked(True)
                    self.handleRejectedPressed()
                clinician = self.model.findClinician(self.sample[0])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list']+1)
                self.cText.setText(self.sample[1])
                self.nText.setText(self.sample[2])
                self.operatory.setText(self.sample[3])
                self.product.setText(self.sample[4])
                self.procedure.setText(self.sample[5])
                self.colDate.setDate(self.view.dtToQDate(self.sample[6]))
                self.recDate.setDate(self.view.dtToQDate(self.sample[7]))
                self.rejectedMessage.setText(self.sample[9])
                self.msg = self.sample[9]
                self.save.setEnabled(True)
                self.saID.setEnabled(False)
                self.rejectedCheckBox.setEnabled(True)
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                self.errorMessage.setText("Found DUWL Order: " + str(saID))
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("This DUWL Order has already been added")

    #@throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        saID = int(self.saID.text())
        if self.clinDrop.currentText():
            if (self.rejectedCheckBox.isChecked() and self.rejectedMessage.text() != "") or not self.rejectedCheckBox.isChecked():
                self.sample = self.model.findSample('Waterlines', int(self.saID.text()), '[Rejection Date]')
                if self.rejectedCheckBox.isChecked() and self.sample[0] is None:
                    rejDate = QDate.currentDate()
                elif self.rejectedCheckBox.isChecked() and self.sample[0] is not None:
                    rejDate = self.sample[0]
                else:
                    rejDate = None
                if self.model.addWaterlineReceiving(
                    saID,
                    self.operatory.text(),
                    self.view.entries[self.clinDrop.currentText()]['db'],
                    self.colDate.date(),
                    self.recDate.date(),
                    self.product.text(),
                    self.procedure.text(),
                    self.cText.toPlainText(),
                    self.nText.toPlainText(),
                    rejDate,
                    self.rejectedMessage.text() if self.rejectedCheckBox.isChecked() else None,
                    self.model.getCurrUser()
                ):
                    clinician = self.clinDrop.currentText().split(', ')
                    self.kitList.append({
                        'underline1': '__________',
                        'clinicianName': clinician[1] + " " + clinician[0],
                        'sampleID': f'{str(saID)[0:2]}-{str(saID)[2:]}',
                        'underline2': '__________',
                        'underline3': '__________'
                    })
                    self.printList[str(saID)] = self.currentKit-1
                    self.currentKit = len(self.kitList)+1
                    self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear()
                    self.handleClearPressed()
                    self.save.setEnabled(False)
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                    self.errorMessage.setText("Saved DUWL Order: " + str(saID))
                    self.view.auditor(self.model.getCurrUser(), "Update", str(saID), 'DUWL_Receive')
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Please enter reason for rejection")
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Please select a clinician")

    #@throwsViewableException
    def handleClearPressed(self):
        self.saID.clear()
        self.saID.setEnabled(True)
        self.clinDrop.setCurrentIndex(0)
        self.cText.clear()
        self.nText.clear()
        self.operatory.clear()
        self.procedure.clear()
        self.product.clear()
        self.save.setEnabled(False)
        self.clear.setEnabled(True)
        self.tabWidget.setCurrentIndex(0)
        self.errorMessage.setText("")
        self.rejectedCheckBox.setCheckState(False)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.rejectionError.clear()
        self.msg = ""
        self.handleRejectedPressed()
        self.updateTable()

    #@throwsViewableException
    def handleClearAllPressed(self):
        self.kitList.clear()
        self.currentKit = 1
        self.printList.clear()
        self.updateTable()

    #@throwsViewableException
    def handleRemovePressed(self):
        del self.kitList[self.printList[self.kitTWid.currentItem().text()]]
        del self.printList[self.kitTWid.currentItem().text()]
        count = 0
        for key in self.printList.keys():
            self.printList[key] = count
            count += 1
        self.updateTable()
        self.currentKit = len(self.kitList)+1
        self.remove.setEnabled(False)

    #@throwsViewableException
    def updateTable(self):
        self.kitTWid.setRowCount(len(self.printList.keys()))
        count = 0
        for item in self.printList.keys():
            self.kitTWid.setItem(count, 0, QTableWidgetItem(item))
            count += 1
        if len(self.printList.keys())>0:
            self.print.setEnabled(True)
        else:
            self.print.setEnabled(False)

    #@throwsViewableException
    def handlePrintPressed(self):
        template = str(Path().resolve())+r'\templates\pending_duwl_cultures_template.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        document.merge_rows('sampleID', self.kitList)
        document.merge(received=self.view.fSlashDate(self.recDate.date()))
        document.write(dst)
        self.view.convertAndPrint(dst)

    #@throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")
