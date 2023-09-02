from PyQt5.uic import loadUi
from pathlib import Path
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from mailmerge import MailMerge
from PyQt5.QtCore import QDate, QTimer
from PyQt5.QtGui import QIcon

from Utility.QAdminLogin import QAdminLogin
from Utility.QViewableException import QViewableException

class QDUWLResult(QMainWindow):
    def __init__(self, model, view):
        super(QDUWLResult, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_DUWL_Result_Form.ui", self)
        self.find.setIcon(QIcon("Icon/searchIcon.png"))
        self.find2.setIcon(QIcon("Icon/filterIcon.png"))
        self.save.setIcon(QIcon("Icon/saveIcon.png"))
        self.clear.setIcon(QIcon("Icon/clearIcon.png"))
        self.home.setIcon(QIcon("Icon/menuIcon.png"))
        self.print.setIcon(QIcon("Icon/printIcon.png"))
        self.back.setIcon(QIcon("Icon/backIcon.png"))
        self.clearAll.setIcon(QIcon("Icon/clearAllIcon.png"))
        self.remove.setIcon(QIcon("Icon/removeIcon.png"))
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)
        self.currentKit = 1
        self.kitList = []
        self.meets = {"Meets": 1, "Fails to Meet": 2}
        self.printList = {}
        self.save.setEnabled(False)
        self.print.setEnabled(False)
        self.repDate.setDate(
            QDate(self.model.date.year, self.model.date.month, self.model.date.day)
        )
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
        self.kitTWid.setHorizontalHeaderLabels(["Sample ID"])
        header = self.kitTWid.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.bacterialCount.editingFinished.connect(self.lineEdited)
        self.kitTWid.itemClicked.connect(self.activateRemove)
        self.print.setEnabled(False)
        self.remove.setEnabled(False)
        self.cdcADA.setEnabled(False)
        self.rejectedCheckBox.clicked.connect(self.handleRejectedPressed)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.msg = "" 

    #@throwsViewableException
    def handleAdvancedSearchPressed(self):
        self.view.showAdvancedSearchScreen(self, "duwlResult")

    #@throwsViewableException
    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    #@throwsViewableException
    def handleRejectedPressed(self):
        if self.rejectedCheckBox.isChecked():
            self.rejectedMessage.setStyleSheet(
                "background-color: rgb(255, 255, 255); border-style: solid; border-width: 1px"
            )
            self.rejectedMessage.setPlaceholderText("Reason?")
            self.rejectedMessage.setEnabled(True)
            self.rejectedMessage.setText(self.msg)
        else:
            self.rejectedMessage.setStyleSheet(
                "background-color: rgb(123, 175, 212); border-style: solid; border-width: 0px"
            )
            self.rejectedMessage.setPlaceholderText("")
            self.rejectedMessage.setEnabled(False)
            self.rejectedMessage.clear()

    #@throwsViewableException
    def lineEdited(self):
        if self.bacterialCount.text().isdigit():
            if int(self.bacterialCount.text()) < 500:
                self.cdcADA.setCurrentIndex(1)
            else:
                self.cdcADA.setCurrentIndex(2)
        else:
            self.bacterialCount.setText("")
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText(
                "Bacterial Count may only contain positive integers"
            )

    #@throwsViewableException
    def activateRemove(self):
        self.remove.setEnabled(True)

    #@throwsViewableException
    def handleBackPressed(self):
        self.view.showResultEntryNav()

    #@throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    #@throwsViewableException
    def handleSearchPressed(self, data):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if data == False:
            if not self.saID.text().isdigit():
                self.handleClearPressed()
                self.saID.setText("xxxxxx")
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText("Sample ID must only contain numbers")
                return
            if len(self.saID.text()) != 6:
                self.handleClearPressed()
                self.saID.setText("xxxxxx")
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText("Sample ID must contains 6 digits")
                return
            self.sample = self.model.findSample(
                "Waterlines",
                int(self.saID.text()),
                "[Clinician], [Bacterial Count], [CDC/ADA], [Reported], [Comments], [Notes], [Rejection Date], [Rejection Reason], [OperatoryID]",
            )
            if self.sample is None:
                self.handleClearPressed()
                self.saID.setText("xxxxxx")
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText("Sample ID not found")
                return
        else:
            self.sample = data
            data = False
            self.saID.setText(str(self.sample[9]))
        saID = int(self.saID.text())
        saIDCheck = str(saID)[0:2] + "-" + str(saID)[2:]
        kitListValues = [value for elem in self.kitList for value in elem.values()]
        if saIDCheck not in kitListValues: 
            if self.sample is not None:
                self.saID.setEnabled(False)
                if self.sample[7] != None:  
                    self.rejectionError.setText("(REJECTED)")
                    self.rejectedCheckBox.setChecked(True)
                    self.handleRejectedPressed()
                clinician = self.model.findClinician(self.sample[0])
                clinicianName = self.view.fClinicianName(
                    clinician[0], clinician[1], clinician[2], clinician[3]
                )
                self.clinDrop.setCurrentIndex(
                    self.view.entries[clinicianName]["list"] + 1
                )
                self.bacterialCount.setText(
                    str(self.sample[1]) if self.sample[1] else None
                )
                self.cdcADA.setCurrentIndex(
                    self.meets[self.sample[2]] if self.sample[2] else 0
                )
                self.repDate.setDate(self.view.dtToQDate(self.sample[3]))
                self.cText.setText(self.sample[4])
                self.nText.setText(self.sample[5])
                self.rejectedMessage.setText(self.sample[7])
                self.msg = self.sample[7]
                self.save.setEnabled(True)
                self.rejectedCheckBox.setEnabled(True)
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: green"
                )
                self.errorMessage.setText("Found DUWL Order: " + self.saID.text())
        else:
            self.saID.setText("xxxxxx")
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("This DUWL Order has already been added")
            return                

    #@throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        saID = int(self.saID.text())
        if self.clinDrop.currentText() != "":
            if (
                self.rejectedCheckBox.isChecked() and self.rejectedMessage.text() != ""
            ) or not self.rejectedCheckBox.isChecked():
                if self.model.addWaterlineResult(
                    saID,
                    self.view.entries[self.clinDrop.currentText()]["db"],
                    self.repDate.date(),
                    int(self.bacterialCount.text()),
                    self.cdcADA.currentText(),
                    self.cText.toPlainText(),
                    self.nText.toPlainText(),
                    QDate.currentDate() if self.rejectedCheckBox.isChecked() else None,
                    self.rejectedMessage.text()
                    if self.rejectedCheckBox.isChecked()
                    else None,
                    self.view.currentTech,
                ):
                    self.kitList.append(
                        {
                            "sampleID": f"{str(saID)[0:2]}-{str(saID)[2:]}",
                            "operatory": self.sample[8],
                            "count": self.bacterialCount.text(),
                            "cdcADA": self.cdcADA.currentText(),
                        }
                    )
                    self.printList[str(saID)] = self.currentKit - 1
                    self.currentKit = len(self.kitList) + 1
                    self.handleClearPressed()
                    self.errorMessage.setStyleSheet(
                        "font: 12pt 'MS Shell Dlg 2'; color: green"
                    )
                    self.errorMessage.setText("Saved DUWL Result Form: " + str(saID))
                    self.save.setEnabled(False)
                    self.view.auditor(self.view.currentTech, "Update", saID, "DUWL_Result")
            else:
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText("Please enter reason for rejection")
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Please select a clinician")

    #@throwsViewableException
    def handleClearPressed(self):
        self.saID.clear()
        self.saID.setEnabled(True)
        self.cText.clear()
        self.nText.clear()
        self.bacterialCount.clear()
        self.cdcADA.setCurrentText(None)
        self.clear.setEnabled(True)
        self.clinDrop.setCurrentIndex(0)
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
        self.currentKit = len(self.kitList) + 1
        self.remove.setEnabled(False)

    #@throwsViewableException
    def updateTable(self):
        self.kitTWid.setRowCount(len(self.printList.keys()))
        count = 0
        for item in self.printList.keys():
            self.kitTWid.setItem(count, 0, QTableWidgetItem(item))
            count += 1
        if len(self.printList.keys()) > 0:
            self.print.setEnabled(True)
        else:
            self.print.setEnabled(False)

    #@throwsViewableException
    def handlePrintPressed(self):
        template = str(Path().resolve()) + r"\templates\duwl_results_template.docx"
        dst = self.view.tempify(template)
        document = MailMerge(template)
        document.merge_rows("sampleID", self.kitList)
        clinician = self.model.findClinicianFull(self.sample[0])
        clinicianName = self.view.fClinicianNameNormal(
            clinician[0], clinician[1], clinician[2], clinician[5]
        )
        document.merge(
            reported=self.view.fSlashDate(self.repDate.date()),
        )
        addressData = {
            "clinicianAddress": "",
            "clinicianName": clinicianName + "\n" if clinicianName is not None else "",
            "designation": clinician[5] + "\n"
            if (clinician[5] != "" and clinician[5] is not None)
            else "",
            "address1": clinician[6] + "\n"
            if (clinician[6] != "" and clinician[6] is not None)
            else "",
            "address2": clinician[7] + "\n"
            if (clinician[7] != "" and clinician[7] is not None)
            else "",
            "cityStateZip": clinician[8]
            + ", "
            + clinician[9]
            + " "
            + str(clinician[10])
            if (
                clinician[8] != ""
                and clinician[8] is not None
                and clinician[9] != ""
                and clinician[9] is not None
                and str(clinician[10]) != ""
                and str(clinician[10]) is not None
            )
            else "",
        }
        addressDataCopy = addressData.copy()
        for x in addressData:
            if addressData.get(x) != "" and x != "clinicianAddress":
                addressDataCopy["clinicianAddress"] = addressDataCopy.get(
                    "clinicianAddress"
                ) + addressDataCopy.get(x)
        addressDataList = [addressDataCopy]
        document.merge_rows("clinicianAddress", addressDataList)

        document.write(dst)
        self.view.convertAndPrint(dst)
    
    #@throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")