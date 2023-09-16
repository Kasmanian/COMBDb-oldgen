from PyQt5.uic import loadUi
from pathlib import Path
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow
from mailmerge import MailMerge
from PyQt5.QtCore import QDate, QTimer
from PyQt5.QtGui import QIcon

from Utility.QAdminLogin import QAdminLogin
from Utility.QViewableException import QViewableException

class QDUWLOrder(QMainWindow):
    def __init__(self, model, view):
        super(QDUWLOrder, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_DUWL_Order_Form.ui", self)
        self.find.setIcon(QIcon("Icon/searchIcon.png"))
        self.find2.setIcon(QIcon("Icon/filterIcon.png"))
        self.addClinician.setIcon(QIcon("Icon/addClinicianIcon.png"))
        self.save.setIcon(QIcon("Icon/saveIcon.png"))
        self.clear.setIcon(QIcon("Icon/clearIcon.png"))
        self.home.setIcon(QIcon("Icon/menuIcon.png"))
        self.print.setIcon(QIcon("Icon/printIcon.png"))
        self.back.setIcon(QIcon("Icon/backIcon.png"))
        self.clearAll.setIcon(QIcon("Icon/clearAllIcon.png"))
        self.remove.setIcon(QIcon("Icon/removeIcon.png"))
        self.currentKit = 1
        self.kitList = []
        self.printList = {}
        self.kitNum.setText("1")
        self.numOrders.setValue(1)
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)
        self.shipDate.setDate(
            QDate(self.model.date.year, self.model.date.month, self.model.date.day)
        )
        self.find.clicked.connect(self.handleSearchPressed)
        self.find2.clicked.connect(self.handleAdvancedSearchPressed)
        self.addClinician.clicked.connect(self.handleAddClinicianPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.clearAll.clicked.connect(self.handleClearAllPressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.remove.clicked.connect(self.handleRemovePressed)
        self.kitTWid.setColumnCount(1)
        self.kitTWid.itemClicked.connect(self.activateRemove)
        self.print.setEnabled(False)
        self.remove.setEnabled(False)
        self.row.setRange(1, 10)
        self.col.setRange(1, 3)
        self.kitTWid.setHorizontalHeaderLabels(["Sample ID"])
        header = self.kitTWid.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.rejectedCheckBox.clicked.connect(self.handleRejectedPressed)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.msg = ""

    @QViewableException.throwsViewableException
    def handleAdvancedSearchPressed(self):
        self.view.showAdvancedSearchScreen(self, "duwlOrder", False)

    @QViewableException.throwsViewableException
    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

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

    @QViewableException.throwsViewableException
    def activateRemove(self):
        self.remove.setEnabled(True)

    @QViewableException.throwsViewableException
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
                self.errorMessage.setText("Sample ID may only contain numbers")
                return
            if len(self.saID.text()) != 6:
                self.handleClearPressed()
                self.saID.setText("xxxxxx")
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText("Sample ID must contain 6 digits")
                return
            self.sample = self.model.findSample(
                "Waterlines",
                int(self.saID.text()),
                "[Clinician], [Comments], [Notes], [Shipped], [Rejection Date], [Rejection Reason]",
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
            self.saID.setText(str(self.sample[6]))

        saID = int(self.saID.text())
        saIDCheck = str(saID)[0:2] + "-" + str(saID)[2:]
        kitListValues = [value for elem in self.kitList for value in elem.values()]
        if saIDCheck not in kitListValues:
            if self.sample is not None:
                self.saID.setEnabled(False)
                if self.sample[5] != None:
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
                self.cText.setText(self.sample[1])
                self.nText.setText(self.sample[2])
                self.shipDate.setDate(self.view.dtToQDate(self.sample[3]))
                self.rejectedMessage.setText(self.sample[5])
                self.msg = self.sample[5]
                self.rejectedCheckBox.setEnabled(True)
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: green"
                )
                self.errorMessage.setText("Found previous order: " + str(saID))
        else:
            self.saID.setText("xxxxxx")
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("This DUWL Order has already been added")
            return

    @QViewableException.throwsViewableException
    def handleAddClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    @QViewableException.throwsViewableException
    def handleBackPressed(self):
        self.view.showDUWLNav()

    @QViewableException.throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    @QViewableException.throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if self.clinDrop.currentText():
            if (
                self.rejectedCheckBox.isChecked() and self.rejectedMessage.text() != ""
            ) or not self.rejectedCheckBox.isChecked():
                if self.saID.text() == "":
                    self.saID.setText("0")
                self.sample = self.model.findSample(
                    "Waterlines",
                    int(self.saID.text()),
                    "[Clinician], [Comments], [Notes], [Shipped], [Rejection Date], [Rejection Reason], [Tech]",
                )
                if self.sample is None:
                    numOrders = (
                        1
                        if int(self.numOrders.text()) == None
                        else int(self.numOrders.text())
                    )
                    for x in range(numOrders):
                        saID = self.view.model.addWaterlineOrder(
                            self.view.entries[self.clinDrop.currentText()]["db"],
                            self.shipDate.date(),
                            self.cText.toPlainText(),
                            self.nText.toPlainText(),
                            self.view.currentTech,
                        )
                        if saID:
                            self.saID.setText(str(saID))
                            self.kitList.append(
                                {
                                    "sampleID": f"{str(saID)[0:2]}-{str(saID)[2:]}",
                                    "clinician": self.clinDrop.currentText().split(",")[
                                        0
                                    ],
                                    "opID": "Operatory ID: ______________________",
                                    "agent": "Cleaning Agent:  ____________________",
                                    "collected": "Collection Date: _________",
                                }
                            )
                            self.printList[str(saID)] = self.currentKit - 1
                            self.currentKit = len(self.kitList) + 1
                            self.kitNum.setText(str(self.currentKit))
                    self.handleClearPressed()
                    self.errorMessage.setStyleSheet(
                        "font: 12pt 'MS Shell Dlg 2'; color: green"
                    )
                    self.errorMessage.setText("Created New DUWL Order: " + str(saID))
                    self.view.auditor(
                        self.view.currentTech, "Create", self.saID.text(), "DUWL_Order"
                    )
                else:
                    if not self.saID.isEnabled():
                        sampleID = self.saID.text()
                        if self.rejectedCheckBox.isChecked() and self.sample[4] is None:
                            rejDate = QDate.currentDate()
                        elif (
                            self.rejectedCheckBox.isChecked()
                            and self.sample[4] is not None
                        ):
                            rejDate = self.view.dtToQDate(self.sample[4])
                        else:
                            rejDate = None
                        saID = self.model.updateWaterlineOrder(
                            int(self.saID.text()),
                            self.view.entries[self.clinDrop.currentText()]["db"],
                            self.shipDate.date(),
                            self.cText.toPlainText(),
                            self.nText.toPlainText(),
                            rejDate,
                            self.rejectedMessage.text()
                            if self.rejectedCheckBox.isChecked()
                            else None,
                            self.view.currentTech,
                        )
                        if saID:
                            self.saID.setText(self.saID.text())
                            self.kitList.append(
                                {
                                    "sampleID": f"{str(self.saID.text())[0:2]}-{str(self.saID.text())[2:]}",
                                    "clinician": self.clinDrop.currentText().split(",")[
                                        0
                                    ],
                                    "opID": "Operatory ID: ______________________",
                                    "agent": "Cleaning Agent:  ____________________",
                                    "collected": "Collection Date: _________",
                                }
                            )
                            self.printList[self.saID.text()] = self.currentKit - 1
                            self.currentKit = len(self.kitList) + 1
                            self.kitNum.setText(str(self.currentKit))
                        self.handleClearPressed()
                        self.errorMessage.setStyleSheet(
                            "font: 12pt 'MS Shell Dlg 2'; color: green"
                        )
                        self.errorMessage.setText(
                            "Existing DUWL Order Updated: " + sampleID
                        )
                        self.view.auditor(self.view.currentTech, "Update", sampleID, "DUWL_Order")
                    else:
                        self.handleClearPressed()
                        self.errorMessage.setStyleSheet(
                            "font: 12pt 'MS Shell Dlg 2'; color: red"
                        )
                        self.errorMessage.setText("Please search order to edit it")
            else:
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText("Please enter reason for rejection")
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Please select a clinician")

    @QViewableException.throwsViewableException
    def handleClearPressed(self):
        self.kitNum.setText(str(self.currentKit))
        self.saID.clear()
        self.cText.clear()
        self.nText.clear()
        self.numOrders.setValue(1)
        self.save.setEnabled(True)
        self.clear.setEnabled(True)
        self.clinDrop.setCurrentIndex(0)
        self.errorMessage.setText("")
        self.tabWidget.setCurrentIndex(0)
        self.shipDate.setDate(
            QDate(self.model.date.year, self.model.date.month, self.model.date.day)
        )
        self.rejectedCheckBox.setCheckState(False)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.rejectionError.clear()
        self.msg = ""
        self.saID.setEnabled(True)
        self.handleRejectedPressed()
        self.updateTable()

    @QViewableException.throwsViewableException
    def handleClearAllPressed(self):
        self.kitList.clear()
        self.currentKit = 1
        self.kitNum.setText("1")
        self.printList.clear()
        self.updateTable()
        self.save.setEnabled(True)

    @QViewableException.throwsViewableException
    def handleRemovePressed(self):
        del self.kitList[self.printList[self.kitTWid.currentItem().text()]]
        del self.printList[self.kitTWid.currentItem().text()]
        count = 0
        for key in self.printList.keys():
            self.printList[key] = count
            count += 1
        self.updateTable()
        self.currentKit = len(self.kitList) + 1
        self.kitNum.setText(str(self.currentKit))
        self.remove.setEnabled(False)

    @QViewableException.throwsViewableException
    def updateTable(self):
        self.kitTWid.setRowCount(len(self.printList.keys()))
        count = 0
        for item in self.printList.keys():
            self.kitTWid.setItem(count, 0, QAdminLogin.QTableWidgetItem(item))
            count += 1
        if len(self.printList.keys()) > 0:
            self.print.setEnabled(True)
        else:
            self.print.setEnabled(False)

    @QViewableException.throwsViewableException
    def handlePrintPressed(self):
        template = str(Path().resolve()) + r"\templates\duwl_labels.docx"
        dst = self.view.tempify(template)
        document = MailMerge(template)
        numLabelBlanks = int(self.col.value() - 1) * 10 + int(self.row.value()) - 1
        numTotalLabels = len(self.kitList) + numLabelBlanks
        numFullPagesToPrint = numTotalLabels // 30
        numRowsToPrint = (numFullPagesToPrint) * 10 + min(
            numTotalLabels - 30 * numFullPagesToPrint, 10
        )
        labelList = [None] * numRowsToPrint
        mergeFields = ["sampleID", "clinician", "opID", "agent", "collected"]
        for i in range(0, numRowsToPrint):
            labelList[i] = {}
            for j in range(0, 3):
                kitListIndex = (i // 10 * 30) + (i % 10) + (j * 10) - numLabelBlanks
                indexAtBlankLabel = kitListIndex < 0 or kitListIndex >= len(
                    self.kitList
                )
                for mergeField in mergeFields:
                    labelList[i][mergeField + str(j + 1)] = (
                        None
                        if indexAtBlankLabel
                        else self.kitList[kitListIndex][mergeField]
                    )
        document.merge_rows("sampleID1", labelList)
        document.write(dst)
        self.view.convertAndPrint(dst)

    @QViewableException.throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")