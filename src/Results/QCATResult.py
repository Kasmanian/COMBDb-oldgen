from PyQt5.uic import loadUi
from pathlib import Path
from PyQt5.QtWidgets import QMainWindow
from mailmerge import MailMerge
from PyQt5.QtCore import QDate, QTimer, QThread
from PyQt5.QtGui import QIcon
import re

from Utility.QViewableException import QViewableException

class QCATResult(QMainWindow):
    def __init__(self, model, view):
        super(QCATResult, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_CAT_Result_Form.ui", self)
        self.find.setIcon(QIcon("Icon/searchIcon.png"))
        self.find2.setIcon(QIcon("Icon/filterIcon.png"))
        self.save.setIcon(QIcon("Icon/saveIcon.png"))
        self.print.setIcon(QIcon("Icon/printIcon.png"))
        self.clear.setIcon(QIcon("Icon/clearIcon.png"))
        self.home.setIcon(QIcon("Icon/menuIcon.png"))
        self.back.setIcon(QIcon("Icon/backIcon.png"))
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)
        self.volume.setText("0.00")
        self.collectionTime.setText("0.00")
        self.flowRate.setText("0.00")
        self.volume.editingFinished.connect(lambda: self.lineEdited(True))
        self.collectionTime.editingFinished.connect(lambda: self.lineEdited(False))
        self.save.setEnabled(False)
        self.print.setEnabled(False)
        self.repDate.setDate(
            QDate(self.model.date.year, self.model.date.month, self.model.date.day)
        )
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.print.clicked.connect(self.threader)
        self.find.clicked.connect(self.handleSearchPressed)
        self.find2.clicked.connect(self.handleAdvancedSearchPressed)
        self.rejectedCheckBox.clicked.connect(self.handleRejectedPressed)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.msg = "" 

    @QViewableException.throwsViewableException
    def threader(self):
        self.thread = QThread()
        if self.handleSavePressed():
            self.thread.started.connect(self.handlePrintPressed)
            self.thread.start()
            self.thread.exit()

    @QViewableException.throwsViewableException
    def handleAdvancedSearchPressed(self):
        self.view.showAdvancedSearchScreen(self, "catResult")

    @QViewableException.throwsViewableException
    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    @QViewableException.throwsViewableException
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
    def lineEdited(self, arg):
        lineEdit = self.volume if arg else self.collectionTime
        pattern = re.compile("^[0-9\.]*$")
        if lineEdit.text() != "" and pattern.match(lineEdit.text()):
            if float(self.collectionTime.text()) != 0:
                vol = float(self.volume.text())
                colTime = float(self.collectionTime.text())
                value = str(vol if arg else colTime)
                rate = round(vol / colTime, 2)
                lineEdit.setText(value)
                self.flowRate.setText(str(rate))
                self.errorMessage.setText(None)
            else:
                self.flowRate.setText("0.00")
        else:
            lineEdit.setText("0.00")

    @QViewableException.throwsViewableException
    def handleBackPressed(self):
        self.view.showResultEntryNav()

    @QViewableException.throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    @QViewableException.throwsViewableException
    def handleSearchPressed(self, data):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if data == False:
            if not self.saID.text().isdigit():
                self.saID.setText("xxxxxx")
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText("Sample ID must only contain numbers")
                return
            self.sample = self.model.findSample(
                "CATs",
                int(self.saID.text()),
                "[Clinician], [First], [Last], [Tech], [Reported], [Type], [Volume (ml)], [Time (min)], [Initial (pH)], [Flow Rate (ml/min)], [Buffering Capacity (pH)], [Strep Mutans (CFU/ml)], [Lactobacillus (CFU/ml)], [Comments], [Notes], [Collected], [Received], [Rejection Date], [Rejection Reason]",
            )
            if self.sample is None or len(self.saID.text()) != 6:
                self.saID.setText("xxxxxx")
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText("Sample ID not found")
        else:
            self.sample = data
            self.saID.setText(str(self.sample[19]))
            data = False
        if self.sample is not None:
            if self.sample[18] != None:
                self.rejectionError.setText("(REJECTED)")
                self.rejectedCheckBox.setChecked(True)
                self.handleRejectedPressed()
            clinician = self.model.findClinician(self.sample[0])
            clinicianName = self.view.fClinicianName(
                clinician[0], clinician[1], clinician[2], clinician[3]
            )
            self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]["list"] + 1)
            self.fName.setText(self.sample[1])
            self.lName.setText(self.sample[2])
            self.repDate.setDate(self.view.dtToQDate(self.sample[4]))
            self.volume.setText(
                str(self.sample[6]) if self.sample[12] is not None else None
            )
            self.collectionTime.setText(
                str(self.sample[7]) if self.sample[12] is not None else None
            )
            self.initialPH.setText(
                str(self.sample[8]) if self.sample[12] is not None else None
            )
            self.flowRate.setText(
                str(self.sample[9]) if self.sample[12] is not None else None
            )
            self.bufferingCapacityPH.setText(
                str(self.sample[10]) if self.sample[12] is not None else None
            )
            self.strepMutansCount.setText(
                str(self.sample[11]) if self.sample[12] is not None else None
            )
            self.lactobacillusCount.setText(
                str(self.sample[12]) if self.sample[12] is not None else None
            )
            self.cText.setText(self.sample[13])
            self.nText.setText(self.sample[14])
            self.rejectedMessage.setText(self.sample[18])
            self.msg = self.sample[18]
            self.saID.setEnabled(False)
            self.save.setEnabled(True)
            self.print.setEnabled(True)
            self.clear.setEnabled(True)
            self.rejectedCheckBox.setEnabled(True)
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
            self.errorMessage.setText("Found CAT Order: " + self.saID.text())

    @QViewableException.throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if (
            self.fName.text()
            and self.lName.text()
            and self.clinDrop.currentText() != ""
        ):
            if float(self.collectionTime.text()) != 0:
                if (
                    self.rejectedCheckBox.isChecked()
                    and self.rejectedMessage.text() != ""
                ) or not self.rejectedCheckBox.isChecked():
                    saID = int(self.saID.text())
                    if self.model.addCATResult(
                        saID,
                        self.view.entries[self.clinDrop.currentText()]["db"],
                        self.fName.text(),
                        self.lName.text(),
                        self.view.currentTech,
                        self.repDate.date(),
                        "Caries",
                        float(self.volume.text()),
                        float(self.collectionTime.text()),
                        float(self.flowRate.text()),
                        float(self.initialPH.text()),
                        float(self.bufferingCapacityPH.text()),
                        int(self.strepMutansCount.text()),
                        int(self.lactobacillusCount.text()),
                        self.cText.toPlainText(),
                        self.nText.toPlainText(),
                        QDate.currentDate()
                        if self.rejectedCheckBox.isChecked()
                        else None,
                        self.rejectedMessage.text()
                        if self.rejectedCheckBox.isChecked()
                        else None,
                    ):
                        self.save.setEnabled(True)
                        self.print.setEnabled(True)
                        self.errorMessage.setStyleSheet(
                            "font: 12pt 'MS Shell Dlg 2'; color: green"
                        )
                        self.errorMessage.setText("Saved CAT Result Form: " + str(saID))
                        self.view.auditor(
                            self.view.currentTech, "Update", self.saID.text(), "CAT_Result"
                        )
                        return True
                else:
                    self.errorMessage.setStyleSheet(
                        "font: 12pt 'MS Shell Dlg 2'; color: red"
                    )
                    self.errorMessage.setText("Please enter reason for rejection")
                    return False
            else:
                self.flowRate.setText("x.xx")
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText("Cannot divide by 0")
                return False
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("* Denotes Required Fields")
            return False
        

    @QViewableException.throwsViewableException
    def handleClearPressed(self):
        self.saID.clear()
        self.saID.setEnabled(True)
        self.clinDrop.setCurrentIndex(0)
        self.fName.clear()
        self.lName.clear()
        self.volume.setText("0.00")
        self.initialPH.clear()
        self.collectionTime.setText("0.00")
        self.bufferingCapacityPH.clear()
        self.flowRate.setText("0.00")
        self.strepMutansCount.clear()
        self.lactobacillusCount.clear()
        self.repDate.setDate(self.view.dtToQDate(None))
        self.cText.clear()
        self.nText.clear()
        self.save.setEnabled(False)
        self.clear.setEnabled(True)
        self.print.setEnabled(False)
        self.errorMessage.setText("")
        self.tabWidget.setCurrentIndex(0)
        self.rejectedCheckBox.setCheckState(False)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.rejectionError.clear()
        self.msg = ""
        self.handleRejectedPressed()

    @QViewableException.throwsViewableException
    def handlePrintPressed(self):
        template = str(Path().resolve()) + r"\templates\cat_results_template.docx"
        dst = self.view.tempify(template)
        document = MailMerge(template)
        clinician = self.clinDrop.currentText().split(", ")
        document.merge(
            saID=f"{self.saID.text()[0:2]}-{self.saID.text()[2:6]}",
            patientName=f"{self.fName.text()} {self.lName.text()}",
            clinicianName=clinician[1] + " " + clinician[0],
            collected=self.view.fSlashDate(self.sample[15]),
            received=self.view.fSlashDate(self.sample[16]),
            flowRate=str(self.flowRate.text()),
            bufferingCapacity=str(self.bufferingCapacityPH.text()),
            smCount="{:.2e}".format(int(self.strepMutansCount.text())),
            lbCount="{:.2e}".format(int(self.lactobacillusCount.text())),
            reported=self.view.fSlashDate(self.repDate.date()),
            techName=self.model.tech,
            comments=self.cText.toPlainText(),
        )
        document.write(dst)
        self.view.convertAndPrint(dst)

    @QViewableException.throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")