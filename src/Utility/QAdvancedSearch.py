from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QIcon

from Utility.QViewableException import QViewableException

class QAdvancedSearch(QMainWindow):
    def __init__(self, model, view, orderForm, selector):
        super(QAdvancedSearch, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        self.orderForm = orderForm
        self.selector = selector
        loadUi("UI Screens/COMBdb_Advanced_Search_Form2.ui", self)
        self.find.setIcon(QIcon("Icon/searchIcon.png"))
        self.back.setIcon(QIcon("Icon/backIcon.png"))
        self.clear.setIcon(QIcon("Icon/clearIcon.png"))
        self.add.setIcon(QIcon("Icon/addIcon.png"))
        self.find.clicked.connect(self.handleSearchPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.add.clicked.connect(self.handleAddPressed)
        self.add.setEnabled(False)
        self.searchTable.itemSelectionChanged.connect(
            lambda: self.handleOrderSelected()
        )
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)
        if (
            self.selector == "duwlOrder"
            or self.selector == "duwlReceive"
            or self.selector == "duwlResult"
        ):
            self.fName.setEnabled(False)
            self.fName.setText("Not searchable")
            self.lName.setEnabled(False)
            self.lName.setText("Not searchable")

    @QViewableException.throwsViewableException
    def handleSearchPressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if self.saID.text() != "":
            if not self.saID.text().isdigit():
                self.handleClearPressed()
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText("Sample ID must only contain numbers")
                return
            if len(self.saID.text()) != 6:
                self.handleClearPressed()
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText("Sample ID must contain 6 digits")
                return

        if self.selector == "cultureOrder":
            # Set search parameters
            sampleID = int(self.saID.text()) if self.saID.text() != "" else 0
            clin = (
                self.view.entries[self.clinDrop.currentText()]["db"]
                if self.clinDrop.currentText() != ""
                else 0
            )
            if (
                sampleID == 0
                and clin == 0
                and self.fName.text() == ""
                and self.lName.text() == ""
            ):
                return
            inputs = {
                "SampleID": sampleID if sampleID != 0 else None,
                "First": self.fName.text() if self.fName.text() != "" else None,
                "Last": self.lName.text() if self.lName.text() != "" else None,
                "Clinician": clin if clin != 0 else None,
            }
            # Query data, join and sort
            cultures = self.model.findSamples(
                "Cultures",
                inputs,
                "[SampleID], [ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes], [Rejection Date], [Rejection Reason]",
            )
            cats = self.model.findSamples(
                "CATs",
                inputs,
                "[SampleID], [ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes], [Rejection Date], [Rejection Reason]",
            )
            self.results = cultures + cats
            self.results = sorted(self.results, key=lambda x: x[0])
            # Initialize table
            self.searchTable.setRowCount(len(self.results))
            for i in range(0, len(self.results)):
                self.searchTable.setItem(
                    i, 0, QTableWidgetItem(str(self.results[i][0]))
                )
                self.searchTable.setItem(
                    i,
                    1,
                    QTableWidgetItem(
                        str(self.results[i][3]) + " " + str(self.results[i][4])
                    ),
                )
                clinician = self.model.findClinician(self.results[i][2])
                self.searchTable.setItem(
                    i,
                    2,
                    QTableWidgetItem(
                        self.view.fClinicianNameNormal(
                            clinician[0], clinician[1], clinician[2], clinician[3]
                        )
                    ),
                )
                self.searchTable.setItem(
                    i, 3, QTableWidgetItem(str(self.results[i][5]))
                )

        elif self.selector == "duwlOrder":
            # Set search parameters
            sampleID = int(self.saID.text()) if self.saID.text() != "" else 0
            clin = (
                self.view.entries[self.clinDrop.currentText()]["db"]
                if self.clinDrop.currentText() != ""
                else 0
            )
            if sampleID == 0 and clin == 0:
                return
            inputs = {
                "SampleID": sampleID if sampleID != 0 else None,
                "Clinician": clin if clin != 0 else None,
            }
            # Query data, join and sort
            self.results = self.model.findSamples(
                "Waterlines",
                inputs,
                "[Clinician], [Comments], [Notes], [Shipped], [Rejection Date], [Rejection Reason], [SampleID]",
            )
            # Initialize table
            self.searchTable.setRowCount(len(self.results))
            for i in range(0, len(self.results)):
                self.searchTable.setItem(
                    i, 0, QTableWidgetItem(str(self.results[i][6]))
                )
                self.searchTable.setItem(i, 1, QTableWidgetItem("N/A"))
                clinician = self.model.findClinician(self.results[i][0])
                self.searchTable.setItem(
                    i,
                    2,
                    QTableWidgetItem(
                        self.view.fClinicianNameNormal(
                            clinician[0], clinician[1], clinician[2], clinician[3]
                        )
                    ),
                )
                self.searchTable.setItem(i, 3, QTableWidgetItem("Waterline"))

        elif self.selector == "duwlReceive":
            # Set search parameters
            sampleID = int(self.saID.text()) if self.saID.text() != "" else 0
            clin = (
                self.view.entries[self.clinDrop.currentText()]["db"]
                if self.clinDrop.currentText() != ""
                else 0
            )
            if sampleID == 0 and clin == 0:
                return
            inputs = {
                "SampleID": sampleID if sampleID != 0 else None,
                "Clinician": clin if clin != 0 else None,
            }
            # Query data, join and sort
            self.results = self.model.findSamples(
                "Waterlines",
                inputs,
                "[Clinician], [Comments], [Notes], [OperatoryID], [Product], [Procedure], [Collected], [Received], [Rejection Date], [Rejection Reason], [SampleID]",
            )
            # Initialize table
            self.searchTable.setRowCount(len(self.results))
            for i in range(0, len(self.results)):
                self.searchTable.setItem(
                    i, 0, QTableWidgetItem(str(self.results[i][10]))
                )
                self.searchTable.setItem(i, 1, QTableWidgetItem("N/A"))
                clinician = self.model.findClinician(self.results[i][0])
                self.searchTable.setItem(
                    i,
                    2,
                    QTableWidgetItem(
                        self.view.fClinicianNameNormal(
                            clinician[0], clinician[1], clinician[2], clinician[3]
                        )
                    ),
                )
                self.searchTable.setItem(i, 3, QTableWidgetItem("Waterline"))

        elif self.selector == "cultureResult":
            # Set search parameters
            sampleID = int(self.saID.text()) if self.saID.text() != "" else 0
            clin = (
                self.view.entries[self.clinDrop.currentText()]["db"]
                if self.clinDrop.currentText() != ""
                else 0
            )
            if (
                sampleID == 0
                and clin == 0
                and self.fName.text() == ""
                and self.lName.text() == ""
            ):  # this will fail for classes that dont use first and last name
                return
            inputs = {
                "SampleID": sampleID if sampleID != 0 else None,
                "First": self.fName.text() if self.fName.text() != "" else None,
                "Last": self.lName.text() if self.lName.text() != "" else None,
                "Clinician": clin if clin != 0 else None,
            }
            # Query data
            self.results = self.model.findSamples(
                "Cultures",
                inputs,
                "[ChartID], [Clinician], [First], [Last], [Tech], [Collected], [Received], [Reported], [Type], [Direct Smear], [Aerobic Results], [Anaerobic Results], [Comments], [Notes], [Rejection Date], [Rejection Reason], [SampleID]",
            )
            # Initialize table
            self.searchTable.setRowCount(len(self.results))
            for i in range(0, len(self.results)):
                self.searchTable.setItem(
                    i, 0, QTableWidgetItem(str(self.results[i][16]))
                )
                self.searchTable.setItem(
                    i,
                    1,
                    QTableWidgetItem(
                        str(self.results[i][2]) + " " + str(self.results[i][3])
                    ),
                )
                clinician = self.model.findClinician(self.results[i][1])
                self.searchTable.setItem(
                    i,
                    2,
                    QTableWidgetItem(
                        self.view.fClinicianNameNormal(
                            clinician[0], clinician[1], clinician[2], clinician[3]
                        )
                    ),
                )
                self.searchTable.setItem(
                    i, 3, QTableWidgetItem(str(self.results[i][8]))
                )

        elif self.selector == "catResult":
            # Set search parameters
            sampleID = int(self.saID.text()) if self.saID.text() != "" else 0
            clin = (
                self.view.entries[self.clinDrop.currentText()]["db"]
                if self.clinDrop.currentText() != ""
                else 0
            )
            if (
                sampleID == 0
                and clin == 0
                and self.fName.text() == ""
                and self.lName.text() == ""
            ):  # this will fail for classes that dont use first and last name
                return
            inputs = {
                "SampleID": sampleID if sampleID != 0 else None,
                "First": self.fName.text() if self.fName.text() != "" else None,
                "Last": self.lName.text() if self.lName.text() != "" else None,
                "Clinician": clin if clin != 0 else None,
            }
            # Query data
            self.results = self.model.findSamples(
                "CATs",
                inputs,
                "[Clinician], [First], [Last], [Tech], [Reported], [Type], [Volume (ml)], [Time (min)], [Initial (pH)], [Flow Rate (ml/min)], [Buffering Capacity (pH)], [Strep Mutans (CFU/ml)], [Lactobacillus (CFU/ml)], [Comments], [Notes], [Collected], [Received], [Rejection Date], [Rejection Reason], [SampleID]",
            )
            # Initialize table
            self.searchTable.setRowCount(len(self.results))
            for i in range(0, len(self.results)):
                self.searchTable.setItem(
                    i, 0, QTableWidgetItem(str(self.results[i][19]))
                )
                self.searchTable.setItem(
                    i,
                    1,
                    QTableWidgetItem(
                        str(self.results[i][1]) + " " + str(self.results[i][2])
                    ),
                )
                clinician = self.model.findClinician(self.results[i][0])
                self.searchTable.setItem(
                    i,
                    2,
                    QTableWidgetItem(
                        self.view.fClinicianNameNormal(
                            clinician[0], clinician[1], clinician[2], clinician[3]
                        )
                    ),
                )
                self.searchTable.setItem(
                    i, 3, QTableWidgetItem(str(self.results[i][5]))
                )

        elif self.selector == "duwlResult":
            # Set search parameters
            sampleID = int(self.saID.text()) if self.saID.text() != "" else 0
            clin = (
                self.view.entries[self.clinDrop.currentText()]["db"]
                if self.clinDrop.currentText() != ""
                else 0
            )
            if sampleID == 0 and clin == 0:
                return
            inputs = {
                "SampleID": sampleID if sampleID != 0 else None,
                "Clinician": clin if clin != 0 else None,
            }
            # Query data, join and sort
            self.results = self.model.findSamples(
                "Waterlines",
                inputs,
                "[Clinician], [Bacterial Count], [CDC/ADA], [Reported], [Comments], [Notes], [Rejection Date], [Rejection Reason], [OperatoryID], [SampleID]",
            )
            # Initialize table
            self.searchTable.setRowCount(len(self.results))
            for i in range(0, len(self.results)):
                self.searchTable.setItem(
                    i, 0, QTableWidgetItem(str(self.results[i][9]))
                )
                self.searchTable.setItem(i, 1, QTableWidgetItem("N/A"))
                clinician = self.model.findClinician(self.results[i][0])
                self.searchTable.setItem(
                    i,
                    2,
                    QTableWidgetItem(
                        self.view.fClinicianNameNormal(
                            clinician[0], clinician[1], clinician[2], clinician[3]
                        )
                    ),
                )
                self.searchTable.setItem(i, 3, QTableWidgetItem("Waterline"))

        else:
            print("Could not add order")

    @QViewableException.throwsViewableException
    def handleOrderSelected(self):
        self.add.setEnabled(True)

    @QViewableException.throwsViewableException
    def handleAddPressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        self.sample = self.results[self.searchTable.currentRow()]
        if self.selector != None:
            self.orderForm.handleSearchPressed(self.sample)
        else:
            print("Could not add order")
        self.close()

    @QViewableException.throwsViewableException
    def handleBackPressed(self):
        self.close()

    @QViewableException.throwsViewableException
    def handleClearPressed(self):
        if (
            self.selector == "duwlOrder"
            or self.selector == "duwlReceive"
            or self.selector == "duwlResult"
        ):
            self.saID.clear()
            self.clinDrop.setCurrentIndex(0)
            self.searchTable.setRowCount(0)
        else:
            self.saID.clear()
            self.fName.clear()
            self.lName.clear()
            self.clinDrop.setCurrentIndex(0)
            self.searchTable.setRowCount(0)

    @QViewableException.throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")