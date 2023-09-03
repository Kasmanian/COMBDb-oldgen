from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QIcon
from PyQt5 import QtWidgets

from Utility.QViewableException import QViewableException
from Utility.QPrefixGraph import QPrefixGraph

class QManagePrefixes(QMainWindow):
    def __init__(self, model, view):
        super(QManagePrefixes, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        self.ref = QPrefixGraph(self.model)
        loadUi("UI Screens/COMBdb_Settings_Manage_Prefixes_Form.ui", self)
        self.add.setIcon(QIcon("Icon/addIcon.png"))
        self.save.setIcon(QIcon("Icon/saveIcon.png"))
        self.clear.setIcon(QIcon("Icon/clearIcon.png"))
        self.home.setIcon(QIcon("Icon/menuIcon.png"))
        self.back.setIcon(QIcon("Icon/backIcon.png"))
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.add.clicked.connect(self.handleAddPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.aeTWid.itemSelectionChanged.connect(
            lambda: self.handlePrefixSelected("Aerobic")
        )
        self.anTWid.itemSelectionChanged.connect(
            lambda: self.handlePrefixSelected("Anaerobic")
        )
        self.abTWid.itemSelectionChanged.connect(
            lambda: self.handlePrefixSelected("Antibiotics")
        )
        self.prefixesTabWidget.currentChanged.connect(self.clearSelection)
        aeHeader = self.aeTWid.horizontalHeader()
        aeHeader.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        anHeader = self.anTWid.horizontalHeader()
        anHeader.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        abHeader = self.abTWid.horizontalHeader()
        abHeader.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.currentPrefix = ""
        self.selectedPrefix = {}
        self.updateTable("Aerobic")
        self.updateTable("Anaerobic")
        self.updateTable("Antibiotics")
        self.save.setEnabled(False)

    @QViewableException.throwsViewableException
    def clearSelection(self):
        if self.prefixesTabWidget.currentIndex() == 0:
            self.anTWid.clearSelection()
            self.abTWid.clearSelection()
        elif self.prefixesTabWidget.currentIndex() == 1:
            self.aeTWid.clearSelection()
            self.abTWid.clearSelection()
        else:
            self.aeTWid.clearSelection()
            self.anTWid.clearSelection()

    @QViewableException.throwsViewableException
    def updateTable(self, type):
        widget = (
            self.aeTWid
            if type == "Aerobic"
            else self.anTWid
            if type == "Anaerobic"
            else self.abTWid
        )
        prefix = self.model.selectPrefixes(type, "Prefix, Word")
        widget.setRowCount(0)
        widget.setRowCount(len(prefix))
        widget.setColumnCount(2)
        widget.setColumnWidth(0, 50)
        widget.setColumnWidth(1, 300)
        for i in range(0, len(prefix)):
            widget.setItem(i, 0, QTableWidgetItem(prefix[i][0]))
            widget.setItem(i, 1, QTableWidgetItem(prefix[i][1]))
        widget.sortItems(0, 0)

    @QViewableException.throwsViewableException
    def handlePrefixSelected(self, type):
        widget = (
            self.aeTWid
            if type == "Aerobic"
            else self.anTWid
            if type == "Anaerobic"
            else self.abTWid
        )
        prefix = widget.item(widget.currentRow(), 0)
        word = widget.item(widget.currentRow(), 1)
        if prefix and word:
            self.selectedPrefix = {prefix.text(): [type, word.text()]}
            self.pName.setText(list(self.selectedPrefix.keys())[0])
            keyList = self.selectedPrefix.get(list(self.selectedPrefix.keys())[0])
            self.type.setCurrentIndex(self.type.findText(keyList[0]))
            if self.type.currentText() == "Antibiotics":
                self.pName.setEnabled(False)
            else:
                self.pName.setEnabled(True)
            self.word.setText(keyList[1])
            self.currentPrefix = self.model.findPrefix(
                self.pName.text(), "Entry, Type, Prefix, Word"
            )
            self.type.setEnabled(False)
            self.add.setEnabled(False)
            self.save.setEnabled(True)

    @QViewableException.throwsViewableException
    def handleBackPressed(self):
        self.view.showSettingsNav()

    @QViewableException.throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    @QViewableException.throwsViewableException
    def handleAddPressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        prefix = self.pName.text()
        word = self.word.text()
        type = self.type.currentText()
        if self.type.currentText() and self.pName.text() and self.word.text():
            if not self.ref.exists("prefix", prefix) and not self.ref.exists(
                "word", word
            ):
                self.model.addPrefixes(
                    self.type.currentText(), self.pName.text(), self.word.text()
                )
                self.ref.populate(type)
                self.updateTable(self.type.currentText())
                self.handleClearPressed()
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: green"
                )
                self.errorMessage.setText(
                    "Successfully added prefix: "
                    + prefix
                    + ":"
                    + word
                    + " to table: "
                    + type
                )
                self.view.auditor(
                    self.view.currentTech, "Edit", self.user.text(), "Settings_Edit_Technician"
                )
            else:
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText(
                    "An entry with that prefix or word already exists"
                )
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Type, Prefix and Word are required")

    @QViewableException.throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if self.pName.text() and self.word.text() and self.type.currentText():
            self.selectedPrefix
            prefixChanged = self.pName.text() not in self.selectedPrefix
            wordChanged = (
                self.word.text()
                != self.selectedPrefix.get(list(self.selectedPrefix.keys())[0])[1]
            )
            prefixPassed = True
            wordPassed = True
            if prefixChanged:
                prefixPassed = not self.ref.exists("prefix", self.pName.text())
            if wordChanged:
                wordPassed = not self.ref.exists("word", self.word.text())
            if prefixPassed and wordPassed:
                self.model.updatePrefixes(
                    self.currentPrefix[0],
                    self.type.currentText(),
                    self.pName.text(),
                    self.word.text(),
                )
                self.ref.populate(self.type.currentText())
                self.updateTable(self.type.currentText())
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: green"
                )
                self.errorMessage.setText("Successfully Updated Prefix")
            else:
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText(
                    "An entry with that prefix or word already exists"
                )
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Type, Prefix and Word are required")

    @QViewableException.throwsViewableException
    def handleClearPressed(self):
        self.aeTWid.clearSelection()
        self.anTWid.clearSelection()
        self.abTWid.clearSelection()
        self.type.setCurrentIndex(0)
        self.pName.clear()
        self.word.clear()
        self.errorMessage.clear()
        self.add.setEnabled(True)
        self.save.setEnabled(False)
        self.pName.setEnabled(True)
        self.type.setEnabled(True)

    @QViewableException.throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")