from PyQt5.uic import loadUi
from pathlib import Path
from PyQt5.QtWidgets import QMainWindow, QFileDialog
import json
from PyQt5.QtGui import QIcon

from Utility.QViewableException import QViewableException

class QFilePath(QMainWindow):
    def __init__(self, model, view):
        super(QFilePath, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Set_File_Path_Form.ui", self)
        self.save.setIcon(QIcon("Icon/saveIcon.png"))
        self.back.setIcon(QIcon("Icon/backIcon.png"))
        self.back.clicked.connect(self.handleBackPressed)
        self.browse.clicked.connect(self.handleBrowsePressed)
        self.save.clicked.connect(self.handleSavePressed)
        with open("local.json", "r+") as JSON:
            self.currDBText = json.load(JSON)
        self.currDB.setText(
            "Current filepath: " + self.currDBText["DBQ"]
        ) if self.currDBText["DBQ"] != "" else self.currDB.setText(
            "Current filepath: None"
        )

    @QViewableException.throwsViewableException
    def handleBackPressed(self):
        self.view.showSettingsNav()
        self.close()

    @QViewableException.throwsViewableException
    def handleBrowsePressed(self):
        fname = QFileDialog.getOpenFileName(
            self, "Open File", "C:", "MS Access Files (*.accdb)"
        )
        self.filePath.setText(fname[0])

    @QViewableException.throwsViewableException
    def handleSavePressed(self):
        with open("local.json", "r+") as JSON:
            data = json.load(JSON)
            data["DBQ"] = str(Path(self.filePath.text()))
            JSON.seek(0)  # rewind
            json.dump(data, JSON)
            JSON.truncate()
        if not self.model.connect():
            self.view.showErrorScreen(
                "Could not open database with the specified path."
            )
        else:
            self.view.setClinicianList()
            self.close()