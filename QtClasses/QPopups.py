from PyQt5.QtWidgets import QMainWindow, QFileDialog
from PyQt5.uic import loadUi
from pathlib import Path
import json

class QArchiver(QMainWindow):
    def __init__(self):
        super(QArchiver, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Archive_Prompt.ui', self)
        self.no.clicked.connect(self.close)

class QConfirmation(QMainWindow):
    def __init__(self):
        super(QConfirmation, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Confirmation_Window.ui', self)
        self.Cancel.clicked.connect(self.close)

class QError(QMainWindow):
    def __init__(self, message):
        super(QError, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Error_Window.ui', self)
        self.ok.clicked.connect(self.close)
        self.errorMessage.setText(str(message))

class QFileBrowser(QMainWindow):
    def __init__(self, app, key, ext):
        super(QFileBrowser, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Set_File_Path_Form.ui', self)
        self.browse.clicked.connect(self.handleBrowsePressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.back.clicked.connect(self.close)
        self.app = app
        self.key = key
        self.ext = ext

    def handleBrowsePressed(self):
        self.filePath.setText(QFileDialog.getOpenFileName(self, 'Open File', 'C:', self.ext)[0])

    def handleSavePressed(self):
        try:
            with open('COMBDb\local.json', 'r+') as f:
                data = json.load(f)
                data[self.key] = str(Path(self.filePath.text()))
                f.seek(0)
                json.dump(data, f)
                f.truncate()
                self.app.handleFilePathSaved()
        except Exception as e:
            self.app.showErrorScreen(e)