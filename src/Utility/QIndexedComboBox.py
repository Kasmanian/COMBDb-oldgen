from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QComboBox

class QIndexedComboBox(QComboBox):
    def __init__(self, row, column, form, kind):
        super(QIndexedComboBox, self).__init__()
        self.row = row
        self.column = column
        self.form = form
        self.kind = kind
        self.currentIndexChanged.connect(self.handleCurrentIndexChanged)

    #@throwsViewableException
    def handleCurrentIndexChanged(self):
        self.form.updateTable(self.kind, self.row, self.column)