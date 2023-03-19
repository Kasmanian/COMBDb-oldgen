from PyQt5.uic import loadUi
from pathlib import Path
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from mailmerge import MailMerge
from docxtpl import DocxTemplate
from PyQt5.QtGui import QIcon

class QRejectionLog(QMainWindow):
    def __init__(self, model, view):
        super(QRejectionLog, self).__init__()
        self.view = view
        self.model = model
        #self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_Settings_Rejection_Log_Form.ui", self)
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.back.clicked.connect(self.handleBackPressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        rejectedCulture = self.model.findRejections('Cultures', '[SampleID], [Type], [Clinician], [Rejection Date], [Rejection Reason]')
        rejectedCAT = self.model.findRejections('CATs', '[SampleID], [Type], [Clinician], [Rejection Date], [Rejection Reason]')
        rejectedDUWL = self.model.findRejections('Waterlines', '[SampleID], [Clinician], [Rejection Date], [Rejection Reason]')
        count = 0
        for entry in rejectedDUWL:
            entry = list(entry)
            entry.insert(1, "Waterline")
            entry = tuple(entry)
            rejectedDUWL[count] = entry
            count += 1
        rejections = rejectedCulture + rejectedCAT + rejectedDUWL
        count = 0
        for tup in rejections:
            tup = list(tup)
            new = []
            new.append(tup[0])
            new.append(tup[1])
            clinician = self.model.findClinician(tup[2])
            new.append(self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3]))
            new.append(self.view.fSlashDate(tup[3]))
            new.append(tup[4])
            rejections[count] = new
            count += 1
        rejections = sorted(rejections, key=lambda x: x[0])
        header = self.rejLogTable.horizontalHeader()
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.Stretch)
        self.rejLogTable.setRowCount(len(rejections))
        for i in range(0, len(rejections)):
            self.rejLogTable.setItem(i, 0, QTableWidgetItem(str(rejections[i][0])))
            self.rejLogTable.setItem(i, 1, QTableWidgetItem(rejections[i][1]))
            self.rejLogTable.setItem(i, 2, QTableWidgetItem(rejections[i][2]))
            self.rejLogTable.setItem(i, 3, QTableWidgetItem(rejections[i][3]))
            self.rejLogTable.setItem(i, 4, QTableWidgetItem(rejections[i][4]))
        self.rejLogTable.sortItems(0,0)
        self.rejLogTable.resizeColumnsToContents()

    #@throwsViewableException
    def handlePrintPressed(self):
        self.printList = []
        for i in range(self.rejLogTable.rowCount()):
            self.printList.append([self.rejLogTable.item(i, 0).text(), self.rejLogTable.item(i, 1).text(), self.rejLogTable.item(i, 2).text(), self.rejLogTable.item(i, 3).text(), self.rejLogTable.item(i, 4).text()])
        template = str(Path().resolve())+r'\templates\rejection_log_template.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        document.merge(
            tech=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
        )
        document.write(dst)
        context = {
            'headers1' : ['Sample ID', 'Type', 'Clincian', 'Rejection Date', 'Rejection Reason'],
            'servers1' : self.printList
        }
        document = DocxTemplate(dst)
        document.render(context)
        document.save(dst)
        self.view.convertAndPrint(dst)
        
    #@throwsViewableException
    def handleBackPressed(self):
        self.view.showSettingsNav()

    #@throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()