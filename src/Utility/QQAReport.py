from PyQt5.uic import loadUi
from pathlib import Path
from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem
import datetime
from mailmerge import MailMerge
from docxtpl import DocxTemplate
from PyQt5.QtCore import QDate, QTimer
from PyQt5.QtGui import QIcon

from Utility.QAdminLogin import QAdminLogin

class QQAReport(QMainWindow): #TODO
    def __init__(self, model, view):
        super(QQAReport, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_QA_Report_Screen.ui", self)
        self.find.setIcon(QIcon('Icon/searchIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        #self.print.clicked.connect(self.handlePrintPressed)
        self.find.clicked.connect(self.handleSearchPressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.fromDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.toDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))

    #@throwsViewableException
    def handleSearchPressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        fromDate = self.view.fSlashDate(self.fromDate.date())
        toDate = self.view.fSlashDate(self.toDate.date())
        fromFormat = datetime.datetime.strptime(fromDate, "%m/%d/%Y").date()
        toFormat = datetime.datetime.strptime(toDate, "%m/%d/%Y").date()
        if fromFormat <= toFormat:
            cultureData = self.model.findSamplesQA('Cultures', '[SampleID], [Type], [Clinician], [Tech], [Received], [Reported]', fromDate, toDate)
            catData = self.model.findSamplesQA('CATs', '[SampleID], [Type], [Clinician], [Tech], [Received], [Reported]', fromDate, toDate)
            data = cultureData + catData
            count = 0
            for tup in data:
                tup = list(tup)
                new = []
                new.append(tup[0])
                new.append(tup[1])
                clinician = self.model.findClinician(tup[2])
                new.append(self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3]))
                if tup[3] != 0:
                    techName = list(self.model.findTech(tup[3], 'First, Middle, Last'))
                    new.append(techName[0] + ' ' + techName[1] + ' ' + techName[2])
                else:
                    new.append("")
                new.append(self.view.fSlashDate(tup[4])) if tup[4] != None else new.append("None")
                new.append(self.view.fSlashDate(tup[5])) if tup[5] != None else new.append("None")
                data[count] = new
                count += 1
            data = sorted(data, key=lambda x: x[0])
            self.qaReportTable.setRowCount(len(data))
            for i in range(len(data)):
                self.qaReportTable.setItem(i, 0, QTableWidgetItem(str(data[i][0])))
                self.qaReportTable.setItem(i, 1, QTableWidgetItem(data[i][1]))
                self.qaReportTable.setItem(i, 2, QTableWidgetItem(data[i][2]))
                self.qaReportTable.setItem(i, 3, QTableWidgetItem(data[i][3]))
                self.qaReportTable.setItem(i, 4, QTableWidgetItem(data[i][5]))
                if str(data[i][5]) != 'None' and str(data[i][4]) != 'None':
                    numDays = (datetime.datetime.strptime(data[i][5], '%m/%d/%Y').date() - datetime.datetime.strptime(data[i][4], '%m/%d/%Y').date()).days
                else:
                    numDays = 'Still in Culture'
                self.qaReportTable.setItem(i, 5, QTableWidgetItem(str(numDays)))
            self.qaReportTable.sortItems(0,0)
            self.qaReportTable.resizeColumnsToContents()
            self.view.auditor(self.model.getCurrUser(), 'Search', 'COMBDb', 'QAReport')
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("From date must come before to date") 

    #@throwsViewableException
    def handlePrintPressed(self):
        self.printList = []
        for i in range(self.qaReportTable.rowCount()):
            self.printList.append([str(self.qaReportTable.item(i, 0).text()), self.qaReportTable.item(i, 1).text(), self.qaReportTable.item(i, 2).text(), self.qaReportTable.item(i, 3).text(), self.qaReportTable.item(i, 4).text(), self.qaReportTable.item(i, 5).text()])
        template = str(Path().resolve())+r'\templates\qa_report_template.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        document.merge(
            tech=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
        )
        document.write(dst)
        context = {
            'headers1' : ['Sample ID', 'Type', 'Clincian', 'Tech', 'Date Reported', 'Days in Culture'],
            'servers1' : self.printList
        }
        document = DocxTemplate(dst)
        document.render(context)
        document.save(dst)
        self.view.convertAndPrint(dst)
        self.view.auditor(self.model.getCurrUser(), 'Print', 'COMBDb', 'QAReport')

    #@throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    #@throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")