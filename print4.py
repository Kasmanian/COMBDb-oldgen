import os
from PyQt5 import QtWebEngineWidgets, QtCore, QtPrintSupport
# ...
def run(self):
    current_dir = os.path.dirname(os.path.abspath(__file__))
    self._page = QtWebEngineWidgets.QWebEnginePage()
    self._page.setHtml(QtCore.QUrl.fromLocalFile(r"C:\Users\simmsk\Desktop\temp.html"))
    self._printer = QtPrintSupport.QPrinter()
    self._printer.setPaperSize(QtCore.QSizeF(80 ,297), QtPrintSupport.QPrinter.Millimeter)
    r = QtPrintSupport.QPrintDialog(self._printer)
    if r.exec_() == QtPrintSupport.QPrintDialog.Accepted:
        self._page.print(self._printer, self.print_completed)

if __name__ == '__main__':
    run()