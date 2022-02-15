from PyQt5.QtWidgets import *

class View:
  def __init__(self, sys):
      self.sys = sys
      self.app = QApplication(sys.argv)

      self.window = QWidget()
      self.window.setWindowTitle('PyQt5 App')
      self.window.setGeometry(100, 100, 280, 80)
      self.window.move(60, 15)
      helloMsg = QLabel('<h1>Maik BIG gay!</h1>', parent=self.window)
      helloMsg.move(60, 15)
      #self.window.show()

  def display(self):
      self.window.show()
      self.sys.exit(self.app.exec_())