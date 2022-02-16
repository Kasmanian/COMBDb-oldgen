from PyQt5.QtWidgets import *
import sys

class View(QApplication):
  def __init__(self):
      super(View, self).__init__(sys.argv)
      self.win = Window()

  def init(self):
      sys.exit(self.exec_())


class Window(QWidget):
  def __init__(self):
      super(Window, self).__init__()
      self.initUI()

  def initUI(self):
      self.setWindowTitle('OMCDb')
      self.showMaximized()