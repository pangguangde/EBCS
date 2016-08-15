# _*_ coding: utf-8
__author__ = 'guangde'

import csv
import traceback
import xlsxwriter
import datetime

from PyQt4 import QtCore, QtGui, uic
from os.path import isfile
import codecs


from PyQt4.QtWebKit import*  

from PyQt4.QtCore import *  
from PyQt4.QtGui import *  
from PyQt4.QtWebKit import *  
from PyQt4.QtNetwork import *  
import sys

qtCreatorFile = "test.ui" # Enter file here.

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

class EmittingStream(QtCore.QObject):  
        textWritten = QtCore.pyqtSignal(str)  
        def write(self, text):  
            self.textWritten.emit(str(text)) 

class MyApp(QtGui.QMainWindow, Ui_MainWindow):
	def __init__(self):
		QtGui.QMainWindow.__init__(self)
		Ui_MainWindow.__init__(self)
		self.setupUi(self)
		self.label.setToolTip(u'    1. 文件名有命名规范\n        中通和申通外部的文件名应为 ‘外部中通/申通.csv’\n        系统内部的文件名应为 ‘系统混合.csv’\n\n    2. csv 文件的列也有要求\n        ‘外部中通/申通.csv’ 文件从左到右依次为 运单号、省份、重量、价格\n        ‘系统混合.csv’     文件从左到右依次为 运单号、省份、重量、快递公司名称')

		sys.stdout = EmittingStream(textWritten=self.normalOutputWritten)

	def normalOutputWritten(self, text):  
		cursor = self.outText.textCursor()  
		cursor.movePosition(QtGui.QTextCursor.End)  
		cursor.insertText(text)  
		self.outText.setTextCursor(cursor)  
		self.outText.ensureCursorVisible() 


if __name__ == "__main__":
	app = QtGui.QApplication(sys.argv)
	window = MyApp()
	print 'test, ahahahhahahaahha'
	window.show()
	sys.exit(app.exec_())