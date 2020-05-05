import sys
from PyQt5 import QtCore, QtGui, QtWidgets
import pandas as pd
from fuzzywuzzy import fuzz
from operation import Ui_MainWindow

class StringMatch(Ui_MainWindow):
    def __init__(self,MainWindow):
        Ui_MainWindow.__init__(self)
        self.setupUi(MainWindow)
        self.input()
    def input(self):
        self.pushButton.clicked.connect(self.simpleLabel)
        self.comboBox.activated.connect(self.comboBoxInput)
    def comboBoxInput(self):
        #Take input from dropbox and call for dimension match known as self.show()
        self.show(self.comboBox.currentText().lower())
    def simpleLabel(self):
        #Take input from LineEdit and call for dimension match known as self.show()
        self.show(self.lineEdit.text().lower())
    def show(self,value):
        self.entry=QtGui.QStandardItemModel()
        self.Distrubutor.setModel(self.entry)
        self.entry1=QtGui.QStandardItemModel()
        self.Company.setModel(self.entry1)
        #read excel file
        data=pd.read_excel("./Mapped sku.xlsx")
        for i in range(data.shape[0]):
            if data.iloc[i,0]!=0:
                #if match is above 95% then Data will be shown in listview
                if fuzz.partial_ratio(value,str(data.iloc[i,0]).lower())>=95:
                    #Distributor code list with matching dimension
                    item=QtGui.QStandardItem(data.iloc[i,0])
                    self.entry.appendRow(item)
                if fuzz.partial_ratio(value,str(data.iloc[i,1]).lower())>=95:
                    #company description list with matching dimension
                    item1=QtGui.QStandardItem(data.iloc[i,1])
                    self.entry1.appendRow(item1)
            # show status in percentage 100% indicate matching is done
            self.progressBar.setProperty("value",((i+1)/data.shape[0])*100)
            # if listview is empty then select other dimension
if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    prog = StringMatch(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())