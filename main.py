# -*- coding: utf-8 -*-
"""
Created on Mon Jun 19 10:37:33 2017
Changed on Wed Okt 04 14:49:00 2017

@author: Paul Grunert
"""

from PyQt5 import QtWidgets, QtCore, uic, QtGui
from xlrd import open_workbook, xldate
from xlutils.copy import copy
from operator import itemgetter
from sys import argv
import webbrowser

from matplotlib.pyplot import figure
from urllib.request import urlopen

from datetime import datetime
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
from matplotlib.dates import YearLocator, DateFormatter, MonthLocator
#import matplotlib.dates as mdates

class AweFinApp(QtWidgets.QMainWindow):
    # Class Variables
    data_list = []
    account_list = []
    filename = ""
    AweFinVersion = "0.2"
    # Class Methods
    def __init__(self):
        super(self.__class__, self).__init__()
        uic.loadUi(".\\UI\\mainwindow.ui",self)
        # default values
        self.addDate.setDateTime(QtCore.QDateTime.currentDateTime());
        self.dabTimeEnd.setDateTime(QtCore.QDateTime.currentDateTime());
        # build up connections
        self.actionAbout.triggered.connect(self.showAbout)
        self.actionLoad_xlsx.triggered.connect(self.loadxlsx)
        self.actionSave_Data.triggered.connect(self.savexlsx)
        self.actionAddTransaction.clicked.connect(self.addTransaction)
        self.filterSelector.clicked.connect(self.dashboardQuery)
    def showAbout(self):
        self.w = QtWidgets.QWidget()
        a = QtWidgets.QLabel(self.w)
        a.setText("AweFin Version: " + self.AweFinVersion)
        a.move(10,10)
        a2 = QtWidgets.QLabel(self.w)
        f = urlopen("http://puhgee.de/awefin/version.html")
        ServerVersion = f.read()
        print(ServerVersion)
        a2.setText("Server Version: " + ServerVersion.decode("utf-8"))
        a2.move(10,30)
        b = QtWidgets.QPushButton(self.w)
        b.setText("Update")
        b.move(10,50)
        b.clicked.connect(self.updateAbout)
        c = QtWidgets.QPushButton(self.w)
        c.setText("Close")
        c.move(120,50)
        c.clicked.connect(self.hideAbout)
        self.w.setWindowTitle("About AweFin")
        self.w.show()
    def hideAbout(self):
        self.w.hide()
    def updateAbout(self):
        webbrowser.open('http://puhgee.de/images/ana2.png') 
    def loadxlsx(self):
        self.filename = QtWidgets.QFileDialog.getOpenFileName(self, "Datei wählen")
        self.statusbar.showMessage("Loading: " + self.filename[0])
        self.show()
        book = open_workbook(self.filename[0])
        sheet = book.sheet_by_index(0)
        item_model = QtGui.QStandardItemModel(self.accountSelector)
        filter_model = QtGui.QStandardItemModel(self.filterSelector)
        filter_model.appendRow(QtGui.QStandardItem("Basic cumulative"))
        filter_model.appendRow(QtGui.QStandardItem("Basic category shares income"))
        filter_model.appendRow(QtGui.QStandardItem("Basic category shares expenditures"))
        for row_index in range(1, sheet.nrows):
            dateinfo = sheet.cell_value(row_index, 2)
            if isinstance(dateinfo, float):
                dateinfo = xldate.xldate_as_datetime(sheet.cell_value(row_index, 2),0)
                dateinfo = dateinfo.strftime("%d.%m.%Y")
            self.data_list.append((sheet.cell_value(row_index, 0),sheet.cell_value(row_index, 1), dateinfo,sheet.cell_value(row_index, 3),sheet.cell_value(row_index, 4),sheet.cell_value(row_index, 5),sheet.cell_value(row_index, 6),sheet.cell_value(row_index, 7),sheet.cell_value(row_index, 8),sheet.cell_value(row_index, 9),sheet.cell_value(row_index, 10)))
            if(sheet.cell_value(row_index, 9)not in self.account_list):
                self.account_list.append(sheet.cell_value(row_index, 9))
                item_model.appendRow(QtGui.QStandardItem(sheet.cell_value(row_index, 9)))
        self.data_list.sort(key=lambda tup: datetime.strptime(tup[2],"%d.%m.%Y"))
        table_model = TableModel(self, self.data_list, ['Transaktions_ID','geprüft','Datum','Beschreibung','Wert','Kategorie1','Kategorie2','Kategorie3','Laufend','Konto','Soll/Haben'])
        self.filterSelector.setModel(filter_model)
        self.accountSelector.setModel(item_model)
        self.transactionList.setModel(table_model)
        self.refreshKPI()
        self.statusbar.showMessage("Data loaded!")
        self.show()
    def savexlsx(self):
        book = copy(open_workbook(self.filename[0]))
        sheet = book.get_sheet(0)
        for row_index in range(1, len(self.data_list)+1):
            sheet.write(row_index, 0, self.data_list[row_index-1][0])
            sheet.write(row_index, 1, self.data_list[row_index-1][1])
            sheet.write(row_index, 2, self.data_list[row_index-1][2])
            sheet.write(row_index, 3, self.data_list[row_index-1][3])
            sheet.write(row_index, 4, self.data_list[row_index-1][4])
            sheet.write(row_index, 5, self.data_list[row_index-1][5])
            sheet.write(row_index, 6, self.data_list[row_index-1][6])
            sheet.write(row_index, 7, self.data_list[row_index-1][7])
            sheet.write(row_index, 8, self.data_list[row_index-1][8])
            sheet.write(row_index, 9, self.data_list[row_index-1][9])
            sheet.write(row_index, 10, self.data_list[row_index-1][10])
        if self.filename[0][-1]=="x":
            book.save(self.filename[0][:-1])    
        else:
            book.save(self.filename[0])
        self.statusbar.showMessage("Data saved!")
        self.show()
    def addTransaction(self):
        transID = len(self.data_list)
        transChecked =  "nein"
        if self.addChecked.isChecked() == True:
            transChecked = "ja"
        transDate = self.addDate.date().toString("dd.MM.yyyy")
        transDescription = self.addDescription.text()
        transValue = self.addValue.value()
        transCat1 = self.addCat1.text()
        transCat2 = self.addCat2.text()
        transCat3 = self.addCat3.text()
        transRun = "nein"
        if self.addRun.isChecked() == True:
            transRun = "ja"
        if transValue >= 0:
            transSkont = "haben"
        else:
            transSkont = "soll"
        if (len(self.accountSelector.selectedIndexes())<=0):
            return
        transAccount = self.accountSelector.model().data(self.accountSelector.selectedIndexes()[0])
        self.data_list.append((transID, transChecked, transDate, 
                               transDescription, transValue, transCat1,
                               transCat2, transCat3, transRun, transAccount,
                               transSkont))
        table_model = TableModel(self, self.data_list,
                                 ['Transaktions_ID', 'geprüft', 'Datum',
                                  'Beschreibung', 'Wert', 'Kategorie1',
                                  'Kategorie2', 'Kategorie3', 'Laufend',
                                  'Konto', 'Soll/Haben'])
        self.transactionList.setModel(table_model)
        self.refreshKPI()
        self.statusbar.showMessage("Transaction " + transDescription + " added!")
        self.show()
    def refreshKPI(self):
        KPIlist = []
        for account in self.account_list:
            accountValue = 0
            for datapoint in self.data_list:
                if (datapoint[9] == account):
                    accountValue += datapoint[4]
            KPIlist.append((account, accountValue))
        KPImodel = TableModel(self, KPIlist, ['Account','Value'])
        
        figure2 = figure()
        figure2.clear()
        ax = figure2.add_subplot(1, 1, 1)
        ax.pie([x[1] for x in KPIlist],labels=[x[0] for x in KPIlist])
        canvas = FigureCanvasQTAgg(figure2)
        canvas.draw()
        for i in range(self.KPIcanvas.count()): self.KPIcanvas.itemAt(i).widget().close()
        self.KPIcanvas.addWidget(canvas)
        self.KPIlist.setModel(KPImodel)
        self.show()
    def dashboardQuery(self):
        selectedFilter = self.filterSelector.model().data(self.filterSelector.selectedIndexes()[0])
        print("Updating dashboard query with Filter " + selectedFilter)
        figure2 = figure()
        figure2.clear()
        ax = figure2.add_subplot(1, 1, 1)
        ax.xaxis.set_major_locator(YearLocator())
        ax.xaxis.set_major_formatter(DateFormatter('%Y'))
        ax.xaxis.set_minor_locator(MonthLocator())
        datalist = []
        
        if selectedFilter == "Basic cumulative":
            for account in self.account_list:
                itemsum = 0
                for datapoint in self.data_list:
                    if (datapoint[9] == account):
                        itemdate = datetime.strptime(datapoint[2],"%d.%m.%Y")
                        itemsum += datapoint[4]
                        if ((itemdate > self.dabTimeBegin.date()) & (itemdate < self.dabTimeEnd.date())):
                            datalist.append((itemdate, itemsum))
                ax.plot([x[0] for x in datalist],[x[1] for x in datalist])
                datalist = []
        elif selectedFilter == "Basic category shares income":
            catlist = []
            catvaluelist = []
            for datapoint in self.data_list:           
                itemdate = datetime.strptime(datapoint[2],"%d.%m.%Y")
                if ((itemdate > self.dabTimeBegin.date()) & (itemdate < self.dabTimeEnd.date())):
                    if datapoint[4] >=0:
                        if(datapoint[5] not in catlist):
                            catlist.append(datapoint[5])
                            catvaluelist.append(datapoint[4])
                        else:
                            catvaluelist[catlist.index(datapoint[5])] += datapoint[4]
            ax.pie(catvaluelist, labels=catlist)
        elif selectedFilter == "Basic category shares expenditures":
            catlist = []
            catvaluelist = []
            for datapoint in self.data_list:           
                itemdate = datetime.strptime(datapoint[2],"%d.%m.%Y")
                if ((itemdate > self.dabTimeBegin.date()) & (itemdate < self.dabTimeEnd.date())):
                    if datapoint[4] <=0:
                        if(datapoint[5] not in catlist):
                            catlist.append(datapoint[5])
                            catvaluelist.append(abs(datapoint[4]))
                        else:
                            catvaluelist[catlist.index(datapoint[5])] += abs(datapoint[4])
            ax.pie(catvaluelist, labels=catlist)
        canvas = FigureCanvasQTAgg(figure2)
        canvas.draw()
        for i in range(self.dabQueryCanvas.count()): self.dabQueryCanvas.itemAt(i).widget().close()
        self.dabQueryCanvas.addWidget(canvas)
        self.show()

class TableModel(QtCore.QAbstractTableModel):
    def __init__(self, parent, mylist, header, *args):
        QtCore.QAbstractTableModel.__init__(self, parent, *args)
        self.mylist = mylist
        self.header = header
    def rowCount(self, parent):
        return len(self.mylist)
    def columnCount(self, parent):
        return len(self.mylist[0])
    def data(self, index, role):
        if not index.isValid():
            return None
        elif role != QtCore.Qt.DisplayRole:
            return None
        return self.mylist[index.row()][index.column()]
    def headerData(self, col, orientation, role):
        if orientation == QtCore.Qt.Horizontal and role == QtCore.Qt.DisplayRole:
            return self.header[col]
        return None
    def sort(self, col, order):
        """sort table by given column number col"""
        #self.emit(QtCore.SIGNAL("layoutAboutToBeChanged()"))
        self.mylist = sorted(self.mylist,
            key = itemgetter(col))
        if order == QtCore.Qt.DescendingOrder:
            self.mylist.reverse()
        self.layoutChanged.emit()
        
def main():
    app = QtWidgets.QApplication(argv)  # A new instance of QApplication
    form = AweFinApp()                 # We set the form to be our ExampleApp (design)
    form.show()                         # Show the form
    app.exec_()                         # and execute the app


if __name__ == '__main__':              # if we're running file directly and not importing it
    main()                              # run the main function
