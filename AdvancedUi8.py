# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Users\khossei4\Desktop\Ui Files\AdvancedUi.ui'
#
# Created by: PyQt5 UI code generator 5.12.2
#
# WARNING! All changes made in this file will be lost!

import pandas               as pd  
import matplotlib.pyplot    as plt     
import numpy                as np 
import urllib.request
import os.path

from PyQt5              import QtCore, QtGui, QtWidgets
from datetime           import date,datetime
    
from EmployementPage3   import Ui_MainWindow1
from MigrationPage      import Ui_MigrationWindow
from NaturalGasPage     import Ui_NaturalGasPage
from GDP                import Ui_GDPWindow
from Electricity        import Ui_ElectricityWindow


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1147, 582)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.table1 = QtWidgets.QTableWidget(self.centralwidget)
        self.table1.setGeometry(QtCore.QRect(40, 220, 511, 81))
        self.table1.setShowGrid(False)
        self.table1.setGridStyle(QtCore.Qt.SolidLine)
        self.table1.setObjectName("table1")
        self.table1.setColumnCount(4)
        self.table1.setRowCount(1)
        item = QtWidgets.QTableWidgetItem()
        self.table1.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table1.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table1.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.table1.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.table1.setHorizontalHeaderItem(3, item)
        self.table1.horizontalHeader().setVisible(True)
        self.table1.horizontalHeader().setHighlightSections(True)
        self.table1.verticalHeader().setVisible(False)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(40, 160, 411, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(40, 320, 411, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.table2 = QtWidgets.QTableWidget(self.centralwidget)
        self.table2.setGeometry(QtCore.QRect(40, 390, 511, 81))
        self.table2.setShowGrid(False)
        self.table2.setGridStyle(QtCore.Qt.SolidLine)
        self.table2.setObjectName("table2")
        self.table2.setColumnCount(4)
        self.table2.setRowCount(1)
        item = QtWidgets.QTableWidgetItem()
        self.table2.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table2.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table2.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.table2.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.table2.setHorizontalHeaderItem(3, item)
        self.table2.horizontalHeader().setVisible(True)
        self.table2.horizontalHeader().setHighlightSections(True)
        self.table2.verticalHeader().setVisible(False)
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(590, 320, 411, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.table3 = QtWidgets.QTableWidget(self.centralwidget)
        self.table3.setGeometry(QtCore.QRect(590, 390, 511, 81))
        self.table3.setShowGrid(False)
        self.table3.setGridStyle(QtCore.Qt.SolidLine)
        self.table3.setObjectName("table3")
        self.table3.setColumnCount(4)
        self.table3.setRowCount(1)
        item = QtWidgets.QTableWidgetItem()
        self.table3.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table3.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table3.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.table3.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.table3.setHorizontalHeaderItem(3, item)
        self.table3.horizontalHeader().setVisible(True)
        self.table3.horizontalHeader().setHighlightSections(True)
        self.table3.verticalHeader().setVisible(False)
        self.table4 = QtWidgets.QTableWidget(self.centralwidget)
        self.table4.setGeometry(QtCore.QRect(590, 220, 511, 81))
        self.table4.setShowGrid(False)
        self.table4.setGridStyle(QtCore.Qt.SolidLine)
        self.table4.setObjectName("table4")
        self.table4.setColumnCount(4)
        self.table4.setRowCount(1)
        item = QtWidgets.QTableWidgetItem()
        self.table4.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table4.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table4.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.table4.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.table4.setHorizontalHeaderItem(3, item)
        self.table4.horizontalHeader().setVisible(True)
        self.table4.horizontalHeader().setHighlightSections(True)
        self.table4.verticalHeader().setVisible(False)
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(590, 160, 411, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.monthLbl = QtWidgets.QLabel(self.centralwidget)
        self.monthLbl.setGeometry(QtCore.QRect(390, 20, 500, 81))
        font = QtGui.QFont()
        font.setPointSize(48)
        self.monthLbl.setFont(font)
        self.monthLbl.setObjectName(date.today().strftime("%B"))
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1147, 26))
        self.menubar.setObjectName("menubar")
        self.summaryPageBtn = QtWidgets.QMenu(self.menubar)
        self.summaryPageBtn.setObjectName("summaryPageBtn")
        self.menuDetail_Specification = QtWidgets.QMenu(self.menubar)
        self.menuDetail_Specification.setObjectName("menuDetail_Specification")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.migrationPageBtn = QtWidgets.QAction(MainWindow)
        self.migrationPageBtn.setObjectName("migrationPageBtn")
        self.naturalPageBtn = QtWidgets.QAction(MainWindow)
        self.naturalPageBtn.setObjectName("naturalPageBtn")
        self.actionEmployement = QtWidgets.QAction(MainWindow)
        icon = QtGui.QIcon.fromTheme("ys")
        self.actionEmployement.setIcon(icon)
        self.actionEmployement.setObjectName("actionEmployement")
        self.actionGDP_Details = QtWidgets.QAction(MainWindow)
        self.actionGDP_Details.setObjectName("actionGDP_Details")
        self.actionPool_Price = QtWidgets.QAction(MainWindow)
        self.actionPool_Price.setObjectName("actionPool_Price")
        self.menuDetail_Specification.addAction(self.actionEmployement)
        self.menuDetail_Specification.addAction(self.migrationPageBtn)
        self.menuDetail_Specification.addAction(self.naturalPageBtn)
        self.menuDetail_Specification.addAction(self.actionGDP_Details)
        self.menuDetail_Specification.addAction(self.actionPool_Price)
        self.menubar.addAction(self.summaryPageBtn.menuAction())
        self.menubar.addAction(self.menuDetail_Specification.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
                                                                                                        
        self.actionEmployement.triggered.connect(self.openEmployementPage)
        self.migrationPageBtn.triggered.connect(self.openMigrationPage)
        self.naturalPageBtn.triggered.connect(self.openNaturalGasPage)
        self.actionGDP_Details.triggered.connect(self.openGDPpage)
        self.actionPool_Price.triggered.connect(self.openElectricityPage)

        self.dlUpdate()
        self.naturalGasTabel()
        self.GDPtable()
        self.netMigrationTable()
        self.EmployementTable()
    
    def retranslateUi(self, MainWindow):
            _translate = QtCore.QCoreApplication.translate
            MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
            item = self.table1.verticalHeaderItem(0)
            item.setText(_translate("MainWindow", "sadasdcas"))
            item = self.table1.horizontalHeaderItem(0)
            item.setText(_translate("MainWindow", "Trend"))
            item = self.table1.horizontalHeaderItem(1)
            item.setText(_translate("MainWindow", "2017"))
            item = self.table1.horizontalHeaderItem(2)
            item.setText(_translate("MainWindow", "2018"))
            item = self.table1.horizontalHeaderItem(3)
            item.setText(_translate("MainWindow", "% Change"))
            self.label_3.setText(_translate("MainWindow", "Natural Gas ($CDN/GJ)"))
            self.label.setText(_translate("MainWindow", "Gross Domestic Product at Basic Prices ($ billion)"))
            item = self.table2.verticalHeaderItem(0)
            item.setText(_translate("MainWindow", "sadasdcas"))
            item = self.table2.horizontalHeaderItem(0)
            item.setText(_translate("MainWindow", "Trend"))
            item = self.table2.horizontalHeaderItem(1)
            item.setText(_translate("MainWindow", "2017"))
            item = self.table2.horizontalHeaderItem(2)
            item.setText(_translate("MainWindow", "2018"))
            item = self.table2.horizontalHeaderItem(3)
            item.setText(_translate("MainWindow", "% Change"))
            self.label_2.setText(_translate("MainWindow", "Employment "))
            item = self.table3.verticalHeaderItem(0)
            item.setText(_translate("MainWindow", "sadasdcas"))
            item = self.table3.horizontalHeaderItem(0)
            item.setText(_translate("MainWindow", "Trend"))
            item = self.table3.horizontalHeaderItem(1)
            item.setText(_translate("MainWindow", "2017"))
            item = self.table3.horizontalHeaderItem(2)
            item.setText(_translate("MainWindow", "2018"))
            item = self.table3.horizontalHeaderItem(3)
            item.setText(_translate("MainWindow", "% Change"))
            item = self.table4.verticalHeaderItem(0)
            item.setText(_translate("MainWindow", "sadasdcas"))
            item = self.table4.horizontalHeaderItem(0)
            item.setText(_translate("MainWindow", "Trend"))
            item = self.table4.horizontalHeaderItem(1)
            item.setText(_translate("MainWindow", "2018"))
            item = self.table4.horizontalHeaderItem(2)
            item.setText(_translate("MainWindow", "2019"))
            item = self.table4.horizontalHeaderItem(3)
            item.setText(_translate("MainWindow", "% Change"))
            self.label_4.setText(_translate("MainWindow", "Net Migration"))
            self.monthLbl.setText(_translate("MainWindow", date.today().strftime('%B %Y')))
            self.summaryPageBtn.setTitle(_translate("MainWindow", "Summary"))
            self.menuDetail_Specification.setTitle(_translate("MainWindow", "Detail Specification"))
            self.menuHelp.setTitle(_translate("MainWindow", "Help"))
            self.migrationPageBtn.setText(_translate("MainWindow", "Migration Detail"))
            self.naturalPageBtn.setText(_translate("MainWindow", "Natural Gas Detail"))
            self.actionEmployement.setText(_translate("MainWindow", "Employement"))
            self.actionGDP_Details.setText(_translate("MainWindow", "GDP Details"))
            self.actionPool_Price.setText(_translate("MainWindow", "Pool Price"))


    def openGDPpage(self):
        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_GDPWindow()
        self.ui.setupUi(self.window)
        self.window.show()

    def openElectricityPage(self):
        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_ElectricityWindow()
        self.ui.setupUi(self.window)
        self.window.show()

    def openMigrationPage(self):
        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_MigrationWindow()
        self.ui.setupUi(self.window)
        self.window.show()

    def openEmployementPage(self):
        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_MainWindow1()
        self.ui.setupUi(self.window)
        self.window.show()

    def openNaturalGasPage(self):
        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_NaturalGasPage()
        self.ui.setupUi(self.window)
        self.window.show()
    
    def naturalGasTabel(self):
        naturalGas = Ui_NaturalGasPage()

        startDate = str(int(date.today().strftime('%Y'))-1)+"/1/1"
        


        startDate = str(startDate)
        startDate = datetime.strptime(startDate, '%Y/%m/%d')


        df2 = naturalGas.dataFrameReport(startDate,Type=3)

        

        year1 = str(int(df2['When'][df2.shape[0]-1].strftime('%Y'))-1)
        year2 = df2['When'][df2.shape[0]-1].strftime('%Y')
        month1 = df2['When'][df2.shape[0]-1].strftime('%B')
        month2 = df2['When'][df2.shape[0]-1].strftime('%B')
        
        startDate = str(int(date.today().strftime('%Y'))-1)+"/"+month1+"/1"

        
        df2 = df2[(df2['When'] >= startDate)]
        df2 = df2.reset_index(drop="True")


        change = naturalGas.getSpecific(year1,year2,month1,month2,df2,Type=3)

        self.table1.setHorizontalHeaderItem(1,QtWidgets.QTableWidgetItem(month1+" "+year1))
        self.table1.setHorizontalHeaderItem(2,QtWidgets.QTableWidgetItem(month2+" "+year2))

        
        # df1 data is inserted into the cells in the proper position
        firstYear = str(df2['Alberta'][0])
        self.table1.setItem(0,1, QtWidgets.QTableWidgetItem(firstYear))
        


        secondYear = str(df2['Alberta'][df2.shape[0]-1])
        self.table1.setItem(0,2, QtWidgets.QTableWidgetItem(secondYear))



        change = str(round((change-1)*100,1))
        self.table1.setItem(0,3, QtWidgets.QTableWidgetItem(change+"%"))

    def netMigrationTable(self):
        netMigration = Ui_MigrationWindow()

        startDate = str(int(date.today().strftime('%Y'))-1)+"/" + '1'+ "/1"
        


        startDate = str(startDate)
        startDate = datetime.strptime(startDate, '%Y/%m/%d')

        df2 = netMigration.dataFrameReport(startDate)

        month = df2['When'][df2.shape[0]-1].strftime('%m')
        

        startDate = str(int(date.today().strftime('%Y'))-1)+"/"+month+"/1"
        df2 = df2[(df2['When'] >= startDate)]
        
        change = netMigration.changeinMigration(df2)
        

        date1 = df2['When'][0]
        date1 = date1.strftime('%B %Y')


        date2 = df2['When'][df2.shape[0]-1]
        date2 = date2.strftime('%B %Y')


        # Headers for table 1 is set based on the last two year in the selected interval
        self.table4.setHorizontalHeaderItem(1,QtWidgets.QTableWidgetItem(date1))
        self.table4.setHorizontalHeaderItem(2,QtWidgets.QTableWidgetItem(date2))

        # df1 data is inserted into the cells in the proper position
        value1 = df2['Alberta'][0] 
        value1 = str(value1)
        self.table4.setItem(0,1, QtWidgets.QTableWidgetItem(value1))
        


        value1 = df2['Alberta'][df2.shape[0]-1]
        value1 = str(value1)
        self.table4.setItem(0,2, QtWidgets.QTableWidgetItem(value1))



        value1 = change
        value1 = str(round((value1 - 1) * 100,1))
        self.table4.setItem(0,3, QtWidgets.QTableWidgetItem(value1+"%"))
    def GDPtable(self):
        gdpPrint = Ui_GDPWindow()

        df1 = gdpPrint.returnDataframe()
        date1 = str(df1['When'][0].strftime('%B %Y'))
        date2 = str(df1['When'][1].strftime('%B %Y'))

        #Headers are populated using the most recent data collected from Alberta dashboard
        self.table2.setHorizontalHeaderItem(1,QtWidgets.QTableWidgetItem(date1))
        self.table2.setHorizontalHeaderItem(2,QtWidgets.QTableWidgetItem(date2))
       
       # df2 data is inserted into the cells in the proper position
        value1 = df1['Alberta'][df1.shape[0]-2] / 1000000000
        value1 = str(value1)
        self.table2.setItem(0,1, QtWidgets.QTableWidgetItem(value1))
        


        value1 = df1['Alberta'][df1.shape[0]-1] / 1000000000
        value1 = str(value1)
        self.table2.setItem(0,2, QtWidgets.QTableWidgetItem(value1))



        value1 = df1['Alberta'][df1.shape[0]-1] / df1['Alberta'][df1.shape[0]-2] 
        value1 = str(round((value1 - 1) * 100,1))
        self.table2.setItem(0,3, QtWidgets.QTableWidgetItem(value1+"%"))

    
    def dlUpdate(self):
        url = "https://economicdashboard.alberta.ca/Download/DownloadFile?extension=xlsx&requestUrl=https%3A%2F%2Feconomicdashboard.alberta.ca%2Fapi%2FNetMigration"
        urllib.request.urlretrieve(url, "Migration.xlsx" )

        url = "https://economicdashboard.alberta.ca/Download/DownloadFile?extension=xlsx&requestUrl=https%3A%2F%2Feconomicdashboard.alberta.ca%2Fapi%2FEmployment"
        urllib.request.urlretrieve(url, "Employement.xlsx" ) 

        url ='https://economicdashboard.alberta.ca/Download/DownloadFile?extension=xlsx&requestUrl=https%3A%2F%2Feconomicdashboard.alberta.ca%2Fapi%2FNaturalGasPrice'
        urllib.request.urlretrieve(url, "NaturalGasPrice.xlsx" )


        url = "https://economicdashboard.alberta.ca/Download/DownloadFile?extension=xlsx&requestUrl=https%3A%2F%2Feconomicdashboard.alberta.ca%2Fapi%2FGrossDomesticProduct"
        urllib.request.urlretrieve(url, "GDP.xlsx")

    def EmployementTable(self):
        
        
        emplTable = Ui_MainWindow1()

        

        df2         = emplTable.returndataFrame()
        df2         = df2.reset_index(drop="True")

        month1      = df2['When'][df2.shape[0]-1].strftime('%m')
        startDate   = str(int(date.today().strftime('%Y'))-1)+"/"+month1+"/1"
        endDate     = date.today().strftime('%Y')+"/"+month1+"/1"


        startDate   = str(startDate)

        endDate     =   datetime.strptime(endDate, '%Y/%m/%d')
        startDate   =   datetime.strptime(startDate, '%Y/%m/%d')

        df2         =   df2[(df2['When'] >= startDate) & (df2['When'] <= endDate)]
        df2         =   df2.reset_index(drop="True")


        #this data is collected for the table header
        date1       = startDate.strftime('%B %Y')
        date2       = endDate.strftime("%B %Y")

        #Headers are populated using the most recent data collected from Alberta dashboard
        self.table3.setHorizontalHeaderItem(1,QtWidgets.QTableWidgetItem(date1))
        self.table3.setHorizontalHeaderItem(2,QtWidgets.QTableWidgetItem(date2))

        #Table is updated
        firstYear = str(df2['Alberta'][0])
        self.table3.setItem(0,1, QtWidgets.QTableWidgetItem(firstYear))
        


        secondYear = str(df2['Alberta'][df2.shape[0]-1])
        self.table3.setItem(0,2, QtWidgets.QTableWidgetItem(secondYear))


        change = df2['Alberta'][df2.shape[0]-1]/df2['Alberta'][0]
        change = str(round((change-1)*100,1))
        self.table3.setItem(0,3, QtWidgets.QTableWidgetItem(change+"%"))


        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
