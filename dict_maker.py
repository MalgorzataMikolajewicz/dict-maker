import sys
from PySide2 import QtCore, QtWidgets, QtGui
import os
import numpy as np
import pandas as pd
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import openpyxl

class DateInput:
    def __init__(self, enter):
        self.enter = enter


class AppWidget(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()  
        self.state = []
        self.layout = self.initialize_layout()             
        self.setLayout(self.layout)
        
    def add_months(self, entered_date):
        months_list = []
        new = pd.Timestamp(entered_date) + relativedelta(years=1, months=1)
        for i in range(60):
            months_list.append(new - relativedelta(months=1))
            new -= relativedelta(months=1)
        lp = []
        mat = []
        n = 0
        m = 0
        for month in months_list:
            n += 1
            lp.append(n)
        for l in lp:
            if l/12 <= m:
                mat.append("MAT{}".format(m-1))
            else:
                m += 1
                mat.append("MAT{}".format(m-1))
        lp = []
        r_quater = []
        n = 0
        for month in months_list:
            n += 1
            lp.append(n)
        q = [1,1,1,2,2,2,3,3,3,4,4,4]*(len(lp)//12)
        for l in lp:
            r_quater.append("RQ{}".format(q[l-1]))

        year = []
        for month in pd.DatetimeIndex(months_list):
            year.append(month.year)

        quater = []
        for month in pd.DatetimeIndex(months_list):
            quater.append("Q{}".format((month.month-1)//3+1))

        half = []
        for month in pd.DatetimeIndex(months_list):
            half.append("H{}".format(month.month//7+1))

        ytd = []
        lastdata = pd.DatetimeIndex(months_list)[0].year
        for month in pd.DatetimeIndex(months_list):
            ytd.append("YTD{}".format(lastdata-month.year))

        dict_mat = pd.DataFrame({'Time period': months_list, 'YEAR': year,'HALF': half, 'QUATER': quater,'MAT': mat, 'YTD': ytd, 'ROLLING QUATER': r_quater})

        dict_mat.iloc[0:12,4:] = np.nan
        rows_to_clean = int(entered_date[-2:])
        dict_mat.iloc[12 + rows_to_clean:24,5] = np.nan
        dict_mat.iloc[24 + rows_to_clean:36,5] = np.nan
        dict_mat.iloc[36 + rows_to_clean:48,5] = np.nan
        dict_mat.iloc[48 + rows_to_clean:60,5] = np.nan

        dict_mat.to_excel("dict.xlsx") 
        
    def on_submit(self):
        for i in self.state:
            print("Dictionary is being generated. Check your folder.")
            entered_date = i.enter.text()
            self.add_months(entered_date)
            sys.exit(0)
        
    def initialize_layout(self):
        path = os.getcwd()
        path = os.path.realpath(path)
        os.startfile(path)
        label = QtWidgets.QLabel('Place the files into the folder that opened, then enter the starting month in yyyy-mm format and press "OK". ')
        label.setFont(QtGui.QFont('Sanserif', 11))
        enter = QtWidgets.QLineEdit()
        row = 0
        grid = QtWidgets.QGridLayout()
        submit = QtWidgets.QPushButton('OK')
        submit.clicked.connect(self.on_submit)
        self.state.append(DateInput(enter))
        grid.addWidget(label, row, 0)
        grid.addWidget(enter, row, 1)
        grid.addWidget(submit, row + 1, 1)
        return grid
        

if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    appWidget = AppWidget()
    appWidget.resize(600, 80)
    appWidget.show()
    sys.exit(app.exec_())


    