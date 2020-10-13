#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys
from PyQt5.QtWidgets import (QWidget, QLabel, QLineEdit,
                             QTextEdit, QGridLayout, QApplication, QPushButton, QFileDialog, QMessageBox)
import calendar
import xlsxwriter
import datetime

from qtpy import QtGui


class Example(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.rir = []
        now = datetime.datetime.now()
        holidays = {datetime.date(now.year, 8, 14)}  # you can add more here
        businessdays = 0
        for i in range(1, 32):
            try:
                thisdate = datetime.date(now.year, now.month, i)
            except(ValueError):
                break
            if thisdate.weekday() < 5 and thisdate not in holidays:  # Monday == 0, Sunday == 6
                businessdays += 1
                self.rir.append(thisdate.day)
        self.mth = now.strftime("%m")
        self.yr = now.strftime("%Y")
        self.title = QLabel('Всего дней в месяце: '+ str(calendar.monthrange(int(self.yr), int(self.mth))[1]) )
        self.author = QLabel('Всего рабочих дней: '+ str(businessdays))
        self.review = QLabel('Числа тех дней по которым работаем: ' + str(self.rir))
        self.btn = QPushButton("Создание отчёта", self)
        self.btn.clicked.connect(self.buttonClicked)
        grid = QGridLayout()
        grid.setSpacing(10)
        grid.addWidget(self.title, 1, 0)
        grid.addWidget(self.author, 2, 0)
        grid.addWidget(self.review, 3, 0)
        grid.addWidget(self.btn, 4 , 0 )
        self.setLayout(grid)
        self.setGeometry(300, 300, 150, 100)
        self.setWindowTitle('Генератор отчётов')
        self.show()


    def  buttonClicked(self):
        try:
          options = QFileDialog.Options()
          options |= QFileDialog.DontUseNativeDialog
          fileName, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","MS Excel Files (*.xlsx);;All Files (*)", options=options)
          workbook = xlsxwriter.Workbook(fileName + '.xlsx')
          worksheet = workbook.add_worksheet()
          worksheet.set_column('A:A', 30)
          worksheet.set_column('B:B',120)
          bold = workbook.add_format({'bold': True})
          bold.set_font_color('red')
          cell_format = workbook.add_format({'bold': True, 'bg_color': 'black'})
          cell_format2 = workbook.add_format({'bold': True, 'bg_color': '#e1f6fa '})
          i = 0
          i2 = 13
          worksheet.write('A' + str(i + 1),str(self.rir[0])+ "." + self.mth + "." + self.yr,bold)
          worksheet.write('B' + str(i + 1),'',cell_format)
          i = i+1
          ti = 0
          toto = 0
          merge_format = workbook.add_format({
            'bold': True,
            'border': 4,
            'align': 'left',
            'valign': 'vcenter',
            'fg_color': '#e1f6fa',
          })
          time_massive = ['10-11:' ,'11-12:' ,'12-12:30: – обед' ,'12:30-13:30:' ,'13:30-14:30:' ,'14:30-15:30:' ,'15:30-16:00: перерыв','16-17:' ,'17-18:', 'Протоколы:', 'Изучение библиотеки:', 'Разработка программы:']
          while  i+1 <= len(self.rir):
            worksheet.write('A' + str(i + i2 + 1), str(self.rir[i]) + "." + self.mth + "." + self.yr,bold)
            worksheet.write('B' + str(i + i2 + 1),'', cell_format)
            ti = 0
            while ti + 1 <= len(time_massive):
                 worksheet.write('A' + str(i + i2 + 1 + ti + 1), time_massive[ti])
                 worksheet.write('A' + str(ti+2), time_massive[ti])
                 worksheet.write('B' + str(i + i2 + 1 + ti + 1), '',merge_format)
                 worksheet.write('B' + str(ti+2), '',merge_format)
                 ti = ti + 1
                 toto = toto + 1
            toto = toto + 1
            i = i +1
            i2 = i2 + 13
          i = 0
          workbook.close()
          Qmessa("Готово")
        except:
          Qmessa("Произошла ошибка")

class Qmessa(QMessageBox):
    def __init__(self, a):
        super().__init__()
        msg = QMessageBox()
        msg.setWindowTitle("Info")
        msg.setText(a)
        result = msg.setStandardButtons(QMessageBox.Ok)
        retval = msg.exec_()



if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())