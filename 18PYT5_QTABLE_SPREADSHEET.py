import sys
import os
import csv
from PyQt5.QtWidgets import (QTableWidget, QApplication, QMainWindow, QTableWidgetItem, QFileDialog)
from PyQt5.QtWidgets import (QAction, qApp)


class MyTable(QTableWidget):
    def __init__(self, r, c):
        super().__init__(r, c)
        self.check_change = True
        self.init_ui()

    def init_ui(self):
        self.cellChanged.connect(self.cell_current)

    def cell_current(self):
        if self.check_change:
            row = self.currentRow()
            col = self.currentColumn()
            value = self.item(row, col)
            value = value.text()
            print("CURRENT CELL: ", row, ",", col)
            print("CELL VALUE: ", value)

            self.show()

    def open_sheet(self):
        self.check_change = False
        path = QFileDialog.getOpenFileName(self, 'Open CSV', os.getenv('HOME'), 'CSV(*.csv)')
        if path[0] != '':
            with open(path[0], newline='') as csv_file:
                self.setRowCount(0)
                self.setColumnCount(10)
                my_file = csv.reader(csv_file, dialect='excel')
                for row_data in my_file:
                    row = self.rowCount()
                    self.insertRow(row)
                    if len(row_data) > 10:
                        self.setColumnCount(len(row_data))
                    for column, staff in enumerate(row_data):
                        item = QTableWidgetItem(staff)
                        self.setItem(row, column, item)
        self.check_change = True

    def save_sheet(self):
        path = QFileDialog.getSaveFileName(self, 'Save CSV', os.getenv('HOME'), 'CSV(*.csv)')
        if path[0] != '':
            with open(path[0], 'w') as csv_file:
                writer = csv.writer(csv_file, dialect='excel')
                for row in range(self.rowCount()):
                    row_data = []
                    for column in range(self.columnCount()):
                        item = self.item(row, column)
                        if item is not None:
                            row_data.append(item.text())
                        else:
                            row_data.append('')
                    writer.writerow(row_data)


class Sheet(QMainWindow):
    def __init__(self):
        super().__init__()

        self.form_widget = MyTable(10, 10)
        self.setCentralWidget(self.form_widget)
        col_headers = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
        self.form_widget.setHorizontalHeaderLabels(col_headers)

        # create menu bar
        bar = self.menuBar()

        # create root menu
        file = bar.addMenu('File')

        # create action for menus
        save_action = QAction('Save', self)
        save_action.setShortcut('Ctrl+S')

        # new_action = QAction('New', self)
        # new_action.setShortcut('Ctrl+N')

        open_action = QAction('Open', self)
        open_action.setShortcut('Ctrl+O')

        quit_action = QAction('Quit', self)
        quit_action.setShortcut('Ctrl+Q')

        # add action to menus
        # file.addAction(new_action)
        file.addAction(save_action)
        file.addAction(open_action)
        file.addAction(quit_action)

        # events
        quit_action.triggered.connect(self.quit_trigger)
        save_action.triggered.connect(self.form_widget.save_sheet)
        open_action.triggered.connect(self.form_widget.open_sheet)

        self.show()

    def quit_trigger(self):
        qApp.quit()


app = QApplication(sys.argv)
spreadsheet = Sheet()
sys.exit(app.exec_())
