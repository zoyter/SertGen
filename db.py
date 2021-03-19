import sqlite3
import sys

from PyQt5 import uic
from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QApplication


class frmDB(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('ui/frmDB.ui', self)  # Грузим интерфейс главной формы
        self.set_window_on_center_screen()  # Центрируем окно
        self.load_config()
        # Привязываем обрабочики
        self.pushButton.clicked.connect(self.onClick)

        self.rbName.clicked.connect(self.onQueryClick)
        self.rbDate.clicked.connect(self.onQueryClick)
        self.rbEvent.clicked.connect(self.onQueryClick)
        self.rbSN.clicked.connect(self.onQueryClick)

        self.column = "name"

    def load_config(self):
        self.db = "db.sqlite"
        self.table = "info"

    def set_window_on_center_screen(self):
        # Окно программы в центр экрана
        screen_geometry = QApplication.desktop().availableGeometry()
        screen_size = (screen_geometry.width(), screen_geometry.height())
        win_size = (self.frameSize().width(), self.frameSize().height())
        x = screen_size[0] // 2 - win_size[0] // 2
        y = screen_size[1] // 2 - win_size[1] // 2
        self.move(x, y)

    def onClick(self):
        query = "SELECT * FROM " + self.table
        if self.txtQuery.text():
            query += " WHERE "
            query += self.column
            query += " LIKE '%" + self.txtQuery.text() + "%'"
        self.get_data(query)

    def onQueryClick(self):
        self.column = self.sender().objectName()
        if self.sender().objectName() == "rbName":
            self.column = "name"
        elif self.sender().objectName() == "rbDate":
            self.column = "date"
        elif self.sender().objectName() == "rbEvent":
            self.column = "event"
        elif self.sender().objectName() == "rbSN":
            self.column = "sn"
        else:
            self.column = "name"

    def get_data(self, query):
        self.tableWidget.clear()
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)

        self.con = sqlite3.connect(self.db)
        self.cur = self.con.cursor()

        result = self.cur.execute(query).fetchall()
        if result:
            headers = list(map(lambda x: x[0], self.cur.description))
            self.tableWidget.setColumnCount(len(result[0]))
            self.tableWidget.setRowCount(len(result))
            self.tableWidget.setHorizontalHeaderLabels(headers)
            for i in range(len(result)):
                for j in range(len(result[0])):
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(result[i][j])))
        self.con.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = frmDB()
    ex.show()
    # ex.con.close()
    sys.exit(app.exec_())
