import sys, os
import platform, subprocess
import webbrowser
import sqlite3
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QDialog, QMessageBox
from PyQt5.QtGui import QPixmap, QMovie
from PyQt5.QtWidgets import QTableWidgetItem
import xlrd

# т.к. в Windows cairosvg не работает,
isLinux = True
if platform.system() == "Windows": # Если виндовс, то будем работать через inkscape
    isLinux = False
else: # если остальное, то пробуем импортировать модуль
    import cairosvg


import db  # второе окно для просмотра базы и поиска в ней информации


def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)


class frmMain(QMainWindow):
    def __init__(self):
        super().__init__()

        uic.loadUi('ui/frmMain.ui', self)  # Грузим интерфейс главной формы
        self.set_window_on_center_screen()  # Центрируем окно
        self.load_config()  # Загрузка конфига из файла

        # Привязываем обработчики
        self.btnOpenDB.clicked.connect(self.onOpenDBClick)
        self.btnNext.clicked.connect(self.onClickNext)
        self.btnBack.clicked.connect(self.onClickBack)

        self.lstSVGFiles.itemClicked.connect(self.onSVGFileClicked)
        self.tabWidget.currentChanged.connect(self.onTabWidgetClick)
        self.btnMake_all_sert.clicked.connect(self.make_all_sert)
        self.lstXLSFiles.itemClicked.connect(self.onXLSFileClicked)

        self.btnOpenResults.clicked.connect(self.onOpenResultClick)
        self.btnOpenTemplatesFolder.clicked.connect(self.onOpenTemplatesFolderClick)
        self.btnOpenXLSFolder.clicked.connect(self.onOpenXLSFolderClick)
        self.btnHelp.clicked.connect(self.onHelptClick)

        self.txtInkscape.textChanged.connect(self.onTextChanged)

        # Донастраиваем всякие мелочи
        self.noimage = QPixmap(self.img + "/noimage.png")  # Картинка, если нет картинки
        self.ok_img = QPixmap(self.img + "/check.png")  # Картинка, если нет картинки
        self.no_img = QPixmap(self.img + "/remove.png")  # Картинка, если нет картинки

        self.current_step = 0  # Текущий шаг в мастере
        self.tabWidget.setCurrentIndex(self.current_step)  # Выставляем нужную вкладку в мастере
        self.is_table_data_loaded = False

        # Декативируем лишние вкладки
        for i in range(1, self.tabWidget.count()):
            self.tabWidget.setTabEnabled(i, False)

        # Проверяем наоличия конвертера в PDF (для windows)
        if isLinux:
            # Убираем поля связанные с inkscape
            self.txtInkscape.hide()
            self.lblInkscape.hide()
            self.lblInkscapeIcon.hide()
        else:
            if not os.path.exists(self.inkscape):
                self.lblInkscapeIcon.setPixmap(self.no_img)
            else:
                self.lblInkscapeIcon.setPixmap( self.ok_img)

    def onOpenDBClick(self): # Открываем второе окно для просмотра базы и поиска по ней
        self.frmDB = db.frmDB()
        self.frmDB.show()
    def onTextChanged(self):
        self.inkscape = self.txtInkscape.text()
        print(self.inkscape)
        if os.path.exists(self.inkscape):
            self.lblInkscapeIcon.setPixmap(self.ok_img)
        else:
            self.lblInkscapeIcon.setPixmap(self.no_img)
            print("ERROR")

    def load_config(self): # TODO: сделать загрузку конфига из файла
        self.db = "db.sqlite"
        self.table = "info"
        self.templates = "03.templates"  # Каталог с шаблонами сертификатов
        self.img = "img"  # Каталог с картинками к программе
        self.tmp = "tmp"  # Временные файлы
        self.tabledata = "01.input"  # Каталог с файлами xlx
        self.result = "02.output"  # Каталог для результатов работы программы
        self.help = os.path.abspath(os.curdir+"/04.help/index.html")
        # надо вписать полный путь к исполняемому файлу inkscape
        #
        self.inkscape = self.txtInkscape.text()
        #
        ################################################################

    def set_window_on_center_screen(self): # Окно в ценрт экрана
        screen_geometry = QApplication.desktop().availableGeometry()
        screen_size = (screen_geometry.width(), screen_geometry.height())
        win_size = (self.frameSize().width(), self.frameSize().height())
        x = screen_size[0] // 2 - win_size[0] // 2
        y = screen_size[1] // 2 - win_size[1] // 2
        self.move(x, y)

    def set_NextBackEnabled(self): # Активируем и деактивируем кнопки Вперёд/назад
        if self.current_step == 0:
            self.btnBack.setEnabled(False)
        else:
            self.btnBack.setEnabled(True)

        if self.current_step == self.tabWidget.count() - 1:
            self.btnNext.setEnabled(False)
        else:
            self.btnNext.setEnabled(True)

    def onClickNext(self):  # Кнопка Вперёд

        if not os.path.exists(self.inkscape):
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Ошибка")
            msg_text = 'Не установлен Inkscape\n'
            msg_text += 'Приложение использует для преобразования SVG файлов в PDF формат внешний редактор - Inkscape\n'
            msg_text += '(https://inkscape.org)\n'
            msg_text += '\n'
            msg.setInformativeText(msg_text)
            msg.setWindowTitle("Ошибка")
            msg.exec_()
            return


        self.current_step += 1  # Переходим к следуюему шагу мастера
        if self.current_step > self.tabWidget.count() - 1:
            self.current_step = self.tabWidget.count() - 1

        self.set_NextBackEnabled()
        # Активируем нужную вкладку и выводим в заголовок название шага
        self.tabWidget.setCurrentIndex(self.current_step)
        out = "<html><head/><body><h1><span>"
        out += self.tabWidget.tabToolTip(self.current_step)
        out += "</span></h1></body></html>"
        self.lblTitle.setText(out)
        self.tabWidget.setTabEnabled(self.current_step, True)
        # Если на Х шаге,
        if self.current_step == 1: # то грузим первый в списке SVG файл
            self.render_svg(self.txtFileName.text())
            self.get_svg_files_list(self.templates)
            self.lstSVGFiles.setCurrentRow(0)
        elif self.current_step == 2: # то грузим первый в списке XLS файл
            if not self.is_table_data_loaded:
                self.get_xls_files()
                self.load_table_data()

    def onClickBack(self):  # Кнопка Назад, тут почти тоже что и в кнопке Вперёд
        self.current_step -= 1
        if self.current_step < 0:
            self.current_step = 0

        self.set_NextBackEnabled()

        self.tabWidget.setCurrentIndex(self.current_step)
        out = "<html><head/><body><h1><span>"
        out += self.tabWidget.tabToolTip(self.current_step)
        out += "</span></h1></body></html>"
        self.lblTitle.setText(out)

    def onTabWidgetClick(self, new_index): # Клики по вкладкам
        self.current_step = new_index
        self.set_NextBackEnabled()

    def onSVGFileClicked(self, item):  # Клик по SVG файлу в списке
        self.txtFileName.setText(item.text())
        self.render_svg(self.txtFileName.text())

    def render_svg(self, filename):  # Делаем превьюшку для SVG и отображаем в окне
        filename = self.templates + "/" + filename
        if os.path.exists(filename):
            self.pixmap = QPixmap(filename)
            self.lblSVG.setPixmap(
                self.pixmap.scaled(self.lblSVG.frameGeometry().width(),
                                   self.lblSVG.frameGeometry().height(),
                                   True, True))
        else:
            self.lblSVG.setPixmap(self.noimage)

    def get_svg_files_list(self, path):  # Получаем список SVG файлов из каталога с шаблонами
        if path:
            self.lstSVGFiles.clear()
            files = sorted(os.listdir(path))
            files = list(i for i in files if i[-3:] == "svg")
            if files:
                self.lstSVGFiles.addItems(files)

    def onXLSFileClicked(self, item):  # Клик по XLS файлу в списке
        self.tableWidget.clear()
        self.load_table_data()

    def get_xls_files(self):  # ПОлучаем список всех табличных файлов
        self.lstXLSFiles.clear()
        files = sorted(os.listdir(self.tabledata))
        files = list(i for i in files if i[-3:] == "xls")
        if files:
            self.lstXLSFiles.addItems(files)
        else:
            print('Нет табличных файлов :-(')

    def load_table_data(self):  # Загружаем данные из табличного файла, по которому кликнули
        if not self.is_table_data_loaded:
            self.lstXLSFiles.setCurrentRow(0)
        filename = self.tabledata + "/" + self.lstXLSFiles.item(self.lstXLSFiles.currentRow()).text()
        book = xlrd.open_workbook(filename)
        sheet = book.sheets()[0]
        headers = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        self.tableWidget.setColumnCount(len(headers[0]))
        self.tableWidget.setRowCount(len(headers) - 1)
        self.tableWidget.setHorizontalHeaderLabels(headers[0])
        for i in range(1, len(headers)):
            for j in range(len(headers[0])):
                self.tableWidget.setItem(i - 1, j, QTableWidgetItem(str(headers[i][j])))
        self.is_table_data_loaded = True
        self.headers = headers[0]  # сохраняем заголовок таблицы с образцами для замены в SVG


    def load_svg(self, filename): # Грузим в память весь SVG файл
        filename = self.templates + "/" + filename
        f = open(filename, encoding="utf8")
        data = f.readlines()
        f.close()
        return data

    def get_row(self, row): # Делаем список из строки таблицы
        r = []
        for col in range(self.tableWidget.columnCount()):
            if self.tableWidget.item(row, col):
                # Костыль для удаления дробной части числа
                r.append(self.tableWidget.item(row, col).text().replace(".0", ""))
        return r

    def create_new_svg(self, original_svg_data, data_to_replace): # Делаем новый временный SVG файл
        buf = original_svg_data[:]
        for i in range(len(self.headers)):
            for j in range(len(buf)):
                try:
                    if self.headers[i] in buf[j]:
                        buf[j] = buf[j].replace(self.headers[i], data_to_replace[i])
                except:
                    print("Ошибка")
        f = open(self.tmp + "/tmp.svg", "w", encoding="utf8")
        f.writelines(buf)
        f.close()

    def create_pdf(self, row_data, row_number):
        # Делаем имя файла из пары столбцов в текущей строке (Имя, номер по порядку и серийсный номер)
        if row_data:
            filename = str(row_number) + row_data[0] + row_data[3] + ".pdf"
            # Если линукс, то пользуемся модулем cairosvg
            if isLinux:
                cairosvg.svg2pdf(url=self.tmp + "/tmp.svg", write_to=self.result + "/" + filename)
            else:
                # print("Это не Linux, нужно установить inkscape")
                command = self.inkscape+' '
                args = '"'+self.tmp + '/tmp.svg" '
                args += '--export-type="pdf" '
                args += '--export-filename='
                args += '"'+self.result + "/" + filename+'"'
                os.system(command+args)


    def make_all_sert(self, filename): # Делаем сертификаты
        original_svg_data = self.load_svg(self.txtFileName.text())  # Оригинальные текст SVG файла
        self.progressBar.setMaximum(self.tableWidget.rowCount() - 1)  # Выставляем максимум для прогресс-бара
        # Просматриваем построчно таблицу
        self.set_all_disabled()
        self.txtLog.append("Создаём сертификат:")
        for row in range(self.tableWidget.rowCount()):
            self.create_new_svg(original_svg_data, self.get_row(row))  # Делаем временный файл SVG
            self.create_pdf(self.get_row(row), row)  # Делаем из SVG уже готовую PDF-ку
            self.progressBar.setValue(row)  # Сдвигаем полосу прогресса
            self.db_add_new_data(self.get_row(row), self.headers)
            self.txtLog.append(" " * 4 + str(row) + "." + self.get_row(row)[0])
        self.txtLog.append("Работа завершена")
        self.set_all_enabled()

    def set_all_disabled(self): # Отключение кнопок после запуска генерации
        self.btnMake_all_sert.setEnabled(False)
        self.btnOpenResults.setEnabled(False)
        self.btnBack.setEnabled(False)

    def set_all_enabled(self): # Включение кнопок после завершения процесса генерации
        self.btnMake_all_sert.setEnabled(True)
        self.btnOpenResults.setEnabled(True)
        self.btnBack.setEnabled(True)

    def onOpenResultClick(self): # Открываем папку с результатом работы
        self.open_dir(self.result)

    def onOpenTemplatesFolderClick(self): # Открываем папку с шаблонами
        self.open_dir(self.templates)

    def onOpenXLSFolderClick(self): # Открываем папку с таблиными файлами
        self.open_dir(self.tabledata)

    def open_dir(self, path): # Открытие папок
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])

    def onHelptClick(self):
        url = self.help
        webbrowser.open(url, new=0, autoraise=True)

    def db_add_new_data(self, data, headers): # Добавляем инфу в базу данных о каждом генерируемом сертификате
        try:
            self.con = sqlite3.connect(self.db)
            self.cur = self.con.cursor()
            query = "INSERT INTO info(name,event,date,sn) VALUES "
            query += "('" + data[0] + "','"
            query += data[1] + "','"
            query += data[2] + "','"
            query += data[3]
            query += "')"
            result = self.cur.execute(query).fetchall()
            self.con.commit()
            self.con.close()
        except:
            print("Ошибка: не могу записать данные в базу")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = frmMain()
    ex.show()
    sys.excepthook = except_hook
    sys.exit(app.exec_())
