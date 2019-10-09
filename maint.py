
import sys
import os
from PyQt5 import QtCore, QtWidgets
import func
import ui_test3
import json


class ExampleApp(QtWidgets.QMainWindow, ui_test3.Ui_MainWindow):

    def load_last_settings(self):

        try:
            self.settings_file = os.path.dirname(
                __file__)+"/last_settings.json"
            with open(self.settings_file, 'r') as read_file:
                self.last_settings = json.load(read_file)
            self.load_con_settings(
                self.last_settings["last_settings_file_dir"])
        except FileNotFoundError:
            self.status_listWidget.addItem("Первый запуск")
        self.status_listWidget.scrollToBottom()

    def button_click_connect(self):

        try:
            self.connection = func.conprop(
                host=self.hostLine.text(), port=self.portLine.text(
                ), db_name=self.dbnLine.text(), user_name=self.unLine.text(), password=self.passLine.text())
            self.connection.connect_db()
            self.load_table_settings()
            self.choose_tableBox.clear()
            for i in range(len(self.connection.tables)):
                self.choose_tableBox.addItem(
                    str("%s" % self.connection.tables[i]))
            self.status_listWidget.addItem(self.connection.massage)
        except AttributeError:
            self.status_listWidget.addItem('Проверьте данные подключения')
        except Exception as e:
            self.status_listWidget.addItem(str(e))
        self.status_listWidget.scrollToBottom()

    def button_click_save_con_settings(self):

        filters = "JSON файлы (*.json)"
        text = "Назовите файл настроек"
        error = 'Отмена сохранения файла настроек'
        try:
            json_string = {"db_name": "%s" % (self.dbnLine.text()), "user_name": "%s" % (self.unLine.text()), "host": "%s" % (
                self.hostLine.text()), "port": "%s" % (self.portLine.text()), "password": "%s" % (self.passLine.text())}
            file_name = self.save_dialog(filters, text, error)
            file_name_f, file_extention = os.path.splitext(file_name)
            if file_extention == '.json':
                file_name_f = file_name
            else:
                file_name_f = file_name+".json"
            with open(file_name_f, "w") as write_file:
                json.dump(json_string, write_file)
            self.save_last_settings(file_name_f)
            self.status_listWidget.addItem("Успешно сохранен файл настроек")
        except(TypeError, UnboundLocalError):
            self.status_listWidget.addItem('Отмена сохранения файла настроек')
        except Exception as e:
            self.status_listWidget.addItem(e)
        self.status_listWidget.scrollToBottom()

    def save_last_settings(self, settings_dir):

        self.last_settings = {"last_settings_file_dir": "%s" % (settings_dir)}
        with open(self.settings_file, "w") as write_file:
            json.dump(self.last_settings, write_file)

    def load_dialog(self, filters, text, error):

        try:
            f_dir = os.path.dirname(__file__)
            file_name = QtWidgets.QFileDialog.getOpenFileName(
                None, text, f_dir, filters)
        except(TypeError, FileNotFoundError):
            self.status_listWidget.addItem(error)
        except AttributeError as e:
            self.status_listWidget.addItem(error)
        except Exception as e:
            self.status_listWidget.addItem(e)
        return file_name[0]

    def save_dialog(self, filters, text, error):

        try:
            f_dir = os.path.dirname(__file__)
            file_name = QtWidgets.QFileDialog.getSaveFileName(
                None, text, f_dir, filters)
        except Exception as e:
            self.status_listWidget.addItem(e)
            self.status_listWidget.addItem(error)
        return file_name[0]

    def load_con_settings(self, file_name):

        try:
            with open(file_name, 'r') as read_file:
                self.conn_data = json.load(read_file)
            self.hostLine.setText(self.conn_data["host"])
            self.portLine.setText(self.conn_data["port"])
            self.dbnLine.setText(self.conn_data["db_name"])
            self.unLine.setText(self.conn_data["user_name"])
            self.passLine.setText(self.conn_data["password"])
            self.status_listWidget.addItem('Успешно загружен файл настроек')

        except(TypeError, FileNotFoundError):
            self.status_listWidget.addItem('Отмена загрузки файла настроек')
        except IsADirectoryError:
            pass
        except Exception as e:
            self.status_listWidget.addItem(e)
        self.status_listWidget.scrollToBottom()

    def button_click_load_con_settings(self):

        filters = "JSON файлы (*.json)"
        text = "Выберете файл настроек"
        error = 'Отмена загрузки файла настроек'
        try:
            file_name = self.load_dialog(filters, text, error)
            self.load_con_settings(file_name)
            self.save_last_settings(file_name)
        except Exception as e:
            self.status_listWidget.addItem(e)
        self.status_listWidget.scrollToBottom()

    def load_table_settings(self):
        self.connection.load_table_settings()

    def button_click_open_excel_file(self):

        filters = "Книга Excel 97-2003 (*.xls);;Книга Excel 2010 (*.xlsx)"
        text = "Выберете Exel файл"
        error = "Открытие файла отменено"
        try:
            self.c_sheetBox.clear()
            file_name = self.load_dialog(filters, text, error)
            self.excel_file = func.excel_edit(file_name=file_name)
            self.excel_file.read_excel()
            print(self.excel_file.sheet_names)
            for i in range(len(self.excel_file.sheet_names)):
                self.c_sheetBox.addItem(self.excel_file.sheet_names[i])
            self.c_sheet_name = self.excel_file.sheet_names[0]
            self.excel_file.read_xls_sheet(self.c_sheet_name)
            self.set_columns_numbers()
            self.status_listWidget.addItem(self.excel_file.massage)
        except Exception as e:
            self.status_listWidget.addItem(str(e))
        self.status_listWidget.scrollToBottom()

    def button_click_delete_spaces(self):

        self.cds_pressed = False
        try:
            self.excel_file.delete_spaces()
            self.status_listWidget.addItem(self.excel_file.massage)
            self.status_listWidget.scrollToBottom()
            self.cds_pressed = True
        except Exception:
            self.status_listWidget.addItem(
                "Нет открытого файла")
            self.cds_pressed = False
        self.status_listWidget.scrollToBottom()

    def button_save_excel_file(self):

        filters = "Книга Excel 97-2003 (*.xls);;Книга Excel 2010 (*.xlsx)"
        text = "Назовите файл Excel"
        error = "Нет открытого файла для сохранения"
        file_name = self.save_dialog(filters, text, error)
        try:
            self.excel_file.save_excel(file_name)
            self.status_listWidget.addItem(self.excel_file.massage)
        except Exception as e:
            self.status_listWidget.addItem(e)
        self.status_listWidget.scrollToBottom()


    def load_data_to_server(self):

        try:
            if self.cds_pressed == True:
                self.connection.load_data_to_server(
                    self.excel_file.vals_1, self.excel_file.numbers)
            else:
                self.connection.load_data_to_server(
                    self.excel_file.vals, self.excel_file.numbers)
            self.status_listWidget.addItem(self.connection.massage)
        except Exception:
            self.status_listWidget.addItem('Нет открытого файла Excel')
        self.status_listWidget.scrollToBottom()

    def onActivated_ct(self, text):

        self.connection.loaded_table_name(text)
        self.status_listWidget.addItem(self.connection.massage)
        self.status_listWidget.scrollToBottom()

    def choose_columns(self):

        try:
            if self.diapasonButton.isChecked():
                self.excel_file.diapason(self.from_d, self.to)
            else:
                self.excel_file.pere(self.pereEdit.text())
        except AttributeError as e:
            self.status_listWidget.addItem("Не открыт Excel файл")
        except Exception as e:
            self.status_listWidget.addItem(str(e))
        self.status_listWidget.scrollToBottom()

    def set_columns_numbers(self):

        self.fromBox.clear()
        self.toBox.clear()
        for i in range(len(self.excel_file.vals[0])):
            self.fromBox.addItem(str(i+1))
            self.toBox.addItem(str(i+1))

    def onActivated_ccdf(self, text):

        self.from_d = int(text)

    def onActivated_ccdt(self, text):

        self.to = int(text)

    def onActivated_cs(self, text):

        self.c_sheet_name = text
        self.excel_file.read_sheet(self.c_sheet_name)

    def init_ui(self):

        self.choose_tableBox.activated[str].connect(self.onActivated_ct)
        self.passLine.setEchoMode(QtWidgets.QLineEdit.Password)
        self.connectButton.clicked.connect(self.button_click_connect)
        self.save_setButton.clicked.connect(
            self.button_click_save_con_settings)
        self.load_setButton.clicked.connect(
            self.button_click_load_con_settings)
        self.ofButton.clicked.connect(self.button_click_open_excel_file)
        self.del_spButton.clicked.connect(self.button_click_delete_spaces)
        self.save_excelButton.clicked.connect(self.button_save_excel_file)
        self.load_dataButton.clicked.connect(self.load_data_to_server)
        self.choose_columnsButton.clicked.connect(self.choose_columns)
        self.fromBox.activated[str].connect(self.onActivated_ccdf)
        self.toBox.activated[str].connect(self.onActivated_ccdt)
        self.c_sheetBox.activated[str].connect(self.onActivated_cs)

    def init_params(self):

        self.cds_pressed = False
        self.connection = func.conprop(0, 0, 0, 0, 0)
        self.from_d = 1
        self.to = 1
        self.diapasonButton.setChecked(True)

    def __init__(self):

        super().__init__()
        self.setupUi(self)
        self.init_ui()
        self.init_params()
        self.load_last_settings()

def main():

    app = QtWidgets.QApplication(sys.argv)
    window = ExampleApp()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()
