import calclibrary
import calclib
import os
import sys
import time as tm
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox
from PyQt5.uic import loadUi

class MainWindow(QMainWindow):
    def __init__(self, title, qcombo, configname):
        super(MainWindow, self).__init__()
        loadUi("gui.ui", self)
        self.setWindowTitle(title)
        self.configname = configname
        self.pushButton_input.clicked.connect(self.browsefiles)
        self.pushButton_output.clicked.connect(self.browsefolder)
        self.combo_version.addItems(qcombo)
        self.combo_version.setCurrentIndex(1)
        self.pushButton_salir.clicked.connect(sys.exit)
        self.pushButton_proc.clicked.connect(self.process)
        self.actionGuardar_config.triggered.connect(self.saveconfig)
        self.actionAcerca_de.triggered.connect(self.about)
        if os.path.exists(self.configname):
            self.get_config()

    def get_config(self):
            init_data = calclib.read_config(self.configname)
            self.lineEdit_in.setText(init_data['Principal']['Origen'])
            self.lineEdit_out.setText(init_data['Principal']['Destino'])

    def browsefiles(self):
        fname = QFileDialog.getOpenFileName(self, 'Open file', os.getcwd(), 'Excel files (*.xls *.xlsx)')
        self.lineEdit_in.setText(fname[0])

    def browsefolder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        self.lineEdit_out.setText(folder)

    def process(self):
        start_time = tm.time()
        if str(self.combo_version.currentText())=='Version 2022':
            status = calclibrary.process_file_v2022(self.lineEdit_in.text(), self.lineEdit_out.text())
        else:
            status = calclibrary.process_file(self.lineEdit_in.text(), self.lineEdit_out.text(), streamlit=False)
        if status == 0:
            stop_time = tm.time()
            self._status.setText('Archivo procesado correctamente en {} segundos'.format(round(stop_time - start_time, 3)))
        elif status == -1:
            self._status.setText('Error en el procesamiento del archivo')

    def saveconfig(self):
        values = {'-ORIG-': self.lineEdit_in.text(),
                  '-DEST-': self.lineEdit_out.text()
        }
        calclib.save_config(values, self.configname)
        self._status.setText('Configuración guardada')

    def about(self):
        dlg = QMessageBox(self)
        dlg.setWindowTitle("Acerca de...")
        dlg.setText(f"Programación: Ricardo Juárez Pérez\nVersión: 0.3\nPyQT: {QtCore.qVersion()}")
        button = dlg.exec()

        if button == QMessageBox.Ok:
            print("OK!")

APP_TITLE = 'Elaboración de saldos'
ININAME = 'config.ini'
COMBO_MENU = ['Version <= 2021', 'Version 2022']

app = QApplication([APP_TITLE])
mainwindow = MainWindow(APP_TITLE, COMBO_MENU, ININAME)
widget = QtWidgets.QStackedWidget()
widget.addWidget(mainwindow)
widget.setFixedWidth(800)
widget.setFixedHeight(300)
widget.show()
sys.exit(app.exec_())