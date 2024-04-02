#!interpreter [optional-arg]
# -*- coding: utf-8 -*-

"""
{Description}
{License_info}
"""
__filename__ = 'coefficient_gui.py'
__author__ = 'Andre Wiegleb'
__created__ = '26.03.2024'
__copyright__ = 'Copyright 2024, knc'
__version__ = '0.1.0'
__maintainer__ = 'Andre Wiegleb'
__status__ = '{dev_status}'

import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, \
    QPushButton, QLineEdit, QFileDialog, QLabel

import xlwings as xw


def get_data(file_path):
    """
    Liest die Daten aus einer Textdatei und gibt sie als Liste zurück.
    Ermittelt die Ecke aus dem Dateinamen. Eckennamen sind die ersten beiden Buchstaben des Dateinamens und werden
    für die Zuordnung der Daten in das Excel-Arbeitsblatt verwendet.

    :param file_path:
    :return: (corner, data), corner: Ecke(z.B. 'Rf'), data: Liste mit den Daten

    """
    corner = file_path.split('/')[-1][0:2]

    # Öffnen der Textdatei im Lesemodus
    with open(file_path, 'r') as file:
        # Lesen aller Zeilen aus der Datei
        lines = [line.strip().split()[-1] if line.strip() else '' for line in file.readlines()]

        data = [
            lines[6:13],
            lines[18:25],
            lines[30:37],
            lines[42:49],
            lines[54:61],
            lines[66:73]
        ]

    return corner, data


def write_data_to_excel(files, wb, cell_range):
    """
    Schreibt die Daten aus den Textdateien in das Excel-Arbeitsblatt.
    :param files: Liste von Dateipfaden
    :param wb: Excel-Arbeitsmappe
    :param cell_range: erste Zelle, in die die Daten geschrieben werden
    :return: None
    """

    for file in files:
        corner, data = get_data(file)
        sheet = wb.sheets[corner]
        sheet.range(cell_range).value = data[0]
        sheet.range(cell_range).offset(row_offset=1).value = data[1]
        sheet.range(cell_range).offset(row_offset=2).value = data[2]
        sheet.range(cell_range).offset(row_offset=3).value = data[3]
        sheet.range(cell_range).offset(row_offset=4).value = data[4]
        sheet.range(cell_range).offset(row_offset=5).value = data[5]


class MainWindow(QWidget):
    """Main Window der Anwendung."""
    def __init__(self):
        super().__init__()

        self.excel_file = None
        self.data_files = None
        self.lineEdit = None
        self.init_ui()

    def init_ui(self):
        """Initialisiert die Benutzeroberfläche."""
        # Layout erstellen
        main_layout = QVBoxLayout()

        # Dateidialog-Button 1 erstellen
        button1 = QPushButton('Open *FmCC File(s)', self)
        button1.clicked.connect(self.open_data_files)
        main_layout.addWidget(button1)

        # Dateidialog-Button 2 erstellen
        button2 = QPushButton('Open Excel Summary', self)
        button2.clicked.connect(self.open_excel_file)
        main_layout.addWidget(button2)

        # Layout für Label, LineEdit und Button erstellen
        layout = QHBoxLayout()

        # Label erstellen
        label = QLabel('Excels First Cell:', self)
        layout.addWidget(label)

        # LineEdit erstellen
        self.lineEdit = QLineEdit(self)
        layout.addWidget(self.lineEdit)

        # Dateidialog-Button erstellen
        button = QPushButton('Execute', self)
        button.clicked.connect(self.execute)
        layout.addWidget(button)

        main_layout.addLayout(layout)

        # Hauptfenster konfigurieren
        self.setLayout(main_layout)
        self.setWindowTitle('KnC Coefficient Tool')
        self.show()

    def open_data_files(self):
        """Öffnet einen Dateidialog, um die Daten-Dateien auszuwählen."""
        options = QFileDialog.Options()
        file_dialog = QFileDialog()
        self.data_files, _ = file_dialog.getOpenFileNames(self, 'DSelect Data Files', '', 'Alle Dateien (*)',
                                                          options=options)

    def open_excel_file(self):
        """Öffnet einen Dateidialog, um die Excel-Datei auszuwählen."""
        options = QFileDialog.Options()
        file_dialog = QFileDialog()
        self.excel_file, _ = file_dialog.getOpenFileName(self, 'Select Excel File', '', 'Alle Dateien (*)',
                                                         options=options)

    def execute(self):
        """Überträgt die Daten aus den Textdateien in die Excel-Datei."""
        try:
            excel_app = xw.App(visible=True)
            wb = excel_app.books.open(self.excel_file)
            write_data_to_excel(self.data_files, wb, self.lineEdit.text())
        except Exception as e:
            print(e)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec_())
