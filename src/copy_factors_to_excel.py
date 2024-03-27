#!interpreter [optional-arg]
# -*- coding: utf-8 -*-

"""
{Description}
{License_info}
"""
__filename__ = 'copy_factors_to_excel.py'
__author__ = 'Andre Wiegleb'
__created__ = '26.03.2024'
__copyright__ = 'Copyright 2024, knc'
__credits__ = ['{credit_list}']
__license__ = '{license}'
__version__ = '{mayor}.{minor}.{rel}'
__maintainer__ = 'Andre Wiegleb'
__email__ = 'andre.wiegleb@mts.com'
__status__ = '{dev_status}'

import xlwings as xw

from icecream import ic
from tkinter import filedialog

# Öffnen eines Dateiauswahldialogs
# files = filedialog.askopenfilenames()


# print(file_path)
excel_file = 'C:/Users/wiegleba/PycharmProjects/knc/data/summary.xls'

# Öffnen der Excel-Datei
app = xw.App(visible=True)
wb = app.books.open(excel_file)


def get_data(file_path):
    """
    Liest die Daten aus einer Textdatei und gibt sie als Liste zurück.
    Ermittelt die Ecke aus dem Dateinamen. Eckennamen sind die ersten beiden Buchstaben des Dateinamens und werden
    für die Zuordnung der Daten in das Excel-Arbeitsblatt verwendet.

    :param file_path:
    :return: (corner, data), corner: Ecke(z.B. 'Rf'), data: Liste mit den Daten

    """
    corner = file_path.split('/')[-1][0:2]
    ic(corner)
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

def write_data_to_excel(files, cell_range):
    """
    Schreibt die Daten aus den Textdateien in das Excel-Arbeitsblatt.
    :param files: Liste von Dateipfaden
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


# Ausführung der Funktion
# write_data_to_excel(files, 'A53')
