# -*- coding: utf-8 -*-
"""
/***************************************************************************
 Spreadsheet
                                 A QGIS plugin
 Plugin for loading spreadsheet data
                              -------------------
        begin                : 2016-07-03
        git sha              : $Format:%H$
        copyright            : (C) 2016 by Sebastian Schulz
        email                : schulz.siwy@gmail.com
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""
from PyQt4 import uic
from PyQt4.QtCore import QSettings
from PyQt4.QtGui import QFileDialog, QDialog
from pyexcel import get_sheet
import re
import os

FORM_CLASS, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'spreadsheet_module.ui'))


class SpreadsheetModule(QDialog, FORM_CLASS):

    def __init__(self, parent, parents=None):
        super(SpreadsheetModule, self).__init__(parents)
        self.setupUi(self)
        self.parent = parent
        self.iface = parent.iface
        self.fileButton.clicked.connect(self.selectFile)
        self.memoryButton.clicked.connect(self.outputOn)
        self.fileSaveButton.clicked.connect(self.outputOff)

    def read_spreadsheet(self, path):
        try:
            return get_sheet(file_name=path).to_array()
        except:
            return None

    def degree_to_decimal(self, coord):
        s = re.search("(\d+)[ENSW](\d+)\'(\d+)\"", coord)
        return round(int(s.group(1)) +
                     (float(s.group(2)) / 60) +
                     (float(s.group(3)) / 3600), 10)

    def convert_coordinates(self, data, indexes):
        output = []
        for line in data:
            line[indexes[0]] = degree_to_decimal(ine[indexes[0]])
            line[indexes[1]] = degree_to_decimal(ine[indexes[1]])
            output.append(line)
        return output

    def listEPSG(self):
        codes = []
        for code in QSettings().value('UI/recentProjectionsAuthId'):
            codes.append(code[5:])
        self.epsgBox.addItems(codes)

    def updateCoordinates(self):
        data = self.read_spreadsheet(self.fileLine.text())
        self.xBox.addItems(data[0])
        self.yBox.addItems(data[0])
        self.listEPSG()

    def selectFile(self):
        filename = QFileDialog.getOpenFileName(
            None,
            'Open spreadsheet file', '',
            'Spreadsheet file (*.xlsx *.xls *.ods)')
        self.fileLine.setText(filename)
        if filename:
            self.updateCoordinates()

    def outputOn(self):
        self.outputBox.setEnabled(True)
        self.outputBox.addItems(['Shapefile', 'GeoJSON'])

    def outputOff(self):
        self.outputBox.setEnabled(False)
        self.outputBox.clear()
