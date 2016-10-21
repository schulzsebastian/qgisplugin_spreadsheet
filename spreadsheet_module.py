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
from PyQt4.QtCore import QSettings, QVariant
from PyQt4.QtGui import QFileDialog, QDialog, QDialogButtonBox
from qgis.core import QgsMapLayerRegistry, QgsField, QgsVectorLayer
from qgis.gui import QgsMessageBar, QgsProjectionSelectionWidget
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
        self.memoryButton.clicked.connect(self.outputOff)
        self.fileSaveButton.clicked.connect(self.outputOn)
        self.executeBox.button(QDialogButtonBox.Ok).setEnabled(False)
        self.epsgBox.setOptionVisible(
            QgsProjectionSelectionWidget.CurrentCrs, False)

    def read_spreadsheet(self, path):
        return get_sheet(file_name=path).to_array()

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

    def updateCoordinates(self):
        self.spreadsheetData = self.read_spreadsheet(self.fileLine.text())
        self.xBox.clear()
        self.yBox.clear()
        try:
            self.xBox.addItems(self.spreadsheetData[0])
            self.yBox.addItems(self.spreadsheetData[0])
            self.executeBox.button(QDialogButtonBox.Ok).setEnabled(True)
        except IndexError:
            self.iface.messageBar().pushMessage(
                'Spreadsheet',
                'Empty file',
                level=QgsMessageBar.WARNING)

    def selectFile(self):
        filename = QFileDialog.getOpenFileName(
            None,
            'Open spreadsheet file',
            '',
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

    def createMemoryLayer(self):
        vl = QgsVectorLayer(
            "Point?crs=" + self.epsgBox.crs().geographicCRSAuthId(),
            "Spreadsheet",
            "memory")
        pr = vl.dataProvider()
        vl.startEditing()
        pr.addAttributes([QgsField(i, QVariant.String)
                          for i in self.spreadsheetData[0]])
        vl.commitChanges()
        QgsMapLayerRegistry.instance().addMapLayer(vl)
        return True

    def createShapefile(self):
        path = QFileDialog.getSaveFileName(
            None,
            'Save as Shapefile',
            '',
            'Select directory and set output filename')
        if path:
            print path
            return True
        return False

    def createGeoJSON(self):
        path = QFileDialog.getSaveFileName(
            None,
            'Save as GeoJSON',
            '',
            'Select directory and set output filename')
        if path:
            print path
            return True
        return False

    def accept(self):
        if self.run():
            super(SpreadsheetModule, self).accept()

    def run(self):
        if len(self.spreadsheetData) < 2:
            self.iface.messageBar().pushMessage(
                'Spreadsheet',
                'No data except the header',
                level=QgsMessageBar.WARNING)
            return False
        if self.memoryButton.isChecked():
            return self.createMemoryLayer()
        elif self.outputBox.currentText() == 'Shapefile':
            return self.createShapefile()
        elif self.outputBox.currentText() == 'GeoJSON':
            return self.createGeoJSON()
        return False
