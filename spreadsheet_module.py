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
from qgis.core import QgsMapLayerRegistry, QgsField, QgsFields, \
        QgsVectorLayer, QgsFeature, QGis, QgsVectorFileWriter, QgsGeometry, \
        QgsPoint
from qgis.gui import QgsMessageBar, QgsProjectionSelectionWidget
from dependencies import *
from xlrd import open_workbook, xldate_as_tuple, XL_CELL_DATE
from datetime import datetime
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
        data = []
        wb = open_workbook(path)
        sh = wb.sheet_by_index(0)
        for row in range(sh.nrows):
            line = []
            for col in range(sh.ncols):
                if sh.cell_type(row, col) == XL_CELL_DATE:
                    dt_tuple = xldate_as_tuple(sh.cell(row, col).value, wb.datemode)
                    date = datetime(dt_tuple[0], dt_tuple[1], dt_tuple[2])
                    date = date.strftime('%d-%m-%Y')
                    line.append(date)
                else:
                    line.append(sh.cell(row, col).value)
            data.append(line)
        return data

    def degree_to_decimal(self, coord):
        s = re.search("(\d+)[ENSW](\d+)\'(\d+)\"", coord)
        return round(int(s.group(1)) +
                     (float(s.group(2)) / 60) +
                     (float(s.group(3)) / 3600), 10)

    def convert_coordinates(self, data, indexes):
        output = []
        for line in data:
            line[indexes[0]] = self.degree_to_decimal(line[indexes[0]])
            line[indexes[1]] = self.degree_to_decimal(line[indexes[1]])
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
        print self.spreadsheetData[0]
        pr.addAttributes([QgsField(i, QVariant.String)
                          for i in self.spreadsheetData[0]])
        if self.skipBox.isChecked():
            self.spreadsheetData = self.spreadsheetData[1:]
        if self.convertBox.isChecked():
            self.spreadsheetData = self.convert_coordinates(
                self.spreadsheetData,
                [self.xBox.currentIndex(), self.yBox.currentIndex()])
        for row in self.spreadsheetData:
            feature = QgsFeature()
            feature.setGeometry(QgsGeometry.fromPoint(
                QgsPoint(
                    row[self.xBox.currentIndex()],
                    row[self.yBox.currentIndex()])))
            feature.setAttributes(row)
            pr.addFeatures([feature])
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
            fields = QgsFields()
            for field in self.spreadsheetData[0]:
                fields.append(QgsField(field, QVariant.String))
            vl = QgsVectorFileWriter(path, "utf-8",
                                     fields,
                                     QGis.WKBPoint,
                                     self.epsgBox.crs(),
                                     "ESRI Shapefile")
            del vl
            vl = self.iface.addVectorLayer(path + '.shp',
                                           path.split('/')[-1],
                                           "ogr")
            pr = vl.dataProvider()
            vl.startEditing()
            for row in self.convert_coordinates(self.spreadsheetData[1:],
                                                [self.xBox.currentIndex(),
                                                 self.yBox.currentIndex()]):
                feature = QgsFeature()
                feature.setGeometry(QgsGeometry.fromPoint(
                    QgsPoint(
                        row[self.xBox.currentIndex()],
                        row[self.yBox.currentIndex()])))
                feature.setAttributes(row)
                pr.addFeatures([feature])
            vl.commitChanges()
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
