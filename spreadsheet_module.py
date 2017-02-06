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
from pyexcel_ods import get_data
from datetime import datetime, date as dt
import json
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
        self.regexLine.setText("(\d+)[ENSW](\d+)\'(\d+)\"")
        self.fileButton.clicked.connect(self.selectFile)
        self.memoryButton.clicked.connect(self.toggleOutput)
        self.fileSaveButton.clicked.connect(self.toggleOutput)
        self.convertBox.clicked.connect(self.toggleConvert)
        self.executeBox.button(QDialogButtonBox.Ok).setEnabled(False)
        self.epsgBox.setOptionVisible(
            QgsProjectionSelectionWidget.CurrentCrs, False)

    def toggleConvert(self):
        if self.convertBox.isChecked():
            self.regexLine.setEnabled(True)
            self.regexLabel.setEnabled(True)
        else:
            self.regexLine.setEnabled(False)
            self.regexLabel.setEnabled(False)

    def toggleOutput(self):
        if self.fileSaveButton.isChecked():
            self.outputLabel.setEnabled(True)
            self.outputBox.setEnabled(True)
            self.outputBox.addItems(['Shapefile'])
        else:
            self.outputLabel.setEnabled(False)
            self.outputBox.setEnabled(False)
            self.outputBox.clear()

    def read_spreadsheet(self, path):
        data = []
        if path.split('.')[-1] == 'ods':
            wb = get_data(path)
            sh = wb[wb.keys()[0]]
            for row in sh:
                fixed_line = [i.strftime('%d-%m-%Y') if isinstance(i, dt)
                              else i for i in row]
                data.append(fixed_line)
        elif path.split('.')[-1] in ['xlsx', 'xls']:
            wb = open_workbook(path)
            sh = wb.sheet_by_index(0)
            for row in range(sh.nrows):
                line = []
                for col in range(sh.ncols):
                    if sh.cell_type(row, col) == XL_CELL_DATE:
                        dt_tuple = xldate_as_tuple(
                            sh.cell(row, col).value, wb.datemode)
                        date = datetime(dt_tuple[0], dt_tuple[1], dt_tuple[2])
                        date = date.strftime('%d-%m-%Y')
                        line.append(date)
                    else:
                        line.append(sh.cell(row, col).value)
                data.append(line)
        return data

    def degree_to_decimal(self, coord):
        s = re.search(self.regexLine.text(), coord)
        sign = 1
        if [char for char in ['s', 'w'] if char in coord.lower()]:
            sign = -1
        return round(int(s.group(1)) +
                     (float(s.group(2)) / 60) +
                     (float(s.group(3)) / 3600), 10) * sign

    def convert_coordinates(self, data, indexes):
        output = []
        try:
            self.degree_to_decimal(data[0][indexes[0]])
            self.degree_to_decimal(data[0][indexes[1]])
        except:
            return output
        for line in data:
            line[indexes[0]] = self.degree_to_decimal(line[indexes[0]])
            line[indexes[1]] = self.degree_to_decimal(line[indexes[1]])
            output.append(line)
        return output

    def validate_coordinates(self):
        if self.skipBox.isChecked():
            self.spreadsheetData = self.spreadsheetData[1:]
        if self.convertBox.isChecked():
            self.spreadsheetData = self.convert_coordinates(
                self.spreadsheetData,
                [self.xBox.currentIndex(), self.yBox.currentIndex()])
        if not self.spreadsheetData:
            self.iface.messageBar().pushMessage(
                'Spreadsheet',
                'Invalid headers',
                level=QgsMessageBar.WARNING)
            return False
        try:
            float(self.spreadsheetData[0][self.xBox.currentIndex()])
            float(self.spreadsheetData[0][self.yBox.currentIndex()])
            return True
        except:
            self.iface.messageBar().pushMessage(
                'Spreadsheet',
                'Invalid coordinates',
                level=QgsMessageBar.WARNING)
            return False

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

    def createLayer(self, save=None):
        vl = QgsVectorLayer(
            "Point?crs=" + self.epsgBox.crs().geographicCRSAuthId(),
            "Spreadsheet",
            "memory")
        pr = vl.dataProvider()
        vl.startEditing()
        pr.addAttributes([QgsField(i, QVariant.String)
                          for i in self.spreadsheetData[0]])
        if self.validate_coordinates():
            for row in self.spreadsheetData:
                feature = QgsFeature()
                feature.setGeometry(QgsGeometry.fromPoint(
                    QgsPoint(
                        row[self.xBox.currentIndex()],
                        row[self.yBox.currentIndex()])))
                feature.setAttributes(row)
                pr.addFeatures([feature])
            vl.commitChanges()
            if not save:
                QgsMapLayerRegistry.instance().addMapLayer(vl)
                return True
            elif save == 'Shapefile':
                path = QFileDialog.getSaveFileName(
                    None,
                    'Save as Shapefile',
                    '',
                    'Select directory and set output filename')
                if path:
                    QgsVectorFileWriter.writeAsVectorFormat(
                        vl,
                        path + '.shp',
                        'utf-8',
                        vl.crs(),
                        'ESRI Shapefile')
                    self.iface.addVectorLayer(path + '.shp',
                                              os.path.basename(path),
                                              'ogr')
                    return True
        return False

    def accept(self):
        if self.run():
            super(SpreadsheetModule, self).accept()

    def run(self):
        # update spreadsheet data (twice-used fix)
        self.spreadsheetData = self.read_spreadsheet(self.fileLine.text())
        if len(self.spreadsheetData) < 2:
            self.iface.messageBar().pushMessage(
                'Spreadsheet',
                'No data except the header',
                level=QgsMessageBar.WARNING)
            return False
        if self.memoryButton.isChecked():
            return self.createLayer()
        else:
            return self.createLayer(save=self.outputBox.currentText())
