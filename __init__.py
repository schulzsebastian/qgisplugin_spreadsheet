# -*- coding: utf-8 -*-
"""
/***************************************************************************
 Spreadsheet
                                 A QGIS plugin
 Plugin for loading spreadhseet data
                             -------------------
        begin                : 2016-07-03
        copyright            : (C) 2016 by Sebastian Schulz
        email                : schulz.siwy@gmail.com
        git sha              : $Format:%H$
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""


# noinspection PyPep8Naming
def classFactory(iface):  # pylint: disable=invalid-name
    """Load Spreadsheet class from file Spreadsheet.

    :param iface: A QGIS interface instance.
    :type iface: QgsInterface
    """
    #
    from .spreadsheet import Spreadsheet
    return Spreadsheet(iface)
