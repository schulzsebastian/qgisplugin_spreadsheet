#!/usr/bin/env python
# -*- coding: utf-8 -*-

from pyexcel import get_sheet


def read_spreadsheet(path):
    try:
        return get_sheet(file_name=path).to_array()
    except:
        return None
