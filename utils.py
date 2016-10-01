#!/usr/bin/env python
# -*- coding: utf-8 -*-

from pyexcel import get_sheet
import re


def read_spreadsheet(path):
    try:
        return get_sheet(file_name=path).to_array()
    except:
        return None


def degree_to_decimal(coord):
    s = re.search("(\d+)[ENSW](\d+)\'(\d+)\"", coord)
    return round(int(s.group(1)) +
                 (float(s.group(2)) / 60) +
                 (float(s.group(3)) / 3600), 10)


def convert_coordinates(data, indexes):
    output = []
    for line in data:
        line[indexes[0]] = degree_to_decimal(ine[indexes[0]])
        line[indexes[1]] = degree_to_decimal(ine[indexes[1]])
        output.append(line)
    return output
