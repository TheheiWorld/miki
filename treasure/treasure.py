#!/user/bin/env python3
# coding=utf-8

from openpyxl.reader.excel import load_workbook

treasure = {}
poi = {}
treasure_match = {}


def load_treasure(filename):
    """
    加载藏宝图
    :param filename:
    :return:
    """
    wb = load_workbook('/Users/juststand/Desktop/' + filename)
    sheets = wb.sheetnames
    sheet = wb[sheets[0]]

    for index, row in enumerate(sheet.__iter__()):
        if index == 0:
            continue
        key = row[4].value
        value = row[3].value
        treasure[key] = value


def load_poi(filename):
    """
    加载poi
    :param filename:
    :return:
    """
    wb = load_workbook('/Users/juststand/Desktop/' + filename)
    sheets = wb.sheetnames
    sheet = wb[sheets[0]]
    for index, row in enumerate(sheet.__iter__()):
        if index == 0:
            continue
        key = row[11].value
        value = row[8].value
        poi[key] = value


def match():
    for key in treasure:
        key_list = str(key).split("+")

        for poi_key in poi:
            match_flag = False
            for treasure_key in key_list:
                if str(treasure_key) in str(poi_key):
                    match_flag = True
                else:
                    match_flag = False
                    break
            if match_flag:
                try:
                    value = treasure_match[key]
                    value.append(poi[poi_key])
                    treasure_match[key] = value
                except KeyError:
                    treasure_match[key] = [poi[poi_key]]
            else:
                pass


if __name__ == '__main__':
    load_treasure("treasure.xlsx")
    load_poi("product.xlsx")
    match()
    print(treasure_match)
    treasure.clear()
    poi.clear()
    treasure_match.clear()
