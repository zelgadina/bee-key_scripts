# an ugly script for placing layouts into folders for ugly articles
# by default, the first csv file found is used

import argparse
import csv
import re
import openpyxl
import os
import shutil
import sys
from collections import OrderedDict
from os import walk

market = '' # user needs set Ozon or WB for running the script

def csv_from_excel(src):
    dst = 'wb.csv'
    wb = openpyxl.load_workbook(src)
    sh = wb.active
    with open(dst, 'w', newline="") as f:
        c = csv.writer(f)
        for r in sh.rows:
            c.writerow([cell.value for cell in r])
    return dst

def find_csv(root_dir, market):
    for address, dirs, files in os.walk(root_dir):
        for name in files:
            match market:
                case 'ozon':
                    if CSV_FILE.match(name):
                        print('File used:', name)
                        return name
                case 'wb':
                    if XLS_FILE.match(name):
                        print('File used:', name)
                        print('Converting to csv...')
                        csv_name = csv_from_excel(name)
                        return csv_name
                case _:
                    print('Not found market', market)
                    return 'Unknown market'


ROOT_DIR = os.path.abspath(os.curdir)
CSV_FILE = re.compile('.*\.csv$')   # for Ozon
XLS_FILE = re.compile('.*\.xls.?$') # for WB

ozon_re_matching_items = OrderedDict([
    (re.compile('^[A-Z]*_?F.*'),      'Футболка'),
    (re.compile('^([A-Z]*_K|K)\w+$'), 'Кружка'),
    (re.compile('^P.*'),              'Коврик'),
    ])

ozon_re_matching_shirt_type = OrderedDict([
    (re.compile('.*((?<!H_)W-|W[MLS]|S_W).*'), 'Сублимация'),
    (re.compile('.*(WX|H_W).*'),               'Белый хлопок'),
    (re.compile('.*(H_B|BKX).*'),              'Чёрный хлопок'),
    ])


ozon_re_matching_shirt_size = OrderedDict([
    (re.compile('.+(XXXL|3XL)$'),          'XXXL'),
    (re.compile('.+(WXXL|[^WX]XXL|2XL)$'), '2XL'),
    (re.compile('.+(WXL|[^23WX]XL)$'),     'XL'),
    (re.compile('.+(WL|[^23XW]L)$'),       'L'),
    (re.compile('.+M$'),                   'M'),
    (re.compile('.+(2XS|14\s+лет\s*)$'),   '14 лет'),
    (re.compile('.+(3XS|12\s+лет\s*)$'),   '12 лет'),
    (re.compile('.+[^23]XS$'),             'XS'),
    (re.compile('.+[^X]S$'),               'S'),
    (re.compile('.+(38|10\s+лет\s*)$'),    '10 лет'),
    (re.compile('.+(36|8\s+лет\s*)$'),     '8 лет'),
    (re.compile('.+(34|6\s+лет\s*)$'),     '6 лет'),
    (re.compile('.+32$'),                  '32'),
    (re.compile('.+30$'),                  '30')
    ])


ozon_re_matching_pad_size = {
    re.compile('.?P.*L$'): 'Большой',
    re.compile('.?P.*M$'): 'Маленький',
}

# research an order list, get and recognize all items: article, type, size
def ozon_get_items(filename):
    items = {}
    with open(filename, newline='') as f:
        ozon = csv.DictReader(f, delimiter=';')
        for row in ozon:
            article = row['Артикул']
            items[article] = {'count': int(row['Количество']), 'type': ozon_recognize_item(article)}
    return items

def wb_get_items(filename):
    items = {}
    with open(filename, newline='') as f:
        wb = csv.DictReader(f, delimiter=',')
        for row in wb:
            article = row['Артикул продавца']
            size = row['Размер']
            items[article] = {'type': ozon_recognize_item(article), 'size': size}
    return items

# magic article matching
def ozon_recognize_item(article):
    item = []
    for key, value in ozon_re_matching_items.items():
        if key.match(article):
            item.append(value)
            if value == 'Футболка':
                for key, value in ozon_re_matching_shirt_type.items():
                    if key.match(article):
                        item.append(value)
                        for key, value in ozon_re_matching_shirt_size.items():
                            if key.match(article):
                                item.append(value)
                return item
            elif value == 'Коврик':
                for key, value in ozon_re_matching_pad_size.items():
                    if key.match(article):
                        item.append(value)
                return item
            elif value == 'Кружка':
                return item
            else:
                print('Not recognized:', value)
    return item

# find all files in layouts folders
def get_all_layouts(search_path):
    result = {}
    for address, dirs, files in os.walk(search_path):
        for name in files:
            if CSV_FILE.match(name):
                continue # if the file .csv in the current directory
            result[name] = os.path.join(address, name)
    return result

# try to match and find a file path for an article; copy the file to an order folder if success
def ozon_find_layouts(layouts, items, orders):
    errors = items.copy()
    for article, properties in items.items():
        for j in layouts.keys():
            for i in range(properties['count']):
                if article in j:
                    dst_dir = os.path.join(orders, *properties['type'], '')
                    os.makedirs(os.path.dirname(dst_dir), exist_ok=True)
                    # i'm so sorry...
                    count = ''
                    if i > 0:
                        count = str(i) + '_'
                    else:
                        del errors[article]
                    shutil.copyfile(layouts[j], os.path.join(dst_dir, count+j))
    return errors

def main():
    parser = argparse.ArgumentParser(description='a script for placing layouts into folders for articles')
    parser.add_argument('-m', '--market', help='ozon|wb', required=True)
    args = parser.parse_args()
    parser.add_argument('-e', '--excel', help='excel (csv or xls) file with orders', 
                        default=find_csv(ROOT_DIR, args.market))
    parser.add_argument('-l', '--layouts', help='layouts dir', 
                        default=os.path.join(ROOT_DIR, args.market))
    parser.add_argument('-o', '--output', help='output dir for sorted orders', 
                        default=os.path.join(ROOT_DIR, args.market+'_orders'))
    args = parser.parse_args()

    if args.market == 'ozon':
        ozon_list = ozon_get_items(args.excel)
        ozon_layouts_tree = get_all_layouts(args.layouts)
        print('Not found in layouts:\n')
        print(ozon_find_layouts(ozon_layouts_tree, ozon_list, args.output))
    elif args.market == 'wb':
        wb_list = wb_get_items(args.excel)
        wb_layouts_tree = get_all_layouts(args.layouts)
        print('Not found in layouts:\n')
        print(ozon_find_layouts(wb_layouts_tree, wb_list, args.output))

if __name__ == "__main__":
   main()
