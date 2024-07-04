# an ugly script for placing layouts into folders for ugly articles
# by default, the first csv file found is used

import csv
import re
import os
import shutil
from os import walk


def find_csv(root_dir):
    for address, dirs, files in os.walk(root_dir):
        for name in files:
            if EXCEL.match(name):
                print('File used:', name)
                return name

ROOT_DIR = os.path.abspath(os.curdir)
EXCEL = re.compile('.*\.csv$')
ozon_items_list = os.path.join(find_csv(ROOT_DIR))
ozon_layouts_dir = os.path.join(ROOT_DIR, 'OZON')
ozon_orders_dir = os.path.join(ROOT_DIR, 'Заказы')

ozon_re_matching_items = {
    re.compile('^([A-Z]*_K|K)\w+\d+$'): 'Кружка',
    re.compile('^P.*'):                 'Коврик',
    re.compile('^[A-Z]*_?F.*'):         'Футболка',
}

ozon_re_matching_shirt_type = {
    re.compile('.*((?<!H_)W-|W[MLS]|S_W).*'): 'Сублимация',
    re.compile('.*(WX|H_W).*'):               'Белый хлопок',
    re.compile('.*(H_B|BKX).*'):              'Чёрный хлопок',
}

ozon_re_matching_shirt_size = {
    re.compile('.+(XXXL|3XL)$'):          'XXXL',
    re.compile('.+(WXXL|[^WX]XXL|2XL)$'): '2XL',
    re.compile('.+(WXL|[^23WX]XL)$'):     'XL',
    re.compile('.+(WL|[^23XW]L)$'):       'L',
    re.compile('.+M$'):                   'M',
    re.compile('.+(2XS|14\s+лет\s*)$'):   '14 лет',
    re.compile('.+(3XS|12\s+лет\s*)$'):   '12 лет',
    re.compile('.+[^23]XS$'):             'XS',
    re.compile('.+[^X]S$'):               'S',
    re.compile('.+(38|10\s+лет\s*)$'):    '10 лет',
    re.compile('.+(36|8\s+лет\s*)$'):     '8 лет',
    re.compile('.+(34|6\s+лет\s*)$'):     '6 лет',
    re.compile('.+32$'):                  '32',
    re.compile('.+30$'):                  '30',
}

ozon_re_matching_pad_size = {
    re.compile('.?P.*L$'): 'Большой',
    re.compile('.?P.*M$'): 'Маленький',
}

# research an order list, get and recognize all items: article, type, size
def ozon_get_items(dirname):
    items = {}
    with open(dirname, newline='') as f:
        ozon = csv.DictReader(f, delimiter=';')
        for row in ozon:
            article = row['Артикул']
            items[article] = {'count': int(row['Количество']), 'type': ozon_recognize_item(article)}
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
        if value == 'Коврик':
            for key, value in ozon_re_matching_pad_size.items():
                if key.match(article):
                    item.append(value)
    return item

# find all files in layouts folders
def get_all_layouts(search_path):
    result = {}
    for address, dirs, files in os.walk(search_path):
        for name in files:
            if EXCEL.match(name):
                continue # if the file .csv in the current directory
            result[name] = os.path.join(address, name)
    return result

# try to match and find a file path for an article; copy the file to an order folder if success
def ozon_find_layouts(layouts, items):
    errors = items.copy()
    for article, properties in items.items():
        for j in layouts.keys():
            for i in range(properties['count']):
                if article in j:
                    dst_dir = os.path.join(ozon_orders_dir, *properties['type'], '')
                    os.makedirs(os.path.dirname(dst_dir), exist_ok=True)
                    # i'm so sorry...
                    count = ''
                    if i > 0:
                        count = str(i) + '_'
                    else:
                        del errors[article]
                    shutil.copyfile(layouts[j], os.path.join(dst_dir, count+j))
    return errors

def change_dir():
    ozon_items_list = input('Файл с заказами Озона:\n')
    ozon_layouts_dir = input('Директория с макетами Озона:\n')
    ozon_orders_dir = input('Куда положить:\n')

def main():
    change = input('Изменить пути к файлам? (y/N)\n')
    if change.lower() == ('y' or 'yes'):
        change_dir()
    ozon_list = ozon_get_items(ozon_items_list)
    # print(ozon_list, "\n\n")

    ozon_layouts_tree = get_all_layouts(ozon_layouts_dir)
    # print(ozon_layouts_tree)
    print('Not found in layouts:\n')
    print(ozon_find_layouts(ozon_layouts_tree, ozon_list))

main()
