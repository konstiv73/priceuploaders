# -*- coding: utf-8 -*-
from xlrd import open_workbook, XL_CELL_EMPTY, XL_CELL_TEXT


class ForwardUploader(object):
    # row, column это индекс, начинаются с 0
    IMAGE_URL = u'http://www.autobody.ru/upload/images/{code}.jpg'
    name = u'Форвард-Москва'
    START_ROW_INDEX = 5
    START_MANUFACTURER_COL = 14
    MANUFACTURER_ROW = 2

    sheets = [
        u'ГРУЗОВИКИ',
        u'КИТАЙ',
        u'АМЕРИКА',
        u'КОРЕЯ',
        u'ЯПОНИЯ',
        u'ЕВРОПА',
    ]

    COUNT_COLS = [6, 8, 9]

    config = {
        'map': {
            'sheet': None,
            'category_0': None,
            'category': None,
            'manufacturer': None,
            'manufacturer_code': None,
            'supplier_name': None,
            'is_available': None,
            'original_num': 2,
            'name': 4,
            'code': 1,
            'year': 3,
            'count': None,
            'waiting_pechatniki': 10,
            'price': 5, }, }

    def __init__(self, file_path):
        self.index = 0
        self.xls_data = self.get_data(file_path)

    def __iter__(self):
        return self

    def __next__(self):
        return self.next()

    def next(self):
        try:
            value = self.xls_data[self.index]
            self.index += 1
        except IndexError:
            raise StopIteration()
        values = dict()
        for key in self.config['map'].keys():
            if value:
                try:
                    values[key] = value[key]
                except KeyError:
                    values[key] = ''
        values['image_url'] = self.IMAGE_URL.format(code=values['code'])
        values['supplier_name'] = self.name
        return values

    def get_data(self, path):
        book = open_workbook(path, on_demand=True)
        data = list()
        for sheet_name in book.sheet_names():
            if sheet_name in self.sheets:
                category = ''
                category_0 = ''
                category_row = None
                category_value = None
                sheet = book.sheet_by_name(sheet_name)
                manufacturers = {}
                for col in range(self.START_MANUFACTURER_COL, sheet.ncols):
                    cell = sheet.cell(self.MANUFACTURER_ROW, col)
                    if cell.ctype != XL_CELL_EMPTY:
                        manufacturers[col] = cell.value
                for row in range(self.START_ROW_INDEX, sheet.nrows):
                    values = {}
                    for col_name, col in self.config['map'].items():
                        if col:
                            cell = sheet.cell(row, col)
                            if cell.ctype != XL_CELL_EMPTY:
                                if cell.ctype == XL_CELL_TEXT:
                                    values[col_name] = cell.value.strip()
                                elif col_name == 'code':
                                    try:
                                        values[col_name] = unicode(
                                            int(cell.value))
                                    except ValueError:
                                        values[col_name] = cell.value
                                else:
                                    values[col_name] = cell.value

                    if len(values) > 1:
                        values['category'] = category
                        values['category_0'] = category_0

                        for col in manufacturers.keys():
                            cell = sheet.cell(row, col)
                            if cell.ctype != XL_CELL_EMPTY:
                                values['manufacturer_code'] = cell.value
                                values['manufacturer'] = manufacturers[col]

                        values['count'] = 0
                        for col in self.COUNT_COLS:
                            cell = sheet.cell(row, col)
                            if cell.ctype != XL_CELL_EMPTY:
                                if cell.value == '+':
                                    values['count'] += 1
                                else:
                                    try:
                                        values['count'] += int(cell.value)
                                    except (TypeError, ValueError):
                                        pass
                        values['is_available'] = int(values['count']) > 0
                        values['sheet'] = sheet.name
                        data.append(values)
                    elif len(values) == 1:
                        category = values['code']
                        if category_row == row - 1:
                            # Предыдущая строка также категория,
                            # значит это родительская категория
                            category_0 = category_value or values['code']
                        category_row = row
                        category_value = values['code']
                    else:
                        continue
                book.unload_sheet(sheet_name)
        return data


# class ForwardSpbUploader(ForwardUploader):
#     name = u'Форвард-Питер'
#     START_ROW_INDEX = 1
#     START_MANUFACTURER_COL = 8
#     MANUFACTURER_ROW = 0

#     COUNT_COLS = [6, ]

#     sheets = [
#         u'Лист1',
#     ]

#     config = {
#         'map': {
#             'sheet': None,
#             'category_0': None,
#             'category': None,
#             'manufacturer': 7,
#             'manufacturer_code': None,
#             'supplier_name': None,
#             'is_available': None,
#             'original_num': 4,
#             'name': 2,
#             'code': 1,
#             'year': None,
#             'count': None,
#             'waiting_pechatniki': None,
#             'price': 5, }, }
