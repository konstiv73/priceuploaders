# -*- coding: utf-8 -*-
from xlrd import open_workbook, XL_CELL_EMPTY, XL_CELL_TEXT
from xlrd.sheet import ctype_text


class DeviceUploader(object):
    name = u'Девайс'
    START_ROW_INDEX = 11
    CATEGORY_PARENT_ROW_COLOR = 22
    CATEGORY_ROW_COLOR = 31
    MANUFACTURER_ROW_COLOR = 27
    PRODUCT_ROW_COLOR = 9

    config = {
        'map': {
            'sheet': None,
            'category': None,
            'manufacturer': None,
            'manufacturer_code': None,
            'supplier_name': None,
            'is_available': None,
            'original_num': None,
            'name': 1,
            'code': 3,
            'year': None,
            'count': 2,
            'price': 4,
        },
    }

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
        values['supplier_name'] = self.name
        return values

    def get_data(self, path):
        book = open_workbook(path, on_demand=True)
        data = list()
        for sheet_name in book.sheet_names():
            sheet = book.sheet_by_name(sheet_name)
            for row in range(self.START_ROW_INDEX, sheet.nrows):
                values = {}
                if sheet.cell(row, 1).ctype != XL_CELL_EMPTY and \
                        sheet.cell(row, 2).ctype != XL_CELL_EMPTY and \
                        sheet.cell(row, 3).ctype != XL_CELL_EMPTY and \
                        sheet.cell(row, 4).ctype != XL_CELL_EMPTY:
                    values['name'] = unicode(sheet.cell(row, 1).value).strip()
                    values['count'] = int(sheet.cell(row, 2).value)
                    values['code'] = str(sheet.cell(row, 3).value)
                    values['price'] = int(sheet.cell(row, 4).value)

                    values['is_available'] = bool(values['count'])
                    values['sheet'] = sheet.name
                    data.append(values)
                else:
                    continue
            book.unload_sheet(sheet_name)
        return data