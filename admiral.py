# -*- coding: utf-8 -*-
from xlrd import open_workbook, XL_CELL_EMPTY, XL_CELL_TEXT


class AdmiralUploader(object):
    START_ROW_INDEX = 8
    name = u'Адмирал'

    config = {
        'map': {
            'sheet': None,
            'category': None,
            'manufacturer': 8,
            'manufacturer_code': 1,
            'supplier_name': None,
            'is_available': None,
            'original_num': 0,
            'name': 2,
            'code': 1,
            'year': None,
            'count': 7,
            'price': 6, }, }

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
            category = ''
            sheet = book.sheet_by_name(sheet_name)
            for row in range(self.START_ROW_INDEX, sheet.nrows):
                values = {}
                for col_name, col in self.config['map'].items():
                    if col is not None:
                        cell = sheet.cell(row, col)
                        if cell.ctype != XL_CELL_EMPTY:
                            if cell.ctype == XL_CELL_TEXT:
                                values[col_name] = cell.value.strip()
                            elif col_name == 'code':
                                try:
                                    values[col_name] = unicode(int(cell.value))
                                except ValueError:
                                    values[col_name] = cell.value
                            else:
                                values[col_name] = cell.value

                if len(values) > 1:
                    values['category'] = category
                    try:
                        if values['count'] == '+':
                            values['count'] = 1
                        values['is_available'] = int(values['count']) > 0
                    except (TypeError, ValueError, KeyError):
                        values['count'] = 0
                        values['is_available'] = False
                    data.append(values)
                elif len(values) == 1:
                    category = values.values()[0]
                    continue
                else:
                    continue
            book.unload_sheet(sheet_name)
        return data
