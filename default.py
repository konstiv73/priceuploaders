# -*- coding: utf-8 -*-
from decimal import Decimal
from xlrd import open_workbook, XL_CELL_EMPTY, XL_CELL_TEXT


class DefaultUploader(object):
    name = u'Сводный прайс'
    START_ROW_INDEX = 2
    START_MANUFACTURER_COL = 8
    MANUFACTURER_ROW = 1

    config = {
        'map': {
            'category_0': None,
            'category': None,
            'manufacturer_code': None,
            'supplier_name': None,
            'is_available': None,
            'code': 2,
            'count': 5,
            'count_office': 6,
            'old_price': 4,
            'offer_days': 7,
            'order_only': None,
            'price': 3, }, }

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
        return values

    def get_data(self, path):
        book = open_workbook(path, on_demand=True)
        data = list()
        supplier_name = u''
        sheet_name = book.sheet_names()[0]
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
                                values[col_name] = unicode(int(cell.value))
                            except ValueError:
                                values[col_name] = cell.value
                        else:
                            values[col_name] = cell.value
            if values.get('price', None):
                values['supplier_name'] = supplier_name
                for col in manufacturers.keys():
                    cell = sheet.cell(row, col)
                    if cell.ctype != XL_CELL_EMPTY:
                        values['manufacturer_code'] = cell.value
                        values['manufacturer'] = manufacturers[col]
                try:
                    if Decimal(values['old_price']) > 0:
                        values['price'], values['old_price'] = (
                            values['old_price'], values['price'])
                except ValueError:
                    values['old_price'] = 0
                try:
                    if values['count'] == '+':
                        values['count'] = 1
                except (TypeError, ValueError, KeyError):
                    values['count'] = 0
                try:
                    if values['count_office'] == '+':
                        values['count_office'] = 1
                except (TypeError, ValueError, KeyError):
                    values['count_office'] = 0
                try:
                    values['offer_days'] = int(values.get('offer_days'))
                    values['order_only'] = values['offer_days'] > 0
                except (TypeError, ValueError, KeyError):
                    values['offer_days'] = 0
                    values['order_only'] = False

                values['is_available'] = int(values.get('count', 0)) > 0 \
                    or int(values.get('count_office', 0)) > 0
                data.append(values)
            elif values.get('price') is None \
                    and values.get('old_price') is None \
                    and values.get('count') is None \
                    and values.get('count_office') is None \
                    and values.get('count_office') is None:
                supplier_name = values['code']
        book.unload_sheet(sheet_name)
        return data
