# -*- coding: utf-8 -*-
from xlrd import open_workbook
from xlrd.sheet import ctype_text


class RedPowerUploader(object):

    CATEGORY_PARENT_ROW_COLOR_1 = 40
    CATEGORY_PARENT_ROW_COLOR_2 = 44
    CATEGORY_ROW_COLOR = 30
    PRODUCT_ROW_COLOR = 9

    # row, column это индекс, начинаются с 0
    # IMAGE_URL = u'http://www.autobody.ru/upload/images/{code}.jpg'
    name = u'Red Power'
    START_ROW_INDEX = 7

    sheets = [u'TDSheet',]

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
            'name': 1,
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
        # values['image_url'] = self.IMAGE_URL.format(code=values['code'])
        values['supplier_name'] = self.name
        return values

    def check_row_content_and_type(self, book, row):
        colors_list = [self.CATEGORY_PARENT_ROW_COLOR_1, self.CATEGORY_PARENT_ROW_COLOR_2,
                       self.CATEGORY_ROW_COLOR]
        # тип данных в ячейке
        code_obj = row[3]
        price_obj = row[7]
        code_cell_type = ctype_text.get(code_obj.ctype, 'unknown type')
        price_cell_type = ctype_text.get(price_obj.ctype, 'unknown type')

        # объект с набором стилей ячейки
        fmt = book.xf_list[code_obj.xf_index]
        color_id = fmt.background.pattern_colour_index

        # отфильтровываем строки без артикула или без цены,
        # если это не строка с категорией
        if (code_cell_type in ['empty', 'unknown type', 'blank']
            or price_cell_type in ['empty', 'unknown type', 'blank']) \
                and color_id not in colors_list:
            return None, None
        elif color_id == self.CATEGORY_ROW_COLOR:
            return True, 'category'
        elif color_id in [self.CATEGORY_PARENT_ROW_COLOR_1, self.CATEGORY_PARENT_ROW_COLOR_2]:
            return True, 'parent_category'
        else:
            return True, 'product'

    def get_data(self, path):
        book = open_workbook(path, on_demand=True, formatting_info=True)
        data = list()
        for sheet_name in book.sheet_names():
            if sheet_name in self.sheets:
                category = ''
                category_0 = ''
                sheet = book.sheet_by_name(sheet_name)

                for row_n in range(self.START_ROW_INDEX, sheet.nrows):
                    row = sheet.row(row_n)
                    content, row_type = self.check_row_content_and_type(book, row)
                    if not content:
                        continue
                    if row_type == 'category':
                        category = unicode(row[1].value)
                    elif row_type == 'parent_category':
                        category_0 = unicode(row[1].value)
                    else:
                        price = row[8].value
                        count = row[2].value
                        values = {
                            'sheet': sheet_name,
                            'category_0': category_0,
                            'category': category,
                            'manufacturer': None,
                            'manufacturer_code': unicode(row[3].value),
                            'is_available': True if count > 0 else False,
                            'name': unicode(row[1].value),
                            'code': unicode(row[3].value),
                            'year': None,
                            'count': count,
                            'price': price,
                        }
                        data.append(values)

                book.unload_sheet(sheet_name)
        return data

