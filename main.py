# -*- coding: utf-8 -*-
import imp
import os
import re
import traceback
from datetime import datetime
from decimal import Decimal, InvalidOperation

from django.conf import settings
from django.core.exceptions import ValidationError
from django.db import connection, reset_queries
from django.utils.functional import cached_property

from apps.utils.money import markup_price, roundup

UPLOADER_CLASS_POSTFIX = 'Uploader'
UPLOADERS = dict()
UPLOADERS_CHOICES = list()

self_file = __file__.replace('.pyc', '.py')
file_dir = os.path.dirname(self_file)

uploader_files = [
    os.path.join(file_dir, f) for f in os.listdir(file_dir)
    if os.path.isfile(os.path.join(file_dir, f))
    and os.path.join(file_dir, f) != self_file
    and not f.startswith('__') and f.endswith('.py')]


def convert(name):
    s1 = re.sub('(.)([A-Z][a-z]+)', r'\1_\2', name)
    return re.sub('([a-z0-9])([A-Z])', r'\1_\2', s1).lower()


for uploader_file in uploader_files:
    module_name = os.path.splitext(os.path.basename(uploader_file))[0]
    uploader_module = imp.load_source(module_name, uploader_file)
    for cl in dir(uploader_module):
        if not cl.startswith('__') and cl.endswith(UPLOADER_CLASS_POSTFIX):
            uploader_class = getattr(uploader_module, cl)
            key = convert(cl.replace(UPLOADER_CLASS_POSTFIX, ''))
            UPLOADERS[key] = uploader_class
            try:
                name = uploader_class.name
            except AttributeError:
                name = key
            UPLOADERS_CHOICES.append((key, name))

VALIDATION = {
    'supplier': (lambda x: True if int(x) > 0 else False, lambda x: int(x)),
    'group_uuid': (lambda x: True if len(unicode(x)) <= 16 else False, lambda x: unicode(x)),
    'supplier_name': (lambda x: True if len(unicode(x)) <= 255 else False, lambda x: unicode(x)),
    'is_spare': (lambda x: True, lambda x: bool(x)),
    'category_0': (lambda x: True if len(unicode(x)) <= 255 else False, lambda x: unicode(x)),
    'category': (lambda x: True if len(unicode(x)) <= 255 else False, lambda x: unicode(x)),
    'name': (lambda x: True if len(unicode(x)) <= 1011 else False, lambda x: unicode(x)),
    'sheet': (lambda x: True if len(unicode(x)) <= 255 else False, lambda x: unicode(x)),
    'code': (lambda x: True if len(unicode(x)) <= 255 else False, lambda x: unicode(x)),
    'original_num': (lambda x: True if len(unicode(x)) <= 255 else False, lambda x: unicode(x)),
    'year': (lambda x: True if len(unicode(x)) <= 50 else False, lambda x: unicode(x)),
    'count': (lambda x: 0 <= int(x) < 9999, lambda x: int(x)),
    'count_office': (lambda x: 0 <= int(x) < 9999, lambda x: int(x)),
    'offer_days': (lambda x: 0 <= int(x) < 9999, lambda x: int(x)),
    'order_only': (lambda x: True, lambda x: bool(x)),
    'is_available': (lambda x: True, lambda x: bool(x)),
    'old_price': (lambda x: 0 <= Decimal(x) < 99999, lambda x: Decimal(x)),
    'price': (lambda x: 0 <= Decimal(x) < 99999, lambda x: Decimal(x)),
    'retail_old_price': (lambda x: 0 <= Decimal(x) < 99999, lambda x: Decimal(x)),
    'retail_price': (lambda x: 0 <= Decimal(x) < 99999, lambda x: Decimal(x)),
    'manufacturer': (lambda x: True if len(unicode(x)) <= 255 else False, lambda x: unicode(x)),
    'manufacturer_code': (lambda x: True if len(unicode(x)) <= 255 else False, lambda x: unicode(x)),
    'image_url': (lambda x: True if len(unicode(x)) <= 255 else False, lambda x: unicode(x)),
    'waiting_pechatniki': (lambda x: True if len(unicode(x)) <= 255 else False, lambda x: unicode(x)),
    'analog': (lambda x: True, lambda x: x), }


class Uploader(object):
    """
    Интерфейс для работы с Импортерами
    """

    def __init__(self, supplier_id, importer_key, price_model, markup=0, extramarkups=None):
        """
        Инициализация
        """
        self.without_errors = True
        self.supplier_id = supplier_id
        self.markup = markup
        self.extramarkups = extramarkups
        self.key = importer_key
        self.log_messages = list()
        self.validation_errors = list()
        self.price_model = price_model
        self.updated_count = 0
        self.not_valid = 0
        self.created_count = 0
        self.price_count = 0
        self.counter = 0
        self.timer = datetime.now()
        self.exist_prices.update(
            retail_price=0,
            old_price=0,
            retail_old_price=0,
            price=0,
            count=0,
            offer_days=0,
            is_available=False)

    @property
    def missed_count(self):
        """
        Количество пропущенных (необновленных) прайсовых позиций
        """
        return self.total_count - sum(
            (
                self.updated_count,
                self.created_count,))

    @cached_property
    def total_count(self):
        """
        Количество, существующих прайсовых позицй
        """
        return self.exist_prices.count()

    @property
    def exist_prices(self):
        """
        Queryset, существующих прайсовых позицй
        """
        return self.price_model.objects.filter(supplier_id=self.supplier_id)

    @property
    def log(self):
        """
        Лог в виде html строки
        """
        return u'<br/>'.join(self.log_messages)

    def get_markup(self, price):
        """
        Проверяем список экстранаценок
        и ищем подходящий диапазон цен
        Если не находим возвращаем self.markup
        :param price:
        :return: markup value or None
        """
        extramarkups = self.extramarkups or []
        res = self.markup
        for e in extramarkups:
            if e['min_price'] <= price <= e['max_price']:
                res = e['markup']
                break
        return res

    def start(self, file_path):
        """
        Запуск импорта
        """
        self.logging(u'Файл: {}'.format(
            os.path.basename(file_path)), need_time=False)
        try:
            self.logging(u'Начало импорта')
            data = UPLOADERS[self.key](file_path)
            counter = 0
            self.logging(u'Загрузка по {} прайсов'.format(
                settings.LOAD_PRICE_COUNT))
            for price_kwargs in data:
                counter += 1
                if counter == 1:
                    self.logging(u'Подготовка данных')
                    self._prepare()
                    self.logging(u'Обработка данных')
                if counter < settings.LOAD_PRICE_COUNT:
                    self._process_price(price_kwargs)
                else:
                    counter = 0
                    self._process_price(price_kwargs)
                    self.logging(u'Сохранение данных')
                    self._db_update()
            else:
                self.logging(u'Сохранение данных')
                self._db_update()
                clear_qs = self.exist_prices.filter(
                    variants__isnull=True,
                    price=0,
                    count=0,
                    is_available=False)
                self.logging(
                    u'Удаление {} отсутствующих прайсов'.format(clear_qs.count()))
                clear_qs.delete()

            self.logging(
                u'Импорт завершен'
                u'<div class="alert alert-success">'
                u'Всего позиций в базе: {total}'
                u'<br/>Позиций в прайсе: {price_count}'
                u'<hr>Синхронизировано: {updated}'
                u'<br/>Создано новых: {created}</div>'
                u'<div class="alert alert-warning">'
                u'Неверные данные (ошибки): {not_valid}<br/>'
                u'{validation_errors}'
                u'</div>'.format(
                    total=self.total_count,
                    price_count=self.price_count,
                    updated=self.updated_count,
                    created=self.created_count,
                    not_valid=self.not_valid,
                    validation_errors=u'<br/>'.join(self.validation_errors), )
            )
        except Exception as e:
            self.without_errors = False
            self.logging(
                u'Произошла ошибка: {0}'
                u'<div class="alert alert-error">{1}</div>'.format(
                    e,
                    traceback.format_exc()))

    def _prepare(self):
        """
        Подготовка данных для импорта
        """
        self.queries_to_update = dict()
        self.duplicates = list()
        self.params = tuple()
        self.prices_to_create = dict()

    def _process_price(self, price_kwargs):
        """
        Подготовка Прайса к обновлению
        """
        self.price_count += 1
        not_valid_fields = 0
        if price_kwargs['code']:
            for field, value in price_kwargs.items():
                try:
                    if VALIDATION[field][0](value):
                        price_kwargs[field] = VALIDATION[field][1](value)
                    else:
                        raise ValueError
                except (ValidationError, ValueError, InvalidOperation):
                    not_valid_fields += 1
                    self.validation_errors.append(
                        u'Validation Error - CODE: {}, FIELD: {}'.format(
                            price_kwargs['code'], field))
            if not_valid_fields == 0:
                # Валидация пройдена
                markup = self.get_markup(price_kwargs['price'])
                price_kwargs['retail_price'] = roundup(markup_price(
                    price_kwargs['price'], markup))
                price_kwargs['retail_old_price'] = roundup(markup_price(
                    price_kwargs.get('old_price', 0), markup))
                price_id = u'{}-{}'.format(self.supplier_id,
                                           price_kwargs['code'])
                try:
                    price = self.price_model.objects.get(
                        supplier_id=self.supplier_id,
                        code=price_kwargs.get('code'))
                    sql_with_params = self.price_model.objects.get_update_query(
                        pk=price.pk, **price_kwargs)
                    if self.queries_to_update.get(price_id):
                        self.duplicates.append(price_kwargs['code'])
                    else:
                        self.queries_to_update[price_id] = sql_with_params[0]
                        self.params += sql_with_params[1]
                        self.updated_count += 1
                except self.price_model.DoesNotExist:
                    price_kwargs['supplier_id'] = self.supplier_id
                    if self.prices_to_create.get(price_id):
                        self.duplicates.append(price_kwargs['code'])
                    else:
                        self.prices_to_create[price_id] = self.price_model(
                            **price_kwargs)
                        self.created_count += 1
            else:
                self.not_valid += 1
        else:
            # Данные не валидны
            self.not_valid += 1

    def _db_update(self):
        """
        Обновление данных в базе
        """
        query = u';'.join(self.queries_to_update.values())
        if query:
            self.logging(u'...Обновление данных в базе')
            cursor = connection.cursor()
            try:
                cursor.execute(query, self.params)
            finally:
                cursor.close()
        if self.prices_to_create:
            self.logging(u'...Создание данных в базе')
            self.price_model.objects.bulk_create(
                self.prices_to_create.values())
        reset_queries()

    def logging(self, message, need_time=True):
        """
        Логирование сообщения
        """
        if need_time:
            message = u'{} - {}'.format(
                datetime.now().strftime('%d.%m.%Y %H:%M:%S'),
                message)
        self.log_messages.append(message)
