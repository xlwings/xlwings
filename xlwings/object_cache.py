import functools
import logging
import xlwings


_logger = logging.getLogger(__name__)


class CachedObjectError(Exception):
    pass


class ObjectCache(object):

    _objects = dict()
    _cells = dict()

    @staticmethod
    def _get_obj_id(obj):
        """
        Returns the id for an object stored in the cache.

        :param obj: Cached object to retrieve the id from.
        :return str: Id for the given cached object.
        """
        return '<%s instance at 0x%x>' % (getattr(obj, '__class__', type(obj)).__name__, id(obj))

    @classmethod
    def status(cls):
        return 'ObjectCache holds {} cached objects referred by {} cells'.format(
            len(cls._objects), sum(len(referring_cells) for _, referring_cells in cls._objects.values()))

    @classmethod
    def update(cls, obj):
        book = xlwings.Book.caller()
        workbook_name = book.name
        sheet_name = book.api.ActiveSheet.Name
        caller_address = xlwings.Range.caller_address()

        obj_id = cls._get_obj_id(obj)

        cls.delete(workbook_name, sheet_name, caller_address)

        _logger.debug(
            "Adding cached object with id {!r} to [{}]{}!{}".format(obj_id, workbook_name, sheet_name, caller_address))

        cls._objects[obj_id] = (obj, {(workbook_name, sheet_name, caller_address): None})
        cls._cells.setdefault(workbook_name, dict()).setdefault(sheet_name, dict())[caller_address] = obj_id

        _logger.debug(cls.status())

        return obj_id

    @classmethod
    def get(cls, obj_id):
        try:
            return cls._objects[obj_id][0]
        except KeyError:
            raise CachedObjectError("Unknown object id {!r}".format(obj_id))

    @classmethod
    def delete(cls, workbook_name, sheet_name, cell_address):
        try:
            obj_id = cls._cells[workbook_name][sheet_name][cell_address]
        except KeyError:
            return

        _logger.debug(
            "Removing cached object with id {!r} at [{}]{}!{}".format(obj_id, workbook_name, sheet_name, cell_address))

        _, referring_cells = cls._objects[obj_id]
        del referring_cells[(workbook_name, sheet_name, cell_address)]
        if not referring_cells:
            del cls._objects[obj_id]

        workbook_cache = cls._cells[workbook_name]
        sheet_cache = workbook_cache[sheet_name]
        del sheet_cache[cell_address]
        if not sheet_cache:
            del workbook_cache[sheet_name]
        if not workbook_cache:
            del cls._cells[workbook_name]

        _logger.debug(cls.status())

    @classmethod
    def delete_all(cls, workbook_name, sheet_name=None, predicate=None):
        workbook_cache = cls._cells.get(workbook_name)
        if workbook_cache is None:
            return
        if sheet_name is None:
            sheet_names = workbook_cache.keys()
        else:
            sheet_names = [sheet_name]
        for sheet_name in sheet_names:
            sheet_cache = workbook_cache.get(sheet_name)
            if sheet_cache is None:
                continue
            for cell_address, obj_id in sheet_cache.items():
                if predicate is None or predicate(cell_address, obj_id):
                    cls.delete(workbook_name, sheet_name, cell_address)


class ObjectCacheEventHandler(object):
    __metaclass__ = xlwings.EventHandler

    @classmethod
    def WorkbookOpen(cls, book_):
        book = xlwings.Book.caller()
        workbook_name = book.name
        ObjectCache.delete_all(workbook_name)

    @classmethod
    def WorkbookNewSheet(cls, book_, sheet_):
        book = xlwings.Book.caller()
        workbook_name = book.name
        sheet_name = book.api.ActiveSheet
        ObjectCache.delete_all(workbook_name, sheet_name)

    # Review https://www.pyxll.com/docs/examples/objectcache.html for more events.


def return_cached_object(func):

    @functools.wraps(func)
    def func_wrapper(*args, **kwargs):
        return ObjectCache.update(func(*args, **kwargs))

    return func_wrapper


def use_cached_object(method=None, object_class=None):

    def use_cached_object_decorator(func):

        @functools.wraps(func)
        def func_wrapper(obj_id, *args, **kwargs):
            obj = ObjectCache.get(obj_id)
            if object_class and not isinstance(obj, object_class):
                raise CachedObjectError("Cached object with id '{}' is not an instance of '{}'".format(
                    obj_id, object_class.__name__))
            return func(obj, *args, **kwargs)

        return func_wrapper

    if method is None:
        return use_cached_object_decorator
    else:
        return use_cached_object_decorator(method)


def use_cached_object_list(method=None, object_class=None):

    def use_cached_object_list_decorator(func):

        @functools.wraps(func)
        def func_wrapper(obj_ids, *args, **kwargs):
            if not hasattr(obj_ids, '__iter__'):
                obj_ids = [obj_ids]
            objs = list()
            for obj_id in obj_ids:
                obj = ObjectCache.get(obj_id)
                if object_class and not isinstance(obj, object_class):
                    raise CachedObjectError("Cached object with id '{}' is not an instance of '{}'".format(
                        obj_id, object_class.__name__))
                objs.append(obj)
            return func(objs, *args, **kwargs)

        return func_wrapper

    if method is None:
        return use_cached_object_list_decorator
    else:
        return use_cached_object_list_decorator(method)

