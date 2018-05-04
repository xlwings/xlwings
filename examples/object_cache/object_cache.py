import xlwings as xw


@xw.func
def object_cache_status():
    return xw.object_cache.ObjectCache.status()


class MyClass(object):

    @classmethod
    def get_negated(cls, my_obj):
        return cls(-my_obj.value)

    def __init__(self, value):
        self._value = value

    @property
    def value(self):
        return self._value

    def get_double(self):
        return 2 * self._value


@xw.return_cached_object
@xw.func
def MyClass_new(value):
    return MyClass(value)


@xw.use_cached_object(object_class=MyClass)
@xw.func
def MyClass_value(my_obj):
    return my_obj.value


@xw.use_cached_object(object_class=MyClass)
@xw.func
def MyClass_get_double(my_obj):
    return my_obj.get_double()


@xw.use_cached_object_list(object_class=MyClass)
@xw.return_cached_object
@xw.func
def MyClass_aggregate(my_objs):
    return MyClass(sum(my_obj.value for my_obj in my_objs))

