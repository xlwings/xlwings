import inspect

import xlwings


class ConversionContext:
    __slots__ = ["range", "value", "source_value", "meta", "engine"]

    def __init__(self, rng=None, value=None, engine_name=None):
        self.range = rng
        if engine_name:
            self.engine = xlwings.engines[engine_name]
        else:
            # rng can only be None if used via COM server UDFs
            self.engine = (
                rng.sheet.book.app.engine
                if rng is not None
                else xlwings.engines["excel"]
            )
        self.value = value
        # used for markdown (could be replaced by handing the parsed ast from
        # the converter stage to the formatting stage
        self.source_value = value
        self.meta = {}


class Options(dict):
    def __init__(self, original):
        super(Options, self).__init__(original)

    def override(self, **overrides):
        self.update(overrides)
        return self

    def erase(self, keys):
        for key in keys:
            self.pop(key, None)
        return self

    def defaults(self, **defaults):
        for k, v in defaults.items():
            self.setdefault(k, v)
        return self


class Pipeline(list):
    def prepend_stage(self, stage, only_if=True):
        if only_if:
            self.insert(0, stage)
        return self

    def append_stage(self, stage, only_if=True):
        if only_if:
            self.append(stage)
        return self

    def insert_stage(
        self, stage, index=None, after=None, before=None, replace=None, only_if=True
    ):
        if only_if:
            if sum(x is not None for x in (index, after, before, replace)) != 1:
                raise ValueError(
                    "Must specify exactly one of arguments: "
                    "index, after, before, replace"
                )
            if index is not None:
                indices = (index,)
            elif after is not None:
                indices = tuple(
                    i + 1 for i, x in enumerate(self) if isinstance(x, after)
                )
            elif before is not None:
                indices = tuple(i for i, x in enumerate(self) if isinstance(x, before))
            elif replace is not None:
                for i, x in enumerate(self):
                    if isinstance(x, replace):
                        self[i] = stage
                return self
            for i in reversed(indices):
                self.insert(i, stage)
        return self

    def __call__(self, *args, **kwargs):
        for stage in self:
            stage(*args, **kwargs)

    async def async_call(self, *args, **kwargs):
        # Two stage conventions are supported here:
        # - Stages with an `async_run` hook (e.g. ToValueStage, whose write_value
        #   may be sync or async) - call it so the awaited result is assigned
        #   correctly; its sync `__call__` would reject an async converter.
        # - Stages without one (e.g. the Async* read stages) - call directly and
        #   await only if the result is awaitable. Plain sync stages return None.
        for stage in self:
            async_run = getattr(stage, "async_run", None)
            if async_run is not None:
                await async_run(*args, **kwargs)
            else:
                result = stage(*args, **kwargs)
                if inspect.isawaitable(result):
                    await result


accessors = {}


class Accessor:
    @classmethod
    def reader(cls, options):
        return Pipeline()

    @classmethod
    def writer(cls, options):
        return Pipeline()

    @classmethod
    def register(cls, *types):
        for type in types:
            accessors[type] = cls

    @classmethod
    def router(cls, value, rng, options):
        return cls


class Converter(Accessor):
    class ToValueStage:
        def __init__(self, write_value, options):
            self.write_value = write_value
            self.options = options
            self.is_async = inspect.iscoroutinefunction(write_value)

        def __call__(self, c):
            if self.is_async:
                # Programming error: someone defined `async def write_value` on a
                # Converter but used it on a synchronous write (e.g. the
                # `Range.value = ...` setter or a desktop UDF return value).
                # Async converters only work through the async write path
                # (custom functions in xlwings Lite). Either make write_value a
                # regular `def`, or write the value via that async path.
                name = getattr(self.write_value, "__qualname__", self.write_value)
                raise TypeError(
                    f"{name} is defined as `async def` but is being used on a "
                    "synchronous write (e.g. `Range.value = ...` or a COM "
                    "UDF). Async write_value converters are only supported by "
                    "custom functions in xlwings Lite. Use a regular `def "
                    "write_value` for synchronous writes."
                )
            c.value = self.write_value(c.value, self.options)

        async def async_run(self, c):
            if self.is_async:
                c.value = await self.write_value(c.value, self.options)
            else:
                c.value = self.write_value(c.value, self.options)

    class FromValueStage:
        def __init__(self, read_value, options):
            self.read_value = read_value
            self.options = options

        def __call__(self, c):
            c.value = self.read_value(c.value, self.options)

    base_type = None
    base = None

    @classmethod
    def base_reader(cls, options, base_type=None):
        if cls.base is not None:
            return cls.base.reader(options)
        else:
            return accessors[base_type or cls.base_type].reader(options)

    @classmethod
    def base_writer(cls, options, base_type=None):
        if cls.base is not None:
            return cls.base.writer(options)
        else:
            return accessors[base_type or cls.base_type].writer(options)

    @classmethod
    def reader(cls, options):
        return cls.base_reader(options).append_stage(
            cls.FromValueStage(cls.read_value, options)
        )

    @classmethod
    def writer(cls, options):
        return cls.base_writer(options).prepend_stage(
            cls.ToValueStage(cls.write_value, options)
        )
