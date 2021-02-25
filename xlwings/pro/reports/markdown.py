from ...conversion import Converter


class MarkdownConverter(Converter):

    @classmethod
    def write_value(cls, value, options):

        value = value.replace('**', '')

        return value


MarkdownConverter.register('markdown')


class FormatMarkdownStage:
    def __init__(self, options):
        self.options = options

    def __call__(self, ctx):
        if ctx.range:
            ctx.range.font.color = (255, 0, 0)
