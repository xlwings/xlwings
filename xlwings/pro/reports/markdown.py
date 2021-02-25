import textwrap
from ...conversion import Converter


class MarkdownOptions:
    class __Heading1:
        def __init__(self):
            self.color = None
            self.size = None
            self.bold = None
            self.italic = None
            self.blank_lines_after = 1

        def __repr__(self):
            return f"""\
                    h1.color: {self.color}
                    h1.size: {self.size}
                    h1.bold: {self.bold}
                    h1.italic: {self.italic}
                    h1.blank_lines_after: {self.blank_lines_after}
                    """

    def __init__(self):
        self.h1 = self.__Heading1()

    def __repr__(self):
        return textwrap.dedent(repr(self.h1))


class MarkdownConverter(Converter):

    @classmethod
    def write_value(cls, value, options):

        value = value.replace('**', '')

        return value


MarkdownConverter.register('markdown', 'md')


class FormatMarkdownStage:
    def __init__(self, options):
        self.options = options

    def __call__(self, ctx):
        md_options = self.options['md_options']
        if ctx.range:
            if md_options.h1.color:
                ctx.range.font.color = md_options.h1.color
