import textwrap

from ... import mistune
from ...conversion import Converter


class MarkdownStyle:
    # TODO: repr
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

    class __Strong:
        def __init__(self):
            self.color = None
            self.size = None
            self.bold = True
            self.italic = None

    class __Emphasis:
        def __init__(self):
            self.color = None
            self.size = None
            self.bold = None
            self.italic = True

    class __Paragraph:
        def __init__(self):
            self.blank_lines_after = 1

    class __UnorderedList:
        def __init__(self):
            self.bullet_character = '\u2022'
            self.blank_lines_after = 1

    def __init__(self):
        self.h1 = self.__Heading1()
        self.paragraph = self.__Paragraph()
        self.unordered_list = self.__UnorderedList()
        self.strong = self.__Strong()
        self.emphasis = self.__Emphasis

    def __repr__(self):
        return textwrap.dedent(repr(self.h1))


class Markdown:
    def __init__(self, text, style=MarkdownStyle()):
        self.text = text
        self.style = style


class MarkdownConverter(Converter):

    @classmethod
    def write_value(cls, value, options):
        return render_text(value, options['style'])


MarkdownConverter.register('markdown', 'md')


class FormatMarkdownStage:
    def __init__(self, options):
        self.options = options

    def __call__(self, ctx):
        format_text(ctx.range, ctx.source_value, self.options['style'])


def traverse_ast_node(tree, data=None, level=0):
    data = {'length': [], 'type': [], 'parent_type': [],
            'text': [], 'parents': [], 'level': []} if data is None else data
    for element in tree:
        data['parents'] = data['parents'][:level]
        if 'children' in element:
            data['parents'].append(element)
            traverse_ast_node(element['children'], data, level=level + 1)
        else:
            data['level'].append(level)
            data['parent_type'].append([parent['type'] for parent in data['parents']])
            data['type'].append(element['type'])
            if element['type'] == 'text':
                data['length'].append(len(element['text']))
                data['text'].append(element['text'])
            elif element['type'] == 'linebreak':
                data['length'].append(1)
                data['text'].append('\n')
    return data


def flatten_ast(value):
    parse_ast = mistune.create_markdown(renderer=mistune.AstRenderer())
    ast = parse_ast(value)
    flat_ast = []
    for node in ast:
        rv = traverse_ast_node([node])
        del rv['parents']
        flat_ast.append(rv)
    return flat_ast


def render_text(text, options):
    flat_ast = flatten_ast(text)


    # for i in flat_ast:
    #     print(i['parent_type'])
    #     print(i['length'])
    #     print(i['text'])

    output = ''
    for node in flat_ast:
        # heading/list currently don't respect the level
        if 'heading' in node['parent_type'][0]:
            output += ''.join(node['text'])
            output += '\n' + options.h1.blank_lines_after * '\n'
        elif 'paragraph' in node['parent_type'][0]:
            output += ''.join(node['text'])
            output += '\n' + options.paragraph.blank_lines_after * '\n'
        elif 'list' in node['parent_type'][0]:
            for j in node['text']:
                output += f'\u2022 {j}\n'
            output += options.unordered_list.blank_lines_after * '\n'

    return output.rstrip('\n')


def format_text(parent, text, style):
    flat_ast = flatten_ast(text)

    position = 0
    for node in flat_ast:
        if 'heading' in node['parent_type'][0]:
            node_length = sum(node['length']) + style.h1.blank_lines_after + 1
            apply_style_to_font(style.h1,
                                parent.characters[position:position + node_length].font)
        elif 'paragraph' in node['parent_type'][0]:
            node_length = sum(node['length']) + style.paragraph.blank_lines_after + 1
            intra_node_position = position
            for ix, j in enumerate(node['parent_type']):
                selection = slice(intra_node_position, intra_node_position + node['length'][ix])
                if 'strong' in j:
                    apply_style_to_font(style.strong, parent.characters[selection].font)
                elif 'emphasis' in j:
                    apply_style_to_font(style.emphasis, parent.characters[selection].font)
                intra_node_position += node['length'][ix]
        elif 'list' in node['parent_type'][0]:
            node_length = sum(node['length']) + style.unordered_list.blank_lines_after
            for _ in node['text']:
                # TODO: check ast level to allow nested **strong** etc.
                node_length += 3  # bullet, space and new line
        position += node_length


def apply_style_to_font(style_object, font_object):
    for attribute in vars(style_object):
        if getattr(style_object, attribute):
            setattr(font_object, attribute, getattr(style_object, attribute))
