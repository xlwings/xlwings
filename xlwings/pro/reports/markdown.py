import textwrap

from ... import mistune
from ...conversion import Converter


class MarkdownOptions:
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


class MarkdownConverter(Converter):

    @classmethod
    def write_value(cls, value, options):
        md_options = options['md_options']
        flat_ast = flatten_ast(value)

        # for i in flat_ast:
        #     print(i['parent_type'])
        #     print(i['length'])
        #     print(i['text'])

        output = ''
        for node in flat_ast:
            # heading/list currently don't respect the level
            if 'heading' in node['parent_type'][0]:
                output += ''.join(node['text'])
                output += '\n' + md_options.h1.blank_lines_after * '\n'
            elif 'paragraph' in node['parent_type'][0]:
                output += ''.join(node['text'])
                output += '\n' + md_options.paragraph.blank_lines_after * '\n'
            elif 'list' in node['parent_type'][0]:
                for j in node['text']:
                    output += f'\u2022 {j}\n'
                output += md_options.unordered_list.blank_lines_after * '\n'

        value = output.rstrip('\n')
        return value


MarkdownConverter.register('markdown', 'md')


class FormatMarkdownStage:
    def __init__(self, options):
        self.options = options

    def __call__(self, ctx):
        md_options = self.options['md_options']
        assert ctx.meta['scalar']
        flat_ast = flatten_ast(ctx.source_value)

        position = 0
        for node in flat_ast:
            if 'heading' in node['parent_type'][0]:
                node_length = sum(node['length']) + md_options.h1.blank_lines_after + 1
                # TODO: loop over all font characteristics
                ctx.range.characters[position:position + node_length].font.color = md_options.h1.color
                ctx.range.characters[position:position + node_length].font.size = md_options.h1.size
            elif 'paragraph' in node['parent_type'][0]:
                node_length = sum(node['length']) + md_options.paragraph.blank_lines_after + 1
                intra_node_position = position
                for ix, j in enumerate(node['parent_type']):
                    if 'strong' in j:
                        selection = slice(intra_node_position, intra_node_position + node['length'][ix])
                        ctx.range.characters[selection].font.bold = True  # TODO: take from options
                    elif 'emphasis' in j:
                        ctx.range.characters[selection].font.italic = True  # TODO: take from options
                    intra_node_position += node['length'][ix]
            elif 'list' in node['parent_type'][0]:
                node_length = sum(node['length']) + md_options.unordered_list.blank_lines_after
                for j in node['text']:
                    # TODO: check ast level to allow nested **strong** etc.
                    node_length += 3  # bullet, space and new line
            position += node_length


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
