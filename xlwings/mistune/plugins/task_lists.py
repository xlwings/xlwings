import re

__all__ = ['plugin_task_lists']


TASK_LIST_ITEM = re.compile(r'^(\[[ xX]\])\s(\s*\S.*)')


def task_lists_hook(md, tokens, state):
    for tok in tokens:
        if tok['type'] == 'list':
            for item in tok['children']:
                _rewrite_list_item(item)
    return tokens


def render_ast_task_list_item(children, level, checked):
    return {
        'type': 'task_list_item',
        'children': children,
        'level': level,
        'checked': checked,
    }


def render_html_task_list_item(text, level, checked):
    checkbox = (
        '<input class="task-list-item-checkbox" '
        'type="checkbox" disabled'
    )
    if checked:
        checkbox += ' checked/>'
    else:
        checkbox += '/>'

    if text.startswith('<p>'):
        text = text.replace('<p>', '<p>' + checkbox, 1)
    else:
        text = checkbox + text

    return '<li class="task-list-item">' + text + '</li>\n'


def plugin_task_lists(md):
    md.before_render_hooks.append(task_lists_hook)

    if md.renderer.NAME == 'html':
        md.renderer.register('task_list_item', render_html_task_list_item)
    elif md.renderer.NAME == 'ast':
        md.renderer.register('task_list_item', render_ast_task_list_item)


def _rewrite_list_item(item):
    if item['type'] != 'list_item':
        return

    children = item['children']
    if children:
        first_child = children[0]
        text = first_child.get('text', '')
        m = TASK_LIST_ITEM.match(text)
        if m:
            mark = m.group(1)
            first_child['text'] = m.group(2)

            params = item['params']
            if mark == '[ ]':
                params = (params[0], False)
            else:
                params = (params[0], True)

            item['type'] = 'task_list_item'
            item['params'] = params
