"""Show implementation call forms when types belong in parameter descriptions."""

import importlib
import inspect

from sphinx import addnodes


def use_implementation_signature(app, domain, objtype, contentnode):
    """Replace overloaded function headings with one unannotated call form."""
    if domain != "py" or objtype != "function":
        return
    if app.config.autodoc_typehints != "description":
        return
    description = contentnode.parent
    signatures = [
        child for child in description if isinstance(child, addnodes.desc_signature)
    ]
    if len(signatures) < 2:
        return
    heading = signatures[0]
    module = heading.get("module")
    fullname = heading.get("fullname")
    if not module or not fullname:
        return
    try:
        obj = importlib.import_module(module)
        for part in fullname.split("."):
            obj = getattr(obj, part)
        if not inspect.isfunction(obj):
            return
        parameters = list(inspect.signature(obj).parameters.values())
    except (ImportError, AttributeError, TypeError, ValueError):
        return

    arglist = addnodes.desc_parameterlist()
    keyword_separator = False
    for index, parameter in enumerate(parameters):
        if parameter.kind == inspect.Parameter.VAR_POSITIONAL:
            keyword_separator = True
        elif parameter.kind == inspect.Parameter.KEYWORD_ONLY and not keyword_separator:
            arglist += addnodes.desc_parameter("", "*")
            keyword_separator = True
        argument = str(parameter.replace(annotation=inspect.Parameter.empty))
        arglist += addnodes.desc_parameter("", argument)
        if parameter.kind == inspect.Parameter.POSITIONAL_ONLY and (
            index + 1 == len(parameters)
            or parameters[index + 1].kind != inspect.Parameter.POSITIONAL_ONLY
        ):
            arglist += addnodes.desc_parameter("", "/")

    for child in list(heading):
        if isinstance(child, addnodes.desc_parameterlist):
            heading.replace(child, arglist)
        elif isinstance(child, addnodes.desc_returns):
            heading.remove(child)
    for extra in signatures[1:]:
        heading["ids"].extend(extra.get("ids", []))
        description.remove(extra)


def setup(app):
    app.connect("object-description-transform", use_implementation_signature)
    return {"parallel_read_safe": True, "parallel_write_safe": True}
