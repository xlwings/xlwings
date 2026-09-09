"""Index tracking for engines with lazy native name references."""


class NameIndex:
    def __init__(self, index, name, sheet_names=()):
        self.index = index
        self.name = name
        self.sheet_names = tuple(sheet_names)
        scope, _ = self.split_name(name)
        try:
            self.sheet_index = self.sheet_names.index(scope)
        except ValueError:
            self.sheet_index = None

    @staticmethod
    def split_name(name):
        scope, separator, local = name.rpartition("!")
        if not separator:
            return None, local
        if scope.startswith("'") and scope.endswith("'"):
            scope = scope[1:-1].replace("''", "'")
        return scope, local

    def resolve(self, get_name, get_names, get_sheet_names=None):
        """Validate a one-based name index and require an exact scoped match.

        get_name returns None for an index that no longer exists. A worksheet
        snapshot identifies the scope after a rename; the old name index never
        supplies a replacement scope. Ambiguous worksheet changes raise KeyError.
        """
        if get_name(self.index) == self.name:
            return self.index
        names = get_names()
        try:
            self.index = names.index(self.name) + 1
            return self.index
        except ValueError:
            pass
        if self.sheet_index is not None and get_sheet_names is not None:
            sheets = tuple(get_sheet_names())
            # Surviving worksheet names must retain their original positions.
            # Allow independent renames, but reject observable reorders/removals.
            positions = {name: i for i, name in enumerate(self.sheet_names)}
            if len(sheets) == len(self.sheet_names) and all(
                name not in positions or positions[name] == i
                for i, name in enumerate(sheets)
            ):
                scope = sheets[self.sheet_index]
                _, local = self.split_name(self.name)
                matches = [
                    (i, name)
                    for i, name in enumerate(names, 1)
                    if self.split_name(name) == (scope, local)
                ]
                if len(matches) == 1:
                    self.index, self.name = matches[0]
                    self.sheet_names = sheets
                    return self.index
        raise KeyError(self.name)
