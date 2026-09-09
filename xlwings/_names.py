"""Index tracking for engines with lazy native name references."""


class NameIndex:
    def __init__(self, index, name):
        self.index = index
        self.name = name

    def resolve(self, get_name, get_names):
        """Validate a one-based index; rescan only when its name has moved.

        get_name returns None for an index that no longer exists. Name strings
        must distinguish scopes, even if native by-name lookup ignores scope.
        """
        if get_name(self.index) != self.name:
            try:
                self.index = get_names().index(self.name) + 1
            except ValueError:
                raise KeyError(self.name) from None
        return self.index
