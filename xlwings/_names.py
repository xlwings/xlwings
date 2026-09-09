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
        current_name = get_name(self.index)
        if current_name != self.name:
            try:
                self.index = get_names().index(self.name) + 1
            except ValueError:
                # A sheet rename changes the scope label without replacing the
                # entry. Prefer an exact match elsewhere before accepting this.
                if (
                    current_name is not None
                    and "!" in self.name
                    and "!" in current_name
                    and self.name.rsplit("!", 1)[-1] == current_name.rsplit("!", 1)[-1]
                ):
                    self.name = current_name
                else:
                    raise KeyError(self.name) from None
        return self.index
