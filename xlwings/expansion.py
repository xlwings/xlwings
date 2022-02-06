from .main import Range

expanders = {}

_empty = (None, "", [[""]], [[None]])


class Expander:
    def register(self, *aliases):
        for alias in aliases:
            expanders[alias] = self

    def expand(self, rng):
        """
        Expands a range

        Arguments
        ---------
        rng: Range
            The reference range

        Returns
        -------
        Range object: The expanded range

        """
        raise NotImplemented()


class TableExpander(Expander):
    def expand(self, rng):
        origin = rng(1, 1)

        if origin.has_array:
            bottom_left = origin.end("down")
        elif origin(2, 1).raw_value in _empty:
            bottom_left = origin
        elif origin(3, 1).raw_value in _empty:
            bottom_left = origin(2, 1)
        else:
            bottom_left = origin(2, 1).end("down")

        if origin.has_array:
            top_right = origin.end("right")
        elif origin(1, 2).raw_value in _empty:
            top_right = origin
        elif origin(1, 3).raw_value in _empty:
            top_right = origin(1, 2)
        else:
            top_right = origin(1, 2).end("right")

        return Range(top_right, bottom_left)


TableExpander().register("table")


class VerticalExpander(Expander):
    def expand(self, rng):
        if rng(2, 1).raw_value in _empty:
            return Range(rng(1, 1), rng(1, rng.shape[1]))
        elif rng(3, 1).raw_value in _empty:
            return Range(rng(1, 1), rng(2, rng.shape[1]))
        else:
            end_row = rng(2, 1).end("down").row - rng.row + 1
            return Range(rng(1, 1), rng(end_row, rng.shape[1]))


VerticalExpander().register("vertical", "down", "d")


class HorizontalExpander(Expander):
    def expand(self, rng):
        if rng(1, 2).raw_value in _empty:
            return Range(rng(1, 1), rng(rng.shape[0], 1))
        elif rng(1, 3).raw_value in _empty:
            return Range(rng(1, 1), rng(rng.shape[0], 2))
        else:
            end_column = rng(1, 2).end("right").column - rng.column + 1
        return Range(rng(1, 1), rng(rng.shape[0], end_column))


HorizontalExpander().register("horizontal", "right", "r")
