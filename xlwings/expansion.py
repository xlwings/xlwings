from .main import Range

expanders = {}

_empty = (None, '')


class Expander(object):

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

    def clear(self, rng, skip, vshape):
        """
        Clears out existing data corresponding to the expansion, in preparation for writing a value

        Arguments
        ---------
        rng: Range
            The reference range
        skip: tuple
            The number of rows, cols to skip when clearing (UDF formula caller range)
        vshape: tuple
            The number of rows, cols of the value which will be written subsequently
        """
        raise NotImplemented()


class TableExpander(Expander):

    def expand(self, rng):
        origin = rng(1, 1)
        if origin(2, 1).raw_value in _empty:
            bottom_left = origin
        elif origin(3, 1).raw_value in _empty:
            bottom_left = origin(2, 1)
        else:
            bottom_left = origin(2, 1).end('down')

        if origin(1, 2).raw_value in _empty:
            top_right = origin
        elif origin(1, 3).raw_value in _empty:
            top_right = origin(1, 2)
        else:
            top_right = origin(1, 2).end('right')

        return Range(top_right, bottom_left)

    def clear(self, rng, skip, vshape):

        # calculate how many rows of existing data are present
        xdata_origin = rng(1+skip[0], 1)
        if xdata_origin.raw_value in _empty:
            xdata_rows = 0
        elif xdata_origin(2, 1).raw_value in _empty:
            xdata_rows = 1
        else:
            xdata_rows = xdata_origin.end('down').row - xdata_origin.row + 1

        # calculate row to clear till
        clear_to_row = max(
            skip[0] + xdata_rows,  # clear out existing data
            1 + vshape[0]  # clear one space after last element
        )

        # calculate how many columns of existing data are present
        xdata_origin = rng(1, 1 + skip[1])
        if xdata_origin.raw_value in _empty:
            xdata_cols = 0
        elif xdata_origin(1, 2).raw_value in _empty:
            xdata_cols = 1
        else:
            xdata_cols = xdata_origin.end('right').column - xdata_origin.column + 1

        # calculate column to clear till
        clear_to_col = max(
            skip[1] + xdata_cols,  # clear out existing data
            1 + vshape[1]  # clear one space after last element
        )

        # We have now determined the shape of the rectangle to clear out, but there are
        # two top-left rectangular regions which require special treatment:
        # - skip: in the case of a UDF this is the shape occupied by the formula (or array formula)
        # - vshape: this is the shape of the value which will be written subsequently
        # In the first case we cannot clear the cells since that would cause the formula to be deleted,
        # in the second case there is no need to clear out those cells since their values will be overwritten.
        # The second case is not as indispensable like the first, however if we ignore it we get a flicker effect
        # as the value is written out.

        shapes = [skip, vshape]
        shapes.sort()

        # check if shape 0 is contained within shape 1
        if shapes[0][1] <= shapes[1][1]:
            shapes.pop(0)

        # check if either shape is degenerate (i.e. if at least one dimension is zero)
        for i in range(len(shapes)-1, -1, -1):
            if shapes[i][0] == 0 or shapes[i][1] == 0:
                shapes.pop(i)

        # add dummy corner at bottom-left of clear region, if needed
        if not shapes or shapes[-1][0] < clear_to_row:
            shapes.append((clear_to_row, 0))

        # clear the relevant cells, one row-block at a time, starting from the appropriate column
        #
        #    .....X
        #    ...XXX
        #    XXXXXX
        #
        prev_row = 1
        for s in shapes:
            if s[0] <= clear_to_row and s[1] < clear_to_col:
                Range(rng(prev_row, s[1] + 1), rng(s[0], clear_to_col)).clear_contents()
            prev_row = s[0] + 1

TableExpander().register('table')


class VerticalExpander(Expander):

    def expand(self, rng):
        if rng(2, 1).raw_value in _empty:
            return Range(rng(1, 1), rng(1, rng.shape[1]))
        elif rng(3, 1).raw_value in _empty:
            return Range(rng(1, 1), rng(2, rng.shape[1]))
        else:
            end_row = rng(2, 1).end('down').row - rng.row + 1
            return Range(rng(1, 1), rng(end_row, rng.shape[1]))

    def clear(self, rng, skip, vshape):

        # calculate how many rows of existing data are present
        xdata_origin = rng(1+skip[0], 1)
        if xdata_origin.raw_value in _empty:
            xdata_rows = 0
        elif xdata_origin(2, 1).raw_value in _empty:
            xdata_rows = 1
        else:
            xdata_rows = xdata_origin.end('down').row - xdata_origin.row + 1

        # calculate row to clear till
        clear_to_row = max(
            skip[0] + xdata_rows,       # clear out existing data
            1 + vshape[0]               # clear one space after last element
        )

        # calculate row to start clearing from
        clear_from_row = max(
            1 + skip[0],                # cannot clear before skip region
            1 + vshape[0]               # no point clearing rows where data will be written
        )

        if clear_from_row <= clear_to_row:
            ncols = len(rng.columns)
            Range(rng(clear_from_row, 1), rng(clear_to_row, ncols)).clear_contents()


VerticalExpander().register('vertical', 'down', 'd')


class HorizontalExpander(Expander):

    def expand(self, rng):
        if rng(1, 2).raw_value in _empty:
            return Range(rng(1, 1), rng(rng.shape[0], 1))
        elif rng(1, 3).raw_value in _empty:
            return Range(rng(1, 1), rng(rng.shape[0], 2))
        else:
            end_column = rng(1, 2).end('right').column - rng.column + 1
        return Range(rng(1, 1), rng(rng.shape[0], end_column))

    def clear(self, rng, skip, vshape):

        # calculate how many columns of existing data are present
        xdata_origin = rng(1, 1 + skip[1])
        if xdata_origin.raw_value in _empty:
            xdata_cols = 0
        elif xdata_origin(1, 2).raw_value in _empty:
            xdata_cols = 1
        else:
            xdata_cols = xdata_origin.end('right').column - xdata_origin.column + 1

        # calculate column to clear till
        clear_to_col = max(
            skip[1] + xdata_cols,   # clear out existing data
            1 + vshape[1]           # clear one space after last element
        )

        # calculate column to start clearing from
        clear_from_col = max(
            1 + skip[1],    # cannot clear before skip region
            1 + vshape[1]   # no point clearing columns where data will be written
        )

        if clear_from_col <= clear_to_col:
            nrows = len(rng.rows)
            Range(rng(1, clear_from_col), rng(nrows, clear_to_col)).clear_contents()


HorizontalExpander().register('horizontal', 'right', 'r')
