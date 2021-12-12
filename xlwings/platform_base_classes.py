class Range:
    @property
    def coords(self):
        raise NotImplementedError()

    @property
    def api(self):
        raise NotImplementedError()

    def __len__(self):
        raise NotImplementedError()

    @property
    def row(self):
        raise NotImplementedError()

    @property
    def column(self):
        raise NotImplementedError()

    @property
    def shape(self):
        raise NotImplementedError()

    @property
    def raw_value(self):
        raise NotImplementedError()

    @raw_value.setter
    def raw_value(self, value):
        raise NotImplementedError()

    def clear_contents(self):
        raise NotImplementedError()

    def clear(self):
        raise NotImplementedError()

    def end(self, direction):
        raise NotImplementedError()

    @property
    def formula(self):
        raise NotImplementedError()

    @formula.setter
    def formula(self, value):
        raise NotImplementedError()

    @property
    def formula2(self):
        raise NotImplementedError()

    @formula2.setter
    def formula2(self, value):
        raise NotImplementedError()

    @property
    def formula_array(self):
        raise NotImplementedError()

    @formula_array.setter
    def formula_array(self, value):
        raise NotImplementedError()

    @property
    def font(self):
        raise NotImplementedError()

    @property
    def column_width(self):
        raise NotImplementedError()

    @column_width.setter
    def column_width(self, value):
        raise NotImplementedError()

    @property
    def row_height(self):
        raise NotImplementedError()

    @row_height.setter
    def row_height(self, value):
        raise NotImplementedError()

    @property
    def width(self):
        raise NotImplementedError()

    @property
    def height(self):
        raise NotImplementedError()

    @property
    def left(self):
        raise NotImplementedError()

    @property
    def top(self):
        raise NotImplementedError()

    @property
    def has_array(self):
        raise NotImplementedError()

    @property
    def number_format(self):
        raise NotImplementedError()

    @number_format.setter
    def number_format(self, value):
        raise NotImplementedError()

    def get_address(self, row_absolute, col_absolute, external):
        raise NotImplementedError()

    @property
    def address(self):
        raise NotImplementedError()

    @property
    def current_region(self):
        raise NotImplementedError()

    def autofit(self, axis=None):
        raise NotImplementedError()

    def insert(self, shift=None, copy_origin=None):
        raise NotImplementedError()

    def delete(self, shift=None):
        raise NotImplementedError()

    def copy(self, destination=None):
        raise NotImplementedError()

    def paste(self, paste=None, operation=None, skip_blanks=False, transpose=False):
        raise NotImplementedError()

    @property
    def hyperlink(self):
        raise NotImplementedError()

    def add_hyperlink(self, address, text_to_display=None, screen_tip=None):
        raise NotImplementedError()

    @property
    def color(self):
        raise NotImplementedError()

    @color.setter
    def color(self, color_or_rgb):
        raise NotImplementedError()

    @property
    def name(self):
        raise NotImplementedError()

    @name.setter
    def name(self, value):
        raise NotImplementedError()

    def __call__(self, arg1, arg2=None):
        raise NotImplementedError()

    @property
    def rows(self):
        raise NotImplementedError()

    @property
    def columns(self):
        raise NotImplementedError()

    def select(self):
        raise NotImplementedError()

    @property
    def merge_area(self):
        raise NotImplementedError()

    @property
    def merge_cells(self):
        raise NotImplementedError()

    def merge(self, across):
        raise NotImplementedError()

    def unmerge(self):
        raise NotImplementedError()

    @property
    def table(self):
        raise NotImplementedError()

    @property
    def characters(self):
        raise NotImplementedError()

    @property
    def wrap_text(self):
        raise NotImplementedError()

    @wrap_text.setter
    def wrap_text(self, value):
        raise NotImplementedError()

    @property
    def note(self):
        raise NotImplementedError()

    def copy_picture(self, appearance, format):
        raise NotImplementedError()

    def to_png(self, path):
        raise NotImplementedError()