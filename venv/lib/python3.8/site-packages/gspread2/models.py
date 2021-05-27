from gspread import utils as _gspread_utils
from gspread.models import Worksheet as _gspread_Worksheet
from . import get_client
from . import utils as _gspread2_utils
from . import styles as _styles


__all__ = ['Cell', 'Worksheet', 'Workbook']


class Cell:
    """Cell within a worksheet.

    Enhanced replacement for gspread's Cell class.

    Attributes:
        row: Cell Row number
        column: Cell Column number
        value: Cell Value. Updating this will instantly reflect on the Google Sheet
    """
    def __init__(self, worksheet, row, column, value):
        self._worksheet = worksheet
        self._row = row
        self._column = column
        self._value = value
        self._font = None
        self._fill = None
        self._alignment = None

        self._value_modified = False
        self._font_modified = False
        self._fill_modified = False

    @property
    def font(self):
        return self._font

    @font.setter
    def font(self, value):
        assert isinstance(value, _styles.Font), 'Value must be a Font instance'
        self._font = value
        self._font_modified = True
        self._worksheet._modified_cells.add(self)

    @property
    def fill(self):
        return self._fill

    @fill.setter
    def fill(self, value):
        if isinstance(value, str):
            value = _styles.colors.Color(value)
        assert isinstance(value, _styles.colors.Color), 'Value must be a Color instance'
        self._fill = value
        self._fill_modified = True
        self._worksheet._modified_cells.add(self)

    @property
    def horizontal_align(self):
        return self._alignment

    @horizontal_align.setter
    def horizontal_align(self, value):
        assert isinstance(value, str), 'Value must be a str instance'
        if value == 'centre':
            value = 'center'
        assert value in ('left', 'right', 'center'), 'Value must be "left", "right" or "center"'
        self._alignment = value
        # TODO: Update alignment online

    def refresh(self):
        """Refresh value from worksheet.

        Returns:
            :class:`Cell` instance
        """
        return self._worksheet.cell(self.row, self.column)

    def as_formula(self):
        return self._worksheet.cell(self.row, self.column, value_render_option='FORMULA')

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, value):
        self._value = value
        self._value_modified = True
        self._worksheet._modified_cells.add(self)
        # self._worksheet.update_cell(self.row, self.column, value)

    @property
    def row(self):
        return self._row

    @property
    def column(self):
        return self._column

    @property
    def col(self):
        return self._column

    @property
    def column_letter(self):
        return _gspread2_utils.get_column_letter(self.column)

    @property
    def coordinates(self):
        return _gspread_utils.rowcol_to_a1(self.row, self.column)

    def __repr__(self):
        return "<Cell '%s'>" % self.coordinates


class Worksheet:
    """Worksheet class that wraps around gspread's Worksheet class.

    All original functions and attributes are still available.
    """
    def __init__(self, workbook, gspread_worksheet):
        self._workbook = workbook
        self._worksheet = gspread_worksheet
        self._modified_cells = set()

    def __repr__(self):
        return "<Worksheet '%s'>" % self._worksheet.title

    def __getattr__(self, item):
        try:
            return getattr(self._worksheet, item)
        except AttributeError:
            raise AttributeError("'Worksheet' object has no attribute '%s'" % item)

    @property
    def name(self):
        """Worksheet Title"""
        return self._worksheet.title

    @property
    def max_row(self):
        return self._worksheet.row_count

    @property
    def max_column(self):
        return self._worksheet.col_count

    def cell(self, *args, value_render_option='FORMATTED_VALUE'):
        if len(args) == 1:
            arg = args[0]
            assert isinstance(arg, str)
            cell = self._worksheet.acell(arg, value_render_option=value_render_option)
            return Cell(self, cell.row, cell.col, cell.value)
        assert len(args) == 2
        assert all(isinstance(x, int) for x in args)
        cell = self._worksheet.cell(*args, value_render_option=value_render_option)
        return Cell(self, cell.row, cell.col, cell.value)

    def _parse_min_max(self, min_row, max_row, min_col, max_col):
        if min_row is None:
            min_row = 1
        if max_row is None:
            max_row = self.max_row
        if min_col is None:
            min_col = 1
        if max_col is None:
            max_col = self.max_column
        # TODO: convert letters to integers
        return min_row, max_row, min_col, max_col

    def iter_rows(self, min_row=None, max_row=None, min_col=None, max_col=None):
        min_row, max_row, min_col, max_col = self._parse_min_max(min_row, max_row, min_col, max_col)
        all_cells = []
        for cell in self.range(min_row, min_col, max_row, max_col):
            obj = Cell(self, cell.row, cell.col, cell.value)
            all_cells.append(obj)
        row_keys = sorted(set([x.row for x in all_cells]))
        cells = [sorted([x for x in all_cells if x.row == row], key=lambda x: x.row) for row in row_keys]
        return cells

    def iter_cols(self, min_row=None, max_row=None, min_col=None, max_col=None):
        min_row, max_row, min_col, max_col = self._parse_min_max(min_row, max_row, min_col, max_col)
        all_cells = []
        for cell in self.range(min_row, min_col, max_row, max_col):
            obj = Cell(self, cell.row, cell.col, cell.value)
            all_cells.append(obj)
        col_keys = sorted(set([x.col for x in all_cells]))
        cells = [sorted([x for x in all_cells if x.column == col], key=lambda x: x.column) for col in col_keys]
        return cells

    def apply_border(self, start_row, end_row, start_col, end_col, border):
        assert all(isinstance(x, int) for x in (start_row, end_row, start_col, end_col))
        assert isinstance(border, _styles.borders.Border), 'border must be a Border instance'
        _styles.api.apply_border(self, start_row, end_row, start_col, end_col, border)

    def update_cells(self):
        if self._modified_cells:
            values_modified = [x for x in self._modified_cells if x._value_modified is True]
            font_modified = [x for x in self._modified_cells if x._font_modified is True]
            fill_modified = [x for x in self._modified_cells if x._fill_modified is True]
            f_modified = font_modified + fill_modified
            if values_modified:
                self._worksheet.update_cells(values_modified)
            if f_modified:
                _styles.api.apply_formatting(self, f_modified)


class Workbook:
    def __init__(self, url, credentials):
        self.url = url
        self._client = get_client(credentials)
        self._workbook = self._client.open_by_url(self.url)  # TODO: Select URL or key depending on input

    @property
    def worksheets(self):
        sheets = self._workbook.worksheets()
        for s in sheets:
            yield Worksheet(self, s)
            
    @property
    def sheetnames(self):
        return [x.title for x in self._workbook.worksheets()]

    def get_sheet_by_name(self, name):
        return self._workbook.worksheet(name)

    @property
    def active(self):
        return Worksheet(self, self._workbook.sheet1)

    def del_worksheet(self, worksheet):
        """Delete worksheet

        Args:
            worksheet: Worksheet instance (gspread or gspread2) or worksheet name
        """
        if isinstance(worksheet, Worksheet):
            return self._workbook.del_worksheet(worksheet._worksheet)
        elif isinstance(worksheet, _gspread_Worksheet):
            return self._workbook.del_worksheet(worksheet)
        else:
            sheet = self.__getitem__(worksheet)
            if sheet is None:
                raise KeyError('Worksheet not found')
            return self._workbook.del_worksheet(sheet._worksheet)

    def __getitem__(self, item):
        if isinstance(item, int):
            sheet = self._workbook.get_worksheet(item)
        else:
            sheet = self._workbook.worksheet(item)
        if sheet is not None:
            return Worksheet(self, sheet)
        raise KeyError('Worksheet {0} does not exist'.format(item))

    def __getattr__(self, item):
        try:
            return getattr(self._workbook, item)
        except AttributeError:
            raise AttributeError("'Workbook' object has no attribute '%s'" % item)
