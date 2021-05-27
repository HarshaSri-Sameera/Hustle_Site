from gspread import utils as _gspread_utils
from . import credentials as _credentials


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
        self.row = row
        self.column = column
        self._value = value

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
        self._worksheet.update_cell(self.row, self.column, self._value)

    @property
    def col(self):
        """Implemented for gspread compatibility"""
        return self.column

    @property
    def coordinates(self):
        """Cell position in worksheet in format: A1"""
        return _gspread_utils.rowcol_to_a1(self.row, self.column)

    def __repr__(self):
        return "<Cell '%s'>" % self.coordinates


class Worksheet:
    """Worksheet class that wraps around gspread's Worksheet class.

    All original functions and attributes are still available.
    """
    def __init__(self, gspread_worksheet):
        self._worksheet = gspread_worksheet
        self._cells = []

    def __repr__(self):
        return "<Worksheet '%s'>" % self._worksheet.title

    # def __getattr__(self, item):
    #     try:
    #         return getattr(self._worksheet, item)
    #     except AttributeError:
    #         raise AttributeError("'Worksheet' object has no attribute '%s'" % item)

    @property
    def name(self):
        """Worksheet Title"""
        return self._worksheet.title

    @property
    def id(self):
        """Worksheet ID"""
        return self._worksheet.id

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

    def iter_rows(self, min_row=None, max_row=None, min_col=None, max_col=None):
        if min_row is None:
            min_row = 1
        if max_row is None:
            max_row = self.max_row
        if min_col is None:
            min_col = 1
        if max_col is None:
            max_col = self.max_column
        all_cells = []
        for cell in self.range(min_row, min_col, max_row, max_col):
            obj = Cell(self, cell.row, cell.col, cell.value)
            all_cells.append(obj)
        row_keys = sorted(set([x.row for x in all_cells]))
        cells = [sorted([x for x in all_cells if x.row == row], key=lambda x: x.row) for row in row_keys]
        return cells


class Workbook:
    def __init__(self, url, credentials):
        self.url = url
        self._client = _credentials.get_client(credentials)
        self._workbook = self._client.open_by_url(self.url)

    @property
    def worksheets(self):
        sheets = self._workbook.worksheets()
        for s in sheets:
            yield Worksheet(s)

    @property
    def active(self):
        return self._workbook.sheet1

    def __getitem__(self, item):
        if isinstance(item, int):
            sheet = self._workbook.get_worksheet(item)
            return Worksheet(sheet) if sheet is not None else None
        sheet = self._workbook.worksheet(item)
        return Worksheet(sheet) if sheet is not None else None

    def __getattr__(self, item):
        try:
            return getattr(self._workbook, item)
        except AttributeError:
            raise AttributeError("'Workbook' object has no attribute '%s'" % item)
