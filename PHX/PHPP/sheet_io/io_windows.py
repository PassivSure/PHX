# -*- Python Version: 3.10 -*-

"""Controller Class for the PHPP Windows worksheet."""

from __future__ import annotations

from collections.abc import Generator

from ph_units.unit_type import Unit

from PHX.PHPP.phpp_localization import shape_model
from PHX.PHPP.phpp_model.windows_rows import WindowRow, get_name_from_glazing_id
from PHX.PHPP.sheet_io._stale_rows import (
    _cell_has_value,
    _clear_cells,
    _format_stale_rows,
    _read_column_with_integrity_guard,
    _typed_cells,
)
from PHX.xl import xl_app
from PHX.xl.xl_data import XlItem, col_offset, merge_xl_item_rows, xl_chr, xl_ord

# -- The inputs a stale window row is cleared of: what PHX itself writes, less the four
# -- psi-install fields, whose 1s are PHPP's own defaults on every unused row.
STALE_ROW_INPUTS = ("variant_input", "quantity", "description", "width", "height", "host", "glazing_id", "frame_id")


class Windows:
    """IO Controller Class for PHPP "Windows" worksheet."""

    def __init__(self, _xl: xl_app.XLConnection, shape: shape_model.Windows):
        self.xl = _xl
        self.shape = shape
        self._header_row: int | None = None
        self._first_entry_row: int | None = None
        self._last_entry_row: int | None = None

    @property
    def header_row(self) -> int:
        """The row number for the Window entry 'Header'."""
        if not self._header_row:
            self._header_row = self.find_header_row()
        return self._header_row

    @property
    def first_entry_row(self) -> int:
        """Return the starting row for the window data entry block."""
        if not self._first_entry_row:
            self._first_entry_row = self.find_first_entry_row()
        return self._first_entry_row

    @property
    def last_entry_row(self) -> int:
        """Return the ending row for the window data entry block."""
        if not self._last_entry_row:
            self._last_entry_row = self.find_last_entry_row()
        return self._last_entry_row

    @property
    def section_last_entry_row(self) -> int:
        """The last window-entry row: the one above PHPP's 'Unhide additional rows' note.

        'last_entry_row' is two rows past that note (the read ranges here have always run to
        it). Those rows carry PHPP's instructions in the quantity column, so the stale-row
        check must stop above them.
        """
        return self.last_entry_row - 3

    @property
    def entry_range_start(self) -> str:
        return f"{self.shape.window_rows.locator_col_entry}{self.first_entry_row}"

    @property
    def entry_range_end(self) -> str:
        return f"{self.shape.window_rows.locator_col_entry}{self.last_entry_row}"

    @property
    def entry_range(self) -> str:
        """Return the range for the window data entry block."""
        return f"{self.entry_range_start}:{self.entry_range_end}"

    @property
    def windows_row_numbers(self) -> list[int]:
        """Return a list of all the window row numbers."""
        return list(range(self.first_entry_row, self.last_entry_row + 1))

    @property
    def used_window_row_numbers(self) -> Generator[int, None, None]:
        def has_window_data(_item) -> bool:
            return not (_item == "-" or _item == "<End of designPH import!>" or _item is None)

        data = self.xl.get_single_column_data(
            self.shape.name,
            str(self.shape.window_rows.inputs.description.column),
            self.first_entry_row,
            self.last_entry_row,
        )
        return (i for i, val in enumerate(data, start=self.first_entry_row) if has_window_data(val))

    def find_header_row(self, _row_start: int = 1, _read_length: int = 100) -> int:
        """Return the row number for the Window entry section 'Header'"""

        xl_data = self.xl.get_single_column_data(
            _sheet_name=self.shape.name,
            _col=self.shape.window_rows.locator_col_header,
            _row_start=_row_start,
            _row_end=_row_start + _read_length,
        )

        for i, val in enumerate(xl_data):
            if self.shape.window_rows.locator_string_header == val:
                return i

        raise Exception(
            f"Error: Cannot find the '{self.shape.window_rows.locator_string_header}' "
            f"header on the '{self.shape.name}' sheet, column {self.shape.window_rows.locator_string_header}?"
        )

    def find_first_entry_row(self, _start_row: int = 1, _read_length: int = 100) -> int:
        """Return the starting row for the window data entry block."""

        end_row = _start_row + _read_length
        xl_data = self.xl.get_single_column_data(
            _sheet_name=self.shape.name,
            _col=self.shape.window_rows.locator_col_entry,
            _row_start=_start_row,
            _row_end=end_row,
        )

        for i, val in enumerate(xl_data, start=_start_row):
            if self.shape.window_rows.locator_string_entry in str(val):
                return i + 2

        if end_row < 10_000:
            return self.find_first_entry_row(_start_row=end_row, _read_length=1_000)

        raise Exception(
            f"Error: Cannot find the '{self.shape.window_rows.locator_string_entry}' "
            f"marker on the '{self.shape.name}' sheet, column {self.shape.window_rows.locator_col_entry}?"
        )

    def find_last_entry_row(self, _start_row: int | None = None) -> int:
        """Return the last row of the Window input section."""
        if not _start_row:
            _start_row = self.first_entry_row
        elif _start_row > 10_000:
            raise Exception(
                f"Error: Cannot find the last row in the '{self.shape.name}' sheet, column {self.shape.window_rows_end.locator_col_entry}?"
            )

        _row_end = _start_row + 500
        xl_data = self.xl.get_single_column_data(
            _sheet_name=self.shape.name,
            _col=self.shape.window_rows_end.locator_col_entry,
            _row_start=_start_row,
            _row_end=_row_end,
        )

        for i, val in enumerate(xl_data, start=_start_row):
            if self.shape.window_rows_end.locator_string_entry in str(val):
                return i + 2
        else:
            return self.find_last_entry_row(_row_end)

    def set_single_window_construction_ids(
        self,
        _row_num: int,
        _glazing_construction_id: str,
        _frame_construction_id: str,
    ) -> None:
        """Set the glazing and frame construction IDs for a single window.

        Arguments:
        ---------
            * _row_num: int
                The row number for the window to set the construction IDs for.
            * _glazing_construction_id: str
                The glazing construction ID to set.
            * _frame_construction_id: str
                The frame construction ID to set.

        Returns:
        --------
            * None
        """
        glazing_col = str(self.shape.window_rows.inputs.glazing_id.column)
        glazing_range = f"{glazing_col}{_row_num}"
        self.xl.write_xl_item(XlItem(self.shape.name, glazing_range, _glazing_construction_id))

        frame_col = str(self.shape.window_rows.inputs.frame_id.column)
        frame_range = f"{frame_col}{_row_num}"
        self.xl.write_xl_item(XlItem(self.shape.name, frame_range, _frame_construction_id))

    def write_single_window(self, _row_num: int, _window_row: WindowRow) -> None:
        """Write a single WindowRow object to the Windows worksheet."""
        for item in _window_row.create_xl_items(self.shape.name, _row_num=_row_num):
            self.xl.write_xl_item(item)

    def write_windows(self, _window_rows: list[WindowRow], *, clear_stale: bool = False) -> None:
        """Write a list of WindowRow objects to the Windows worksheet.

        Blanks a template's outer-frame lengths on the rows written: PHX writes none, and an
        inherited fixed-glazing length takes its run off the casement PHPP derives from the
        window size, while an inherited mullion fails PHPP's check against our frame. Warns if
        filled rows remain below the written block; pass clear_stale=True to blank those
        rows' inputs as well.
        """
        start = self.first_entry_row
        row_items = [
            row.create_xl_items(self.shape.name, _row_num=i) for i, row in enumerate(_window_rows, start=start)
        ]
        for item in merge_xl_item_rows(row_items):
            self.xl.write_xl_item(item)

        written_end = start + len(row_items) - 1
        stale_rows = self._stale_rows(written_end + 1, self.section_last_entry_row)
        frame_rows = list(range(start, written_end + 1))
        if clear_stale:
            frame_rows += stale_rows
        self._blank_frame_lengths(frame_rows)
        if stale_rows and clear_stale:
            for field_name in STALE_ROW_INPUTS:
                column = str(getattr(self.shape.window_rows.inputs, field_name).column)
                rows = set(_typed_cells(self.xl, self.shape.name, column, stale_rows[0], stale_rows[-1]))
                _clear_cells(self.xl, self.shape.name, column, sorted(rows.intersection(stale_rows)))

    def _stale_rows(self, _row_start: int, _row_end: int) -> list[int]:
        """Warn about the filled rows left below the written block, and return them.

        A row is stale when its quantity holds a typed value. Without one PHPP does not
        count the window; with one it does, whatever building it came from - and the
        shading writer reads its description back as one of ours.
        """
        quantity_column = str(self.shape.window_rows.inputs.quantity.column)
        stale_rows = _typed_cells(self.xl, self.shape.name, quantity_column, _row_start, _row_end)
        if stale_rows:
            self.xl.output(
                f"Warning: '{self.shape.name}' worksheet window section contains stale "
                f"{_format_stale_rows(stale_rows)}. These are leftover entries from a previous export "
                "or the template's example building, and PHPP still counts them."
            )
        return stale_rows

    @property
    def frame_length_columns(self) -> list[str]:
        """The outer-frame length columns, or none for an edition that does not map them."""
        inputs = self.shape.window_rows.inputs
        if inputs.frame_lengths_first is None or inputs.frame_lengths_last is None:
            return []
        first, last = xl_ord(str(inputs.frame_lengths_first.column)), xl_ord(str(inputs.frame_lengths_last.column))
        return [xl_chr(i) for i in range(first, last + 1)]

    def _blank_frame_lengths(self, _rows: list[int]) -> None:
        """Blank whatever a template left in the outer-frame lengths on the given rows.

        Unlike the Areas inputs, a formula here is not proof that PHPP owns the cell: the
        Example's frame lengths are formulas ('AB24 = =S24', a fixed light the full height).
        So each column is calibrated against the rows no window occupies (empty quantity),
        where whatever is present is PHPP's: a column holding a formula there is left alone
        on every row. With no unused row to calibrate on, only typed values are cleared.
        Only cells that read as filled are touched, so a blank template gets no writes.
        """
        if not _rows or not (columns := self.frame_length_columns):
            return
        first, last = self.first_entry_row, self.section_last_entry_row
        quantity_column = str(self.shape.window_rows.inputs.quantity.column)
        quantities = _read_column_with_integrity_guard(self.xl, self.shape.name, quantity_column, first, last)
        unused_rows = {row for row, value in enumerate(quantities, start=first) if not _cell_has_value(value)}
        targets = set(_rows)

        for column in columns:
            data = _read_column_with_integrity_guard(self.xl, self.shape.name, column, first, last)
            filled = {row for row, value in enumerate(data, start=first) if _cell_has_value(value)}
            if not (filled & targets):
                continue
            if unused_rows:
                calibration = sorted(filled & unused_rows)
                if calibration and self.xl.cell_has_formula(self.shape.name, f"{column}{calibration[0]}"):
                    continue  # -- PHPP's own formula: never cleared
                to_clear = filled & targets
            else:
                to_clear = {
                    row for row in filled & targets if not self.xl.cell_has_formula(self.shape.name, f"{column}{row}")
                }
            _clear_cells(self.xl, self.shape.name, column, sorted(to_clear))

    def get_all_window_names(self) -> list[str]:
        """Return a list of all the window names found in the worksheet."""

        # -- Get all the window names from the description row
        xl_data = self.xl.get_single_column_data(
            _sheet_name=self.shape.name,
            _col=self.shape.window_rows.inputs.description.column,  # type: ignore
            _row_start=self.first_entry_row,
            _row_end=self.last_entry_row,
        )

        return [str(_) for _ in xl_data if _]

    def get_total_window_area(self, _tolerance: float = 1.0) -> Unit:
        """Return the total window area from the PHPP Windows worksheet."""
        angle_data = self.xl.get_single_column_data(
            self.shape.name,
            str(self.shape.window_rows.inputs.vertical_angle.column),
            self.first_entry_row,
            self.last_entry_row,
        )
        area_data = self.xl.get_single_column_data(
            self.shape.name,
            str(self.shape.window_rows.inputs.window_area.column),
            self.first_entry_row,
            self.last_entry_row,
        )

        areas = []
        for angle, area in zip(angle_data, area_data, strict=False):
            if angle is None or area is None:
                continue
            if abs(90 - float(angle)) < _tolerance:
                areas.append(float(area))

        return Unit(sum(areas), str(self.shape.window_rows.inputs.window_area.unit))

    def get_total_skylight_area(self, _tolerance: float = 1.0) -> Unit:
        """Return the total skylight area from the PHPP Windows worksheet."""
        angle_data = self.xl.get_single_column_data(
            self.shape.name,
            str(self.shape.window_rows.inputs.vertical_angle.column),
            self.first_entry_row,
            self.last_entry_row,
        )
        area_data = self.xl.get_single_column_data(
            self.shape.name,
            str(self.shape.window_rows.inputs.window_area.column),
            self.first_entry_row,
            self.last_entry_row,
        )

        areas = []
        for angle, area in zip(angle_data, area_data, strict=False):
            if angle is None or area is None:
                continue
            if abs(90 - float(angle)) > _tolerance:
                areas.append(float(area))

        return Unit(sum(areas), str(self.shape.window_rows.inputs.window_area.unit))

    def get_all_glazing_names(self) -> set[str]:
        """Return a set of all the construction names used in the Areas worksheet."""
        glazing_data = self.xl.get_single_column_data(
            self.shape.name,
            str(self.shape.window_rows.inputs.glazing_id.column),
            self.first_entry_row,
            self.last_entry_row,
        )
        return {get_name_from_glazing_id(_d) for _d in glazing_data}

    def activate_variants(self) -> None:
        """Set the frame and glass values to link to the Variants worksheet."""

        for row_num in range(self.first_entry_row, self.last_entry_row):
            # -- Link Glazing to the variants type
            self.xl.write_xl_item(
                XlItem(
                    self.shape.name,
                    f"{self.shape.window_rows.inputs.glazing_id.column}{row_num}",
                    f"={col_offset(str(self.shape.window_rows.inputs.variant_input.column), 1)}{row_num}",
                )
            )

            # -- Link Frame to the variants type
            self.xl.write_xl_item(
                XlItem(
                    self.shape.name,
                    f"{self.shape.window_rows.inputs.frame_id.column}{row_num}",
                    f"={col_offset(str(self.shape.window_rows.inputs.variant_input.column), 2)}{row_num}",
                )
            )

    def scale_window_size(self, _row_num: int, _scale_factor: float) -> None:
        """Scale the size of a single window based on an overall scale-factor.

        This is used during baseline model generation to scale the size of the windows
        in order to set the total WWR to the specified code-required value (40%).

        Arguments:
        ----------
            * _row_num: int
                The row number of the window to scale.
            * _scale_factor: float
                The factor to scale the window by.
        """

        # -- Current dimensions of the rectangle
        # -- Get the length and width from the PHPP Windows worksheet
        width = self.xl.get_data(
            self.shape.name,
            f"{self.shape.window_rows.inputs.width.column}{_row_num}",
        )
        height = self.xl.get_data(
            self.shape.name,
            f"{self.shape.window_rows.inputs.height.column}{_row_num}",
        )

        # -- Calculate current area
        current_area = float(height) * float(width)  # type: ignore
        desired_area = current_area * _scale_factor

        # -- Calculate scaling factor
        # ** 0.5 is equivalent to taking the square root of the expression on the
        # left-hand side. This is because raising a number to the power of 1/2 is
        # the same as taking the square root of that number.
        edge_scaling_factor = (desired_area / current_area) ** 0.5

        # -- Scale the dimensions
        new_width = width * edge_scaling_factor
        new_height = height * edge_scaling_factor

        # -- Set the new dimensions
        self.xl.write_xl_item(
            XlItem(
                self.shape.name,
                f"{self.shape.window_rows.inputs.width.column}{_row_num}",
                new_width,
            )
        )
        self.xl.write_xl_item(
            XlItem(
                self.shape.name,
                f"{self.shape.window_rows.inputs.height.column}{_row_num}",
                new_height,
            )
        )

    def row_is_window(self, _row_num: int, _tolerance: float = 5.0) -> bool:
        """Return True if the row is a window, False otherwise."""
        col = str(self.shape.window_rows.inputs.vertical_angle.column)
        _range = f"{col}{_row_num}"
        v = self.xl.get_data(self.shape.name, _range)

        return abs(90.0 - float(v)) < _tolerance

    def row_is_skylight(self, _row_num: int, _tolerance: float = 5.0) -> bool:
        """Return True if the row is a skylight, False otherwise."""
        col = str(self.shape.window_rows.inputs.vertical_angle.column)
        _range = f"{col}{_row_num}"
        v = self.xl.get_data(self.shape.name, _range)

        return not abs(90.0 - float(v)) < _tolerance
