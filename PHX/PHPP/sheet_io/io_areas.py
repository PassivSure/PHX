# -*- Python Version: 3.10 -*-

"""Controller Class for the PHPP "Areas" worksheet."""

from __future__ import annotations

from collections.abc import Generator

from ph_units.unit_type import Unit

from PHX.PHPP.phpp_localization import shape_model
from PHX.PHPP.phpp_model import areas_data, areas_surface, areas_thermal_bridges
from PHX.PHPP.sheet_io._stale_rows import (
    _cell_has_value,
    _contiguous_column_groups,
    _contiguous_row_groups,
    _format_stale_rows,
    _range_address,
    _read_column_with_integrity_guard,
)
from PHX.xl import xl_app, xl_data
from PHX.xl.xl_data import col_offset


class AreasInputLocation:
    """Generic input item for Areas worksheet items."""

    def __init__(
        self,
        _xl: xl_app.XLConnection,
        _sheet_name: str,
        _search_col: str,
        _search_item: str,
        _input_row_offset: int,
    ):
        self.xl = _xl
        self.sheet_name = _sheet_name
        self.search_col = _search_col
        self.search_item = _search_item
        self.input_row_offset = _input_row_offset

    def find_input_row(self, _row_start: int = 1, _row_end: int = 200) -> int:
        """Return the row number where the search-item is found input."""
        xl_data = self.xl.get_single_column_data(
            _sheet_name=self.sheet_name,
            _col=self.search_col,
            _row_start=_row_start,
            _row_end=_row_end,
        )

        for i, val in enumerate(xl_data, start=_row_start):
            if self.search_item in str(val):
                return i + self.input_row_offset

        raise Exception(
            f'\n\tError: Not able to find the "{self.search_item}" input '
            f'section of the "{self.sheet_name}" worksheet? Please be sure '
            f'the item is note with the "{self.search_item}" flag in column {self.search_col}?'
        )


class Surfaces:
    """Reads and writes surface data to the PHPP 'Areas' worksheet."""

    def __init__(
        self,
        _xl: xl_app.XLConnection,
        _shape: shape_model.Areas,
        _group_type_exposures: dict[int, str],
    ) -> None:
        self.xl = _xl
        self.shape = _shape
        self._section_header_row: int | None = None
        self._section_first_entry_row: int | None = None
        self._section_last_entry_row: int | None = None
        self.surface_cache = {}
        self.group_type_exposures = _group_type_exposures

    @property
    def section_header_row(self) -> int:
        """Return the row number of the 'Area input' section header."""
        if not self._section_header_row:
            self._section_header_row = self.find_section_header_row()
        return self._section_header_row

    @property
    def section_first_entry_row(self) -> int:
        """Return the row number of the very first user-input entry row in the 'Area input' section."""
        if not self._section_first_entry_row:
            self._section_first_entry_row = self.find_section_first_entry_row()
        return self._section_first_entry_row

    @property
    def section_last_entry_row(self) -> int:
        """Return the row number of the last user-input entry row in the 'Area input' section."""
        if not self._section_last_entry_row:
            self._section_last_entry_row = self.find_section_last_entry_row()
        return self._section_last_entry_row

    def find_section_header_row(self, _row_start: int = 1, _row_end: int = 100) -> int:
        """Return the row number of the 'Area input' section header."""

        xl_data = self.xl.get_single_column_data(
            _sheet_name=self.shape.name,
            _col=self.shape.surface_rows.locator_col_header,
            _row_start=_row_start,
            _row_end=_row_end,
        )

        # -- The data begins at '_row_start', so enumerate from there: the
        # -- caller wants a worksheet row number, not an index into the block.
        for i, val in enumerate(xl_data, start=_row_start):
            if val == self.shape.surface_rows.locator_string_header:
                return i

        raise Exception(
            f'\n\tError: Not able to find the "Areas input" input section of '
            f'the "{self.shape.name}" worksheet? Please be sure the section begins '
            f'with the "{self.shape.surface_rows.locator_string_header}" flag in '
            f"column {self.shape.surface_rows.locator_col_header}."
        )

    def find_section_first_entry_row(self) -> int:
        """Return the row number of the very first user-input entry row in the 'Area input' section."""

        xl_data = self.xl.get_single_column_data(
            _sheet_name=self.shape.name,
            _col=self.shape.surface_rows.locator_col_entry,
            _row_start=self.section_header_row,
            _row_end=self.section_header_row + 25,
        )

        for i, val in enumerate(xl_data, start=self.section_header_row):
            try:
                val = str(int(val))  # type: ignore - Value comes in as  "1.0" from Excel?
            except:
                continue

            if val == self.shape.surface_rows.locator_string_entry:
                return i

        raise Exception('\n\tError: Not able to find the first surface entry row in the "Areas input" section?')

    def find_section_last_entry_row(self, _start_row: int | None = None) -> int:
        """Return the row number of the last user-input entry row in the 'Area input' section."""

        if not _start_row:
            _start_row = self.section_first_entry_row
        elif _start_row > 10_000:
            raise Exception('\n\tError: Not able to find the last surface entry row in the "Areas input" section?')

        _row_end = _start_row + 500
        xl_data = self.xl.get_single_column_data(
            _sheet_name=self.shape.name,
            _col=self.shape.surface_rows.locator_col_entry,
            _row_start=_start_row,
            _row_end=_row_end,
        )

        for i, val in enumerate(xl_data, start=_start_row):
            try:
                # -- See if the row has a number
                int(val)  # type: ignore
            except:
                # -- if not, that is the end of the input block
                return i - 1
        else:
            return self.find_section_last_entry_row(_row_end)

    def get_surface_phpp_id_by_name(self, _name: str, _use_cache: bool = False) -> str:
        """Return the PHPP-Style id ("1-NorthRoofSurface", ...) when given the surface name."""

        # -- Try and return the Cached value first.
        if _use_cache:
            try:
                return self.surface_cache[_name]
            except KeyError:
                pass

        row = self.xl.get_row_num_of_value_in_column(
            sheet_name=self.shape.name,
            row_start=self.section_first_entry_row,
            row_end=self.section_first_entry_row + 500,
            col=str(self.shape.surface_rows.inputs.description.column),
            find=_name,
        )

        if not row:
            raise Exception(
                f"Error: Cannot locate the phpp surface named: {_name} in"
                f"column {self.shape.surface_rows.inputs.description.column}?"
            )

        # -- Figure out the right full PHPP name, with the ID-number prefix
        prefix_range = f"{col_offset(str(self.shape.surface_rows.inputs.description.column), -1)}{row}"
        prefix_value = self.xl.get_data(self.shape.name, prefix_range)

        try:
            prefix_value = str(prefix_value)  # prefix_value comes in from excel as "1.0"
            prefix_value = float(prefix_value)  # Can't convert "1.0" -> 1 directly??
            prefix_value = int(prefix_value)
        except:
            msg = (
                f"\n\tError: Something went wrong trying to find the PHPP-ID "
                f"number for the surface: '{_name}'? Expected to find an integer "
                f"value but got: '{prefix_value}' at {self.shape.name}:{prefix_range}"
            )
            raise Exception(msg)

        self.xl.output(f"Getting PHPP Surface id for {_name}")
        name = f"{prefix_value}-{_name}"

        # -- Save in cache
        self.surface_cache[_name] = name

        return name

    def get_all_construction_names(self) -> set[str]:
        """Return a set of all the construction names used in the Areas worksheet."""
        assembly_data = self.xl.get_single_column_data(
            self.shape.name,
            str(self.shape.surface_rows.inputs.assembly_id.column),
            self.section_first_entry_row,
            self.section_last_entry_row,
        )
        return {areas_surface.get_name_from_assembly_id(_d) for _d in assembly_data}

    @property
    def all_surface_rows(
        self,
    ) -> Generator[tuple[int, areas_surface.ExistingSurfaceRow], None, None]:
        """Return a generator of all the row_nums and surface rows in the Areas worksheet.

        Yields:
        -------
            * Tuple[int, ExistingSurfaceRow]: The row-number and the surface row object.
        """

        for i in range(self.section_first_entry_row, self.section_last_entry_row):
            existing_surface_row = areas_surface.ExistingSurfaceRow(
                self.shape,
                self.xl.get_single_row_data(self.shape.name, i),
                self.group_type_exposures,
            )
            yield i, existing_surface_row


class ThermalBridges:
    """Reads and writes thermal bridge data to the PHPP 'Areas' worksheet."""

    def __init__(self, _xl: xl_app.XLConnection, _shape: shape_model.Areas) -> None:
        self.xl = _xl
        self.shape = _shape
        self._section_header_row: int | None = None
        self._section_first_entry_row: int | None = None
        self._section_last_entry_row: int | None = None

    @property
    def section_header_row(self) -> int:
        if not self._section_header_row:
            self._section_header_row = self.find_section_header_row()
        return self._section_header_row

    @property
    def section_first_entry_row(self) -> int:
        if not self._section_first_entry_row:
            self._section_first_entry_row = self.find_section_first_entry_row()
        return self._section_first_entry_row

    @property
    def section_last_entry_row(self) -> int:
        """Return the row number of the last user-input entry row in the 'Thermal Bridge input' section."""
        if not self._section_last_entry_row:
            self._section_last_entry_row = self.find_section_last_entry_row()
        return self._section_last_entry_row

    def find_section_last_entry_row(self) -> int:
        """Return the last row of the entry block: the row before the 'No.' counter stops being a number."""
        start = self.section_first_entry_row
        xl_data = self.xl.get_single_column_data(
            _sheet_name=self.shape.name,
            _col=self.shape.thermal_bridge_rows.locator_col_entry,
            _row_start=start,
            _row_end=start + 500,
        )
        for i, val in enumerate(xl_data, start=start):
            try:
                int(val)  # type: ignore
            except (TypeError, ValueError):
                return i - 1
        return start + 500

    def find_section_header_row(self, _row_start: int = 100, _row_end: int = 500) -> int:
        """Return the row number of the 'Thermal Bridge input' section header."""

        xl_data = self.xl.get_single_column_data(
            _sheet_name=self.shape.name,
            _col=self.shape.thermal_bridge_rows.locator_col_header,
            _row_start=_row_start,
            _row_end=_row_end,
        )

        for i, val in enumerate(xl_data, start=_row_start):
            if val == self.shape.thermal_bridge_rows.locator_string_header:
                return i

        # -- If isn't found in the first batch of data, recursively look further.
        if _row_end < 10_000:
            return self.find_section_header_row(
                _row_start=_row_end,
                _row_end=_row_end + 1_000,
            )

        raise Exception(
            f'\n\tError: Not able to find the "Thermal Bridge input" input section of '
            f'the "{self.shape.name}" worksheet? Please be sure the section begins '
            f'with the "{self.shape.thermal_bridge_rows.locator_string_header}" flag in '
            f"column {self.shape.thermal_bridge_rows.locator_col_header}."
        )

    def find_section_first_entry_row(self) -> int:
        """Return the row number of the very first user-input entry row in the 'Thermal Bridge input' section."""

        xl_data = self.xl.get_single_column_data(
            _sheet_name=self.shape.name,
            _col=self.shape.thermal_bridge_rows.locator_col_entry,
            _row_start=self.section_header_row,
            _row_end=self.section_header_row + 25,
        )

        for i, val in enumerate(xl_data, start=self.section_header_row):
            try:
                val = str(int(val))  # type: ignore - Value comes in as  "1.0" from Excel?
            except:
                continue

            if val == self.shape.thermal_bridge_rows.locator_string_entry:
                return i

        raise Exception(
            '\n\tError: Not able to find the first Thermal Bridge entry row in the "Thermal Bridge input" section?'
        )


class Areas:
    """IO Controller for the PHPP Areas worksheet."""

    def __init__(self, _xl: xl_app.XLConnection, _shape: shape_model.Areas) -> None:
        self.xl = _xl
        self.shape = _shape
        self.group_type_exposures = self.get_group_type_exposures()
        self.surfaces = Surfaces(self.xl, self.shape, self.group_type_exposures)
        self.thermal_bridges = ThermalBridges(self.xl, self.shape)

    def write_thermal_bridges(
        self, _tbs: list[areas_thermal_bridges.ThermalBridgeRow], *, clear_stale: bool = False
    ) -> None:
        """Write all of the the thermal bridge data to the PHPP Areas worksheet.

        Blanks a template's subtraction length on the rows written (PHPP takes it off the
        length PHX writes). Warns if filled rows remain below the written block - including
        when there are no thermal bridges to write at all; pass clear_stale=True to blank
        those rows' input columns as well.
        """
        inputs = self.shape.thermal_bridge_rows.inputs
        start = self.thermal_bridges.section_first_entry_row
        row_items = [tb.create_xl_items(self.shape.name, _row_num=i) for i, tb in enumerate(_tbs, start=start)]
        for item in xl_data.merge_xl_item_rows(row_items):
            self.xl.write_xl_item(item)
        self._blank_template_inputs(inputs, ("subtraction_length",), start, start + len(row_items) - 1)
        self._warn_or_clear_stale_rows(
            "Thermal bridge input",
            inputs,
            start + len(row_items),
            self.thermal_bridges.section_last_entry_row,
            clear_stale,
        )

    def write_surfaces(self, _surfaces: list[areas_surface.SurfaceRow], *, clear_stale: bool = False) -> None:
        """Write all of the the surface data to the PHPP Areas worksheet.

        Blanks a template's a / b dimensions and user-defined subtraction on the rows written:
        PHPP sums a row as quantity x (a x b + area - subtraction), so a leftover a x b is added
        to the area PHX writes. Warns if filled rows remain below the written block; pass
        clear_stale=True to blank those rows' input columns as well.
        """
        inputs = self.shape.surface_rows.inputs
        start = self.surfaces.section_first_entry_row
        row_items = [srf.create_xl_items(self.shape.name, _row_num=i) for i, srf in enumerate(_surfaces, start=start)]
        for item in xl_data.merge_xl_item_rows(row_items):
            self.xl.write_xl_item(item)
        self._blank_template_inputs(
            inputs, ("dimension_a", "dimension_b", "subtraction_area"), start, start + len(row_items) - 1
        )
        self._warn_or_clear_stale_rows(
            "Area input",
            inputs,
            start + len(row_items),
            self.surfaces.section_last_entry_row,
            clear_stale,
        )

    def _typed_cells(self, _column: str, _row_start: int, _row_end: int) -> list[int]:
        """Rows in the range whose cell in the column holds a typed value - not empty, not a formula.

        A formula is PHPP's own (easyPH's kitchen-exhaust thermal bridge takes its quantity
        from 'easyPH!O135'), so it is never reported or cleared. The formula check runs only
        on the cells that read as filled, which on a blank template is none of them.
        """
        if _row_start > _row_end:
            return []
        data = _read_column_with_integrity_guard(self.xl, self.shape.name, _column, _row_start, _row_end)
        return [
            row
            for row, value in enumerate(data, start=_row_start)
            if _cell_has_value(value) and not self.xl.cell_has_formula(self.shape.name, f"{_column}{row}")
        ]

    def _clear_cells(self, _column: str, _rows: list[int]) -> None:
        for row_start, row_end in _contiguous_row_groups(_rows):
            self.xl.clear_range_data(self.shape.name, _range_address(_column, _column, row_start, row_end))

    def _blank_template_inputs(
        self,
        _inputs: shape_model.AreasSurfaceInputs | shape_model.AreasThermalBridgeInputs,
        _field_names: tuple[str, ...],
        _row_start: int,
        _row_end: int,
    ) -> None:
        """Blank the given inputs on the rows PHX wrote, where a template left a typed value.

        Fields not mapped for the PHPP edition (None) are skipped.
        """
        for field_name in _field_names:
            item = getattr(_inputs, field_name)
            if item is None or item.column is None:
                continue
            self._clear_cells(item.column, self._typed_cells(item.column, _row_start, _row_end))

    def _warn_or_clear_stale_rows(
        self,
        _section_name: str,
        _inputs: shape_model.AreasSurfaceInputs | shape_model.AreasThermalBridgeInputs,
        _row_start: int,
        _row_end: int,
        _clear_stale: bool,
    ) -> None:
        """Warn about, and optionally clear, filled rows left below the written block.

        A row is stale when its quantity holds a typed value. Without one PHPP does not
        count the row; with one it does, whatever building it came from.
        """
        stale_rows = self._typed_cells(str(_inputs.quantity.column), _row_start, _row_end)
        if not stale_rows:
            return

        self.xl.output(
            f"Warning: '{self.shape.name}' worksheet {_section_name} section contains stale "
            f"{_format_stale_rows(stale_rows)}. These are leftover entries from a previous export "
            "or the template's example building, and PHPP still counts them."
        )
        if not _clear_stale:
            return

        columns = [
            item.column
            for field_name in type(_inputs).model_fields
            if (item := getattr(_inputs, field_name)) is not None and item.column is not None
        ]
        for row_start, row_end in _contiguous_row_groups(stale_rows):
            for col_start, col_end in _contiguous_column_groups(columns):
                self.xl.clear_range_data(self.shape.name, _range_address(col_start, col_end, row_start, row_end))

    def _create_input_location_object(self, _phpp_model_obj: areas_data.AreasInput) -> AreasInputLocation:
        """Create and setup the AreasInputLocation object with the correct data."""
        phpp_obj_shape: shape_model.AreasDataInput = getattr(self.shape, _phpp_model_obj.input_type)
        return AreasInputLocation(
            _xl=self.xl,
            _sheet_name=self.shape.name,
            _search_col=phpp_obj_shape.locator_col,
            _search_item=phpp_obj_shape.locator_string,
            _input_row_offset=phpp_obj_shape.input_row_offset,
        )

    def write_item(self, _phpp_model_obj: areas_data.AreasInput) -> None:
        """Write the VerificationInputItem item out to the PHPP Areas Worksheet."""
        input_object = self._create_input_location_object(_phpp_model_obj)
        input_row = input_object.find_input_row()
        xl_item = _phpp_model_obj.create_xl_item(self.shape.name, input_row)
        self.xl.write_xl_item(xl_item)

    def write_custom_group_summaries(
        self,
        _groups: set[tuple[int, str, str]],
    ) -> None:
        """Write temp zone letters and descriptions for custom groups in the summary section.

        WHY: Custom groups (12-14) are used for non-standard face-type + exposure
        combinations (e.g., floor over open air, wall adjacent to unconditioned zone).
        PHPP needs the summary row populated so it routes the group's UA to the correct
        transmission loss column (R112 for ambient "A", R113 for ground "B").

        Takes a set of (group_number, temp_zone_letter, description) tuples and writes
        them all in a single pass over the summary section, avoiding repeated Excel reads.
        """
        if not _groups:
            return

        groups_to_write = {group_num: (letter, desc) for group_num, letter, desc in _groups}

        zone_group_numbers = self.xl.get_single_column_data(
            self.shape.name,
            self.shape.summary_rows.group_number,
            1,
            30,
        )

        for i, val in enumerate(zone_group_numbers, start=1):
            try:
                num = int(val)
            except (TypeError, ValueError):
                continue

            if num not in groups_to_write:
                continue

            letter, desc = groups_to_write.pop(num)
            self.xl.write_xl_item(
                xl_data.XlItem(
                    self.shape.name,
                    f"{self.shape.summary_rows.temp_zones}{i}",
                    letter,
                )
            )
            self.xl.write_xl_item(
                xl_data.XlItem(
                    self.shape.name,
                    f"{self.shape.summary_rows.area_type}{i}",
                    desc,
                )
            )

            if not groups_to_write:
                return

    def get_group_type_exposures(self) -> dict[int, str]:
        """Return the group type exposures dictionary from the PHPP Areas worksheet."""
        zone_type_letters = self.xl.get_single_column_data(
            self.shape.name,
            self.shape.summary_rows.temp_zones,
            1,
            30,
        )
        zone_group_numbers = self.xl.get_single_column_data(
            self.shape.name,
            self.shape.summary_rows.group_number,
            1,
            30,
        )

        d = {}
        for group_num, temp_zone in zip(zone_group_numbers, zone_type_letters, strict=False):
            if not group_num or not temp_zone:
                continue

            try:
                d[int(group_num)] = temp_zone
            except TypeError:
                pass
            except ValueError:
                pass
        return d

    def get_total_net_wall_area(self) -> Unit:
        """Return the total net (apertures punched) wall area from the PHPP Areas worksheet."""
        item = self.shape.defined_ranges.exposed_wall_area
        value = self.xl.get_single_data_item(self.shape.name, f"{item.column}{item.row}")
        return Unit(float(value or 0.0), str(item.unit))

    def get_total_net_roof_area(self) -> Unit:
        """Return the total net (apertures punched) Roof area from the PHPP Areas worksheet."""
        item = self.shape.defined_ranges.roof_ceiling_area
        value = self.xl.get_single_data_item(self.shape.name, f"{item.column}{item.row}")
        return Unit(float(value or 0.0), str(item.unit))

    def get_total_vertical_window_area(self) -> Unit:
        """Return the total window area from the PHPP Areas worksheet.


        Note: this value includes any non-horizontal windows and so may include "skylights",
        if the skylights are on sloped roof surfaces. Use the
        io_windows.get_total_window_area() function to get the total window area only.
        """
        unit_type = self.shape.defined_ranges.window_area_north.unit
        start = self.shape.defined_ranges.window_area_north.xl_range
        end = self.shape.defined_ranges.window_area_west.xl_range
        _range = f"{start}:{end}"

        xl_data: list[float] = self.xl.get_data(self.shape.name, _range)  # type: ignore
        unit_data: list[Unit] = [Unit(v, str(unit_type)) for v in xl_data]

        return sum(unit_data) or Unit(0.0, str(unit_type))

    def get_total_horizontal_window_area(self) -> Unit:
        """Return the total skylight area from the PHPP Areas worksheet.

        Note: this value includes ONLY flat windows and not all the "skylights",
        if the skylights are on sloped roof surfaces. Use the
        io_windows.get_total_skylight_area() function to get the total skylight area.
        """
        item = self.shape.defined_ranges.window_area_horizontal
        value = self.xl.get_single_data_item(self.shape.name, f"{item.column}{item.row}")
        return Unit(float(value or 0.0), str(item.unit))

    def set_surface_row_construction(self, _row_num: int, _phpp_constriction_id: str) -> None:
        """Set the construction-id for the surface row in the PHPP Areas worksheet."""
        col = self.shape.surface_rows.inputs.assembly_id.column
        item = xl_data.XlItem(self.shape.name, f"{col}{_row_num}", _phpp_constriction_id)
        self.xl.write_xl_item(item)

    def set_surface_row_solar_absorptivity(self, _row_num: int, _absorptivity: float = 0.75) -> None:
        """Set the solar absorptivity for the surface row in the PHPP Areas worksheet."""
        col = self.shape.surface_rows.inputs.absorptivity.column
        item = xl_data.XlItem(self.shape.name, f"{col}{_row_num}", _absorptivity)
        self.xl.write_xl_item(item)

    def set_surface_row_emissivity(self, _row_num: int, _emissivity: float = 0.90) -> None:
        """Set the emissivity for the surface row in the PHPP Areas worksheet."""
        col = self.shape.surface_rows.inputs.emissivity.column
        item = xl_data.XlItem(self.shape.name, f"{col}{_row_num}", _emissivity)
        self.xl.write_xl_item(item)
