# -*- Python Version: 3.10 -*-

"""Controller Class for the PHPP "Ground" worksheet."""

from __future__ import annotations

from PHX.PHPP.phpp_localization import shape_model
from PHX.PHPP.phpp_model import ground_data
from PHX.PHPP.sheet_io.io_exceptions import FindSectionMarkerException
from PHX.xl import xl_app


class Ground:
    """IO Controller for the PHPP "Ground" worksheet."""

    def __init__(self, _xl: xl_app.XLConnection, _shape: shape_model.Ground):
        self.xl = _xl
        self.shape = _shape

    def find_section_header_row(self, _row_start: int = 1, _row_end: int = 100) -> int:
        """Return the row number of the 'Floor slab type' header, which every input row is offset from."""
        if not (block := self.shape.input_block):
            raise ground_data.GroundExportError(f"The '{self.shape.name}' worksheet layout is not mapped.")

        xl_data = self.xl.get_single_column_data(
            _sheet_name=self.shape.name,
            _col=block.locator_col_header,
            _row_start=_row_start,
            _row_end=_row_end,
        )

        # -- The data begins at '_row_start', so enumerate from there: the
        # -- caller wants a worksheet row number, not an index into the block.
        for i, val in enumerate(xl_data, start=_row_start):
            if block.locator_string_header == val:
                return i

        raise FindSectionMarkerException(block.locator_string_header, self.shape.name, block.locator_col_header)

    def write_foundation(self, _foundation_block: ground_data.GroundFoundationBlock) -> None:
        """Write one foundation into building section 1."""
        for xl_item in _foundation_block.create_xl_items(self.shape.name, self.find_section_header_row()):
            self.xl.write_xl_item(xl_item)

    def warn_or_clear_floor_slab_selection(self, _clear: bool) -> None:
        """Warn about, and optionally clear, a floor-slab type the workbook already selects.

        For a project with no foundation. PHX then writes nothing to the sheet, so a type
        left selected in building section 1 - a previous export's, or a template's example
        building (PHI's 10.4a Example selects an unheated basement with a 25 m perimeter) -
        is still calculated. Clearing the four selectors and the perimeter leaves PHPP on
        its simplified ground estimate. Does nothing when the layout is not mapped.
        """
        if not (block := self.shape.input_block):
            return

        try:
            header_row = self.find_section_header_row()
        except FindSectionMarkerException:
            return  # -- nothing PHX would have written here either
        section = block.sections[0]
        selector_offsets = [getattr(block.inputs, name).row_offset for name in ground_data.TYPE_SELECTORS]
        selector_data = self.xl.get_single_column_data(
            _sheet_name=self.shape.name,
            _col=section.selector,
            _row_start=header_row + min(selector_offsets),
            _row_end=header_row + max(selector_offsets),
        )
        if not isinstance(selector_data, list):
            selector_data = [selector_data]
        first = header_row + min(selector_offsets)
        selected = [
            f"{section.selector}{first + i}"
            for i, value in enumerate(selector_data)
            if (first + i - header_row) in selector_offsets and value not in (None, "")
        ]
        if not selected:
            return

        self.xl.output(
            f"Warning: '{self.shape.name}' worksheet selects a floor-slab type ({', '.join(selected)}) "
            "but the project has no foundation. PHPP still calculates the selected type."
        )
        if not _clear:
            return

        for address in selected + [f"{section.left}{header_row + block.inputs.perimeter.row_offset}"]:
            self.xl.clear_range_data(self.shape.name, address)
