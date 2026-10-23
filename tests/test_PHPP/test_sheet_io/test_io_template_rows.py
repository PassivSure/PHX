# -*- Python Version: 3.10 -*-

"""Template values PHPP still calculates after an export: Areas rows, U-values shares, Ground type.

PHX writes only the cells it has. Whatever a workbook already held in the rest of a row, or
in rows below the written block, PHPP still counts - and PHI's own 'PHPP_EN_V10.4a_Example'
is a common starting workbook, complete with its example building. Rows and values here
mirror that workbook.
"""

from types import SimpleNamespace

import pytest

from PHX.model.phx_site import PhxSite
from PHX.PHPP.phpp_app import PHPPConnection
from PHX.PHPP.sheet_io.io_areas import Areas
from PHX.PHPP.sheet_io.io_ground import Ground
from PHX.PHPP.sheet_io.io_u_values import UValues
from PHX.PHPP.sheet_io.io_windows import Windows
from PHX.xl import xl_data
from PHX.xl.xl_app import XLConnection
from tests.test_PHPP.test_sheet_io.conftest import load_shape
from tests.test_xl_replay.fake_xl_framework import FakeXLFramework, cell_name, parse_range

SHEET = "Areas"


def _areas_seed(**cells) -> dict:
    seed = {"K31": "Area input", "K143": "Thermal bridge input", "K51": "Aend", "K151": "TBend"}
    seed.update({f"K{row}": row - 40 for row in range(41, 51)})  # -- the 'Area no.' counter, evaluated
    seed.update({f"K{row}": row - 145 for row in range(146, 151)})
    seed.update(cells)
    return seed


def _spy_on_clears(_connection: XLConnection) -> set[str]:
    """Record every cell PHX asks to clear.

    The replay fake writes a scalar to a range's first cell only, where Excel fills the
    range (xlwings broadcasts it), so a block clear is read off the request instead.
    """
    cleared: set[str] = set()
    clear = _connection.clear_range_data

    def spy(_sheet_name: str, _range: str) -> None:
        col_1, row_1, col_2, row_2 = parse_range(_range)
        cleared.update(cell_name(c, r) for c in range(col_1, col_2 + 1) for r in range(row_1, row_2 + 1))
        clear(_sheet_name, _range)

    _connection.clear_range_data = spy
    return cleared


def _areas(_seed: dict, _shape_filename: str = "EN_10_4A.json"):
    shape = load_shape(_shape_filename)
    fake_xl = FakeXLFramework(sheet_names=[SHEET], seed={SHEET: _seed})
    outputs: list[str] = []
    connection = XLConnection(xl_framework=fake_xl, output=outputs.append)
    cleared = _spy_on_clears(connection)
    return SimpleNamespace(fake_xl=fake_xl, cleared=cleared), Areas(connection, shape.AREAS), outputs


class _Row:
    """Stands in for a SurfaceRow / ThermalBridgeRow: writes a description and a quantity."""

    def __init__(self, _name: str):
        self.name = _name

    def create_xl_items(self, _sheet_name: str, _row_num: int) -> list[xl_data.XlItem]:
        return [xl_data.XlItem(_sheet_name, f"L{_row_num}", self.name), xl_data.XlItem(_sheet_name, f"N{_row_num}", 1)]


def _written(_run) -> dict:
    return _run.fake_xl.written_state().get(SHEET, {})


# -----------------------------------------------------------------------------
# -- Areas: the rows PHX writes


def test_a_written_surface_row_loses_the_templates_a_x_b_and_subtraction():
    # -- PHPP sums the row as quantity x (a x b + area - subtraction): the Example's 7.13 x 7.475
    # -- on our floor's row made 153.2 m2 read as 206.5 m2.
    run, areas, _ = _areas(_areas_seed(P41=7.13, R41=7.475, V41=2.0, P42=11.35))

    areas.write_surfaces([_Row("Floor"), _Row("Wall")])

    written = _written(run)
    assert {"P41", "R41", "V41", "P42"} <= run.cleared
    assert written["L41"] == "Floor"


def test_a_formula_in_a_written_row_is_left_alone():
    run, areas, _ = _areas(_areas_seed(P41="=X274", R41=10.31))

    areas.write_surfaces([_Row("Floor")])

    written = _written(run)
    assert "P41" not in written
    assert "R41" in run.cleared  # -- positive control: the typed value beside it goes


def test_an_empty_template_gets_no_extra_writes():
    run, areas, outputs = _areas(_areas_seed())

    areas.write_surfaces([_Row("Floor")])

    assert set(_written(run)) == {"L41", "N41"}
    assert not any("stale" in _ for _ in outputs)


def test_an_edition_without_the_dimension_columns_mapped_is_not_blanked():
    run, areas, _ = _areas(_areas_seed(P41=7.13), "EN_10_3.json")

    areas.write_surfaces([_Row("Floor")])

    assert "P41" not in _written(run)


# -----------------------------------------------------------------------------
# -- Areas: rows below the written block


def _stale_surface_seed() -> dict:
    # -- The Example's partition wall (100.9 m2 towards a neighbour), two rows below ours.
    return _areas_seed(
        L43="Partition wall",
        M43="18-Building element towards neighbour",
        N43=1,
        P43=11.35,
        R43=8.89,
        AA43="04ud-Partition wall",
        AE43=90,
        AF43=90,
    )


def test_a_stale_surface_row_is_reported_and_kept_by_default():
    run, areas, outputs = _areas(_stale_surface_seed())

    areas.write_surfaces([_Row("Floor"), _Row("Wall")])

    assert any("Area input section contains stale row 43" in _ for _ in outputs)
    assert "L43" not in _written(run)


def test_clear_stale_blanks_a_stale_surface_rows_inputs():
    run, areas, _ = _areas(_stale_surface_seed())

    areas.write_surfaces([_Row("Floor"), _Row("Wall")], clear_stale=True)

    written = _written(run)
    assert {"L43", "M43", "N43", "P43", "R43", "AA43", "AE43", "AF43"} <= run.cleared
    assert "K43" not in written  # -- PHPP's own counter


def test_a_row_whose_quantity_is_a_formula_is_phpps_own():
    # -- easyPH's last thermal-bridge row: "Do not change the line!", quantity read from easyPH.
    seed = _areas_seed(
        L150="Do not change the line!",
        N150='=IF(LEFT(easyPH!O135,1)="3",1,0)',
        P150=1.0,
        L147="Int. wall-basement",
        M147="17-Thermal Bridges FS/BC",
        N147=1,
        P147=11.35,
    )
    run, areas, outputs = _areas(seed)

    areas.write_thermal_bridges([], clear_stale=True)

    written = _written(run)
    assert "N147" in run.cleared  # -- the Example's bridge goes...
    assert not any(a.endswith("150") for a in written)  # -- ...PHPP's own row stays
    assert any("stale row 147" in _ for _ in outputs)


def test_stale_thermal_bridges_are_cleared_when_the_project_has_none():
    # -- The Example's 11.35 m basement-ceiling bridge reached 'Ground!P19' (0.696 W/K) on a
    # -- model without a single thermal bridge.
    seed = _areas_seed(
        L146="Ext. wall-basement", M146="15-Thermal Bridges Ambient", N146=1, P146=24.85, R146=1.0, X146=0.95
    )
    run, areas, _ = _areas(seed)

    areas.write_thermal_bridges([], clear_stale=True)

    written = _written(run)
    assert {"L146", "M146", "N146", "P146", "R146", "X146"} <= run.cleared


def test_a_written_thermal_bridge_row_loses_the_templates_subtraction_length():
    run, areas, _ = _areas(_areas_seed(R146=2.0))

    areas.write_thermal_bridges([_Row("Corner")])

    assert "R146" in run.cleared


# -----------------------------------------------------------------------------
# -- U-values


def test_a_section_share_left_by_a_previous_assembly_is_cleared():
    shape = load_shape("EN_10_4A.json")
    seed = {
        "L7": "Description of building assembly",
        "L28": "Description of building assembly",
        "O42": 0.02,
        "Q42": 0.1,
    }
    fake_xl = FakeXLFramework(sheet_names=[shape.UVALUES.name], seed={shape.UVALUES.name: seed})
    u_values = UValues(XLConnection(xl_framework=fake_xl), shape.UVALUES)

    u_values.clear_section_percentages()

    written = fake_xl.written_state().get(shape.UVALUES.name, {})
    assert written == {"O42": None, "Q42": None}


# -----------------------------------------------------------------------------
# -- Ground


def _ground_connection(**cells):
    shape = load_shape("EN_10_4A.json")
    seed = {"C22": shape.GROUND.input_block.locator_string_header, **cells}
    fake_xl = FakeXLFramework(sheet_names=[shape.GROUND.name], seed={shape.GROUND.name: seed})
    outputs: list[str] = []
    connection = XLConnection(xl_framework=fake_xl, output=outputs.append)
    phpp_conn = object.__new__(PHPPConnection)
    phpp_conn.easyPh = False
    phpp_conn.xl = connection
    phpp_conn.shape = shape
    phpp_conn.version = SimpleNamespace(number_major="10", number_minor="4", language="EN")
    phpp_conn.ground = Ground(connection, shape.GROUND)
    no_foundation = SimpleNamespace(
        variants=[SimpleNamespace(name="V", site=PhxSite(), phius_cert=SimpleNamespace(ph_building_data=None))]
    )
    return fake_xl, phpp_conn, outputs, no_foundation


def test_a_selected_floor_slab_type_is_reported_without_a_foundation():
    # -- The 10.4a Example's unheated basement: C33 'x' and a 25 m perimeter.
    fake_xl, phpp_conn, outputs, project = _ground_connection(C33="x", H19=25)

    phpp_conn.write_project_ground(project)

    assert fake_xl.written_state().get("Ground", {}) == {}
    assert any("selects a floor-slab type (C33)" in _ for _ in outputs)


def test_clear_stale_unselects_it():
    fake_xl, phpp_conn, _, project = _ground_connection(C33="x", H19=25)

    phpp_conn.write_project_ground(project, clear_stale=True)

    written = fake_xl.written_state().get("Ground", {})
    assert written == {"C33": None, "H19": None}


@pytest.mark.parametrize("clear_stale", (False, True))
def test_nothing_selected_means_nothing_written(clear_stale):
    fake_xl, phpp_conn, outputs, project = _ground_connection(H19=25)

    phpp_conn.write_project_ground(project, clear_stale=clear_stale)

    assert fake_xl.written_state().get("Ground", {}) == {}
    assert not any(_.startswith("Warning") for _ in outputs)


# -----------------------------------------------------------------------------
# -- Windows: the outer-frame lengths and the rows below the written block

WINDOWS = "Windows"


def _windows_seed(**cells) -> dict:
    # -- Entry rows 23..39: 'Quan-' two rows above the first, the 'Unhide additional rows'
    # -- note one below the last. PHPP derives the left casement (W = S - AB) on every row.
    seed = {"L21": "Quan-\ntity", "L40": "Unhide additional rows for data entry"}
    seed.update({f"W{row}": f"=S{row}-AB{row}" for row in range(23, 40)})
    seed.update(cells)
    return seed


def _example_window(row: int, name: str) -> dict:
    # -- As in PHI's Example: a fixed light the full height (AB = S) and half-width
    # -- threshold, transom and bottom-fixed lengths, typed or as formulas.
    return {
        f"L{row}": 2,
        f"N{row}": name,
        f"R{row}": 2.24,
        f"S{row}": 2.55,
        f"T{row}": "1-External wall south",
        f"U{row}": "01ud-Triple-low-e",
        f"V{row}": "01ud-Wooden frame",
        f"AA{row}": f"=R{row}/2",
        f"AB{row}": f"=S{row}",
        f"AE{row}": 1.12,
        f"AM{row}": 2,
        f"AN{row}": 1,
    }


def _windows(_seed: dict, _shape_filename: str = "EN_10_4A.json"):

    shape = load_shape(_shape_filename)
    fake_xl = FakeXLFramework(sheet_names=[WINDOWS], seed={WINDOWS: _seed})
    outputs: list[str] = []
    connection = XLConnection(xl_framework=fake_xl, output=outputs.append)
    cleared = _spy_on_clears(connection)
    return SimpleNamespace(fake_xl=fake_xl, cleared=cleared), Windows(connection, shape.WINDOWS), outputs


class _WindowRow:
    """Stands in for a WindowRow: writes a quantity, a description and a size."""

    def __init__(self, _name: str):
        self.name = _name

    def create_xl_items(self, _sheet_name: str, _row_num: int) -> list[xl_data.XlItem]:
        return [
            xl_data.XlItem(_sheet_name, f"L{_row_num}", 1),
            xl_data.XlItem(_sheet_name, f"N{_row_num}", self.name),
            xl_data.XlItem(_sheet_name, f"R{_row_num}", 1.0),
            xl_data.XlItem(_sheet_name, f"S{_row_num}", 1.5),
        ]


def test_the_window_section_ends_above_phpps_note():
    _, windows, _ = _windows(_windows_seed())

    assert (windows.first_entry_row, windows.section_last_entry_row) == (23, 39)


def test_a_written_window_row_loses_the_templates_frame_lengths():
    # -- AB = S on the Example's row made our window all fixed light: no left casement at all.
    run, windows, _ = _windows(_windows_seed(**_example_window(23, "2.01 Floor-to-ceiling")))

    windows.write_windows([_WindowRow("Our window")])

    assert {"AA23", "AB23", "AE23", "AM23"} <= run.cleared
    assert "W23" not in run.cleared  # -- PHPP's derived casement stays
    assert "AN23" not in run.cleared  # -- PHX writes psi-install itself


def test_a_frame_length_formula_phpp_keeps_on_unused_rows_is_left_alone():
    seed = _windows_seed(**_example_window(23, "2.01"))
    seed.update({f"AC{row}": f"=X{row}*0" for row in range(23, 40)})  # -- PHPP's own, on every row

    run, windows, _ = _windows(seed)
    windows.write_windows([_WindowRow("Our window")])

    assert not any(cell.startswith("AC") for cell in run.cleared)
    assert "AB23" in run.cleared  # -- positive control: the Example's own formula beside it goes


def test_an_empty_window_template_gets_no_extra_writes():
    run, windows, outputs = _windows(_windows_seed())

    windows.write_windows([_WindowRow("Our window")])

    assert set(run.fake_xl.written_state().get(WINDOWS, {})) == {"L23", "N23", "R23", "S23"}
    assert not run.cleared
    assert not any("stale" in _ for _ in outputs)


def test_an_edition_without_the_frame_lengths_mapped_is_not_blanked():
    run, windows, _ = _windows(_windows_seed(**_example_window(23, "2.01")), "EN_10_3.json")

    windows.write_windows([_WindowRow("Our window")])

    assert not run.cleared


def _stale_window_seed() -> dict:
    seed = _windows_seed(**_example_window(23, "2.01"))
    seed.update(_example_window(26, "1.03 Floor-to-ceiling"))
    return seed


def test_a_stale_window_row_is_reported_and_kept_by_default():
    run, windows, outputs = _windows(_stale_window_seed())

    windows.write_windows([_WindowRow("Our window")])

    assert any("stale row 26" in _ for _ in outputs)
    assert not any(cell.endswith("26") for cell in run.cleared)


def test_a_stale_window_row_is_cleared_on_request():
    run, windows, _ = _windows(_stale_window_seed())

    windows.write_windows([_WindowRow("Our window")], clear_stale=True)

    assert {"L26", "N26", "R26", "S26", "T26", "U26", "V26", "AA26", "AB26", "AE26", "AM26"} <= run.cleared
    assert "W26" not in run.cleared
    assert "AN26" not in run.cleared  # -- PHPP's default 1 on every row
    assert windows.get_all_window_names() == ["Our window"]


def test_no_windows_at_all_still_reports_the_templates():
    run, windows, outputs = _windows(_stale_window_seed())

    windows.write_windows([], clear_stale=True)

    assert any("stale rows 23, 26" in _ for _ in outputs)
    assert {"L23", "N23", "AB23", "L26", "AB26"} <= run.cleared
    assert windows.get_all_window_names() == []


def test_a_formula_quantity_is_not_a_stale_window():
    run, windows, outputs = _windows(_windows_seed(L30="=easyPH!O12", N30="Kitchen hood"))

    windows.write_windows([_WindowRow("Our window")], clear_stale=True)

    assert not any("stale" in _ for _ in outputs)
    assert not any(cell.endswith("30") for cell in run.cleared)
