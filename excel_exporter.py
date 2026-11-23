import io
from datetime import datetime
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def _num(val, decimals=2):
    if val is None or val == "":
        return 0.0
    try:
        return round(float(val), decimals)
    except (TypeError, ValueError):
        return 0.0


def _money(val):
    if val is None or val == "":
        return 0
    try:
        return round(float(val))
    except (TypeError, ValueError):
        return 0


def _make_border(color="CBD5E1", style="thin"):
    side = Side(border_style=style, color=color)
    return Border(left=side, right=side, top=side, bottom=side)


def _style_chart_common(chart, title, y_title, x_title="Periode", width=16.5, height=9.2):
    chart.title = title
    chart.style = 10
    chart.width = width
    chart.height = height
    chart.legend.position = "b"
    chart.x_axis.title = x_title
    chart.x_axis.axPos = "b"
    chart.x_axis.delete = False
    chart.x_axis.tickLblPos = "low"
    chart.y_axis.title = y_title
    chart.y_axis.axPos = "l"
    chart.y_axis.delete = False
    chart.y_axis.majorGridlines = ChartLines()


def build_utility_workbook(rows):
    """Build an Excel workbook containing detailed billing breakdown and native Excel charts."""
    wb = openpyxl.Workbook()

    # Palette (matching TA Residence dashboard)
    NAVY_DARK = "1E3A8A"
    NAVY_MED = "1D4ED8"
    SLATE_DARK = "0F172A"
    SLATE_TEXT = "334155"
    SLATE_MUTED = "64748B"
    SLATE_BG = "F8FAFC"
    ZEBRA_BG = "F1F5F9"
    AMBER_HDR = "D97706"
    AMBER_SUB = "FEF3C7"
    AMBER_CELL = "FFFBEB"
    CYAN_HDR = "0891B2"
    CYAN_SUB = "CFFAFE"
    CYAN_CELL = "ECFEFF"
    TOTAL_BG = "E2E8F0"
    AVG_BG = "F8FAFC"

    thin_border = _make_border("CBD5E1", "thin")
    medium_bottom = Border(
        left=Side(border_style="thin", color="CBD5E1"),
        right=Side(border_style="thin", color="CBD5E1"),
        top=Side(border_style="medium", color=NAVY_DARK),
        bottom=Side(border_style="double", color=NAVY_DARK),
    )

    first = rows[0]
    unit = first.get("unit") or "-"
    owner = first.get("owner_name") or "-"
    tower = first.get("tower") or "-"
    kva = first.get("kva")
    mcb = first.get("mcb")
    daya_label = f"{kva} KVA" if kva is not None else "-"
    if mcb is not None:
        mcb_str = int(mcb) if float(mcb).is_integer() else mcb
        daya_label = f"{daya_label} (MCB {mcb_str}A)"

    periods = [r.get("period") or "-" for r in rows]
    period_range = (
        periods[0]
        if len(periods) == 1
        else f"{periods[0]} s/d {periods[-1]} ({len(periods)} Bulan)"
    )
    exported_at = datetime.now().strftime("%d-%m-%Y %H:%M")

    # =========================================================================
    # SHEET 1: Rincian & Grafik
    # =========================================================================
    ws = wb.active
    ws.title = "Rincian & Grafik"
    ws.views.sheetView[0].showGridLines = True

    # Title Banner (Rows 1-2)
    ws.merge_cells("A1:K1")
    ws["A1"] = "TA RESIDENCE — LAPORAN RINCIAN PEMAKAIAN & TAGIHAN UTILITAS"
    ws["A1"].font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY_DARK)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:K2")
    ws["A2"] = "Monitoring Pemakaian Listrik (kWh) & Air (m³) — Termasuk PPJ (2,4%) & PPN (11%)"
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="DBEAFE")
    ws["A2"].fill = PatternFill("solid", fgColor=NAVY_MED)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 20

    ws.row_dimensions[3].height = 8

    # Unit Info Block (Rows 4-6)
    info_left = [
        ("Kode Unit", unit),
        ("Nama Pemilik", owner),
        ("Tower", f"Tower {tower}"),
    ]
    info_right = [
        ("Daya Listrik", daya_label),
        ("Periode Laporan", period_range),
        ("Tanggal Export", exported_at),
    ]
    label_font = Font(name="Calibri", size=10, bold=True, color=SLATE_MUTED)
    val_font = Font(name="Calibri", size=10, bold=True, color=SLATE_DARK)
    info_fill = PatternFill("solid", fgColor=SLATE_BG)

    for idx in range(3):
        r = 4 + idx
        ws.row_dimensions[r].height = 20

        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.cell(r, 1, info_left[idx][0]).font = label_font
        ws.cell(r, 1).alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.cell(r, 3, info_left[idx][1]).font = val_font
        ws.cell(r, 3).alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
        ws.cell(r, 7, info_right[idx][0]).font = label_font
        ws.cell(r, 7).alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=11)
        ws.cell(r, 9, info_right[idx][1]).font = val_font
        ws.cell(r, 9).alignment = Alignment(horizontal="left", vertical="center")

        for c in range(1, 12):
            cell = ws.cell(r, c)
            cell.fill = info_fill
            cell.border = thin_border

    ws.row_dimensions[7].height = 10

    # Data table coordinates (needed early for KPI formulas)
    data_start_row = 15
    data_end_row = data_start_row + len(rows) - 1

    # KPI Summary Cards (Rows 8-10)
    kpi_specs = [
        (1, 2, "TOTAL LISTRIK (kWh)", f"=SUM(D{data_start_row}:D{data_end_row})", '#,##0.0" kWh"', AMBER_SUB, "92400E"),
        (3, 4, "TOTAL TAGIHAN LISTRIK", f"=SUM(E{data_start_row}:E{data_end_row})", '"Rp "#,##0', AMBER_CELL, NAVY_DARK),
        (5, 6, "TOTAL AIR (m³)", f"=SUM(H{data_start_row}:H{data_end_row})", '#,##0.0" m³"', CYAN_SUB, "155E75"),
        (7, 8, "TOTAL TAGIHAN AIR", f"=SUM(I{data_start_row}:I{data_end_row})", '"Rp "#,##0', CYAN_CELL, "0E7490"),
        (9, 11, "GRAND TOTAL TAGIHAN (Inc. PPN 11%)", f"=SUM(J{data_start_row}:J{data_end_row})", '"Rp "#,##0', SLATE_DARK, "FBBF24"),
    ]
    ws.row_dimensions[8].height = 18
    ws.row_dimensions[9].height = 28

    for c_start, c_end, title, formula, num_fmt, bg_hex, fg_hex in kpi_specs:
        ws.merge_cells(start_row=8, start_column=c_start, end_row=8, end_column=c_end)
        ws.merge_cells(start_row=9, start_column=c_start, end_row=9, end_column=c_end)

        is_dark = bg_hex == SLATE_DARK
        hdr_cell = ws.cell(8, c_start, title)
        hdr_cell.font = Font(name="Calibri", size=8, bold=True, color="CBD5E1" if is_dark else SLATE_MUTED)
        hdr_cell.alignment = Alignment(horizontal="center", vertical="center")

        val_cell = ws.cell(9, c_start, formula)
        val_cell.font = Font(name="Calibri", size=13, bold=True, color=fg_hex)
        val_cell.number_format = num_fmt
        val_cell.alignment = Alignment(horizontal="center", vertical="center")

        fill = PatternFill("solid", fgColor=bg_hex)
        for rr in (8, 9):
            for cc in range(c_start, c_end + 1):
                cell = ws.cell(rr, cc)
                cell.fill = fill
                cell.border = thin_border

    ws.row_dimensions[10].height = 8
    ws.row_dimensions[11].height = 6

    # Table Section Title (Row 12)
    ws.merge_cells("A12:K12")
    ws["A12"] = "RINCIAN PEMAKAIAN & BIAYA PER PERIODE"
    ws["A12"].font = Font(name="Calibri", size=11, bold=True, color=NAVY_DARK)
    ws["A12"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[12].height = 22

    # Table Group Header (Row 13) & Sub-header (Row 14)
    ws.row_dimensions[13].height = 22
    ws.row_dimensions[14].height = 26

    ws.merge_cells("A13:A14")
    ws["A13"] = "Periode (Bulan)"
    ws.merge_cells("B13:E13")
    ws["B13"] = "LISTRIK (kWh) — Inc. PPJ 2,4%"
    ws.merge_cells("F13:I13")
    ws["F13"] = "AIR (m³)"
    ws["J13"] = "Inc. PPN 11%"
    ws.merge_cells("K13:K14")
    ws["K13"] = "Catatan"

    sub_headers = [
        (2, "Meter Awal", AMBER_SUB, "78350F"),
        (3, "Meter Akhir", AMBER_SUB, "78350F"),
        (4, "Listrik (kWh)", AMBER_SUB, "78350F"),
        (5, "Tagihan Listrik (Rp)", AMBER_SUB, "78350F"),
        (6, "Meter Awal", CYAN_SUB, "164E63"),
        (7, "Meter Akhir", CYAN_SUB, "164E63"),
        (8, "Air (m³)", CYAN_SUB, "164E63"),
        (9, "Tagihan Air (Rp)", CYAN_SUB, "164E63"),
        (10, "Grand Total (Rp)", "DBEAFE", NAVY_DARK),
    ]

    # Style Row 13 group headers
    group_styles = [
        (1, 1, NAVY_DARK, "FFFFFF"),
        (2, 5, AMBER_HDR, "FFFFFF"),
        (6, 9, CYAN_HDR, "FFFFFF"),
        (10, 10, NAVY_DARK, "FFFFFF"),
        (11, 11, NAVY_DARK, "FFFFFF"),
    ]
    for c_start, c_end, bg, fg in group_styles:
        for c in range(c_start, c_end + 1):
            cell = ws.cell(13, c)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.font = Font(name="Calibri", size=10, bold=True, color=fg)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

    for c in (1, 11):
        ws.cell(14, c).fill = PatternFill("solid", fgColor=NAVY_DARK)
        ws.cell(14, c).border = thin_border

    for col_idx, text, bg, fg in sub_headers:
        cell = ws.cell(14, col_idx, text)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = Font(name="Calibri", size=9, bold=True, color=fg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Populate Data Rows
    for idx, rec in enumerate(rows):
        r = data_start_row + idx
        ws.row_dimensions[r].height = 20
        is_alt = idx % 2 == 1

        row_values = [
            rec.get("period") or "-",
            _num(rec.get("elec_meter_start"), 2),
            _num(rec.get("elec_meter_end"), 2),
            _num(rec.get("elec_usage"), 2),
            _money(rec.get("elec_total")),
            _num(rec.get("water_meter_start"), 2),
            _num(rec.get("water_meter_end"), 2),
            _num(rec.get("water_usage"), 2),
            _money(rec.get("water_total")),
            _money(rec.get("grand_total")),
            rec.get("note") or "NORMAL",
        ]

        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(r, col_idx, val)
            cell.border = thin_border

            if col_idx == 1:
                cell.font = Font(name="Calibri", size=10, bold=True, color=SLATE_DARK)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if is_alt:
                    cell.fill = PatternFill("solid", fgColor=ZEBRA_BG)
            elif col_idx in (2, 3, 4):
                cell.font = Font(name="Calibri", size=10, bold=(col_idx == 4), color=SLATE_DARK if col_idx == 4 else SLATE_TEXT)
                cell.number_format = "#,##0.0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.fill = PatternFill("solid", fgColor=AMBER_CELL if not is_alt else "FEF9E7")
            elif col_idx == 5:
                cell.font = Font(name="Calibri", size=10, bold=True, color=NAVY_MED)
                cell.number_format = '"Rp "#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.fill = PatternFill("solid", fgColor=AMBER_CELL if not is_alt else "FEF9E7")
            elif col_idx in (6, 7, 8):
                cell.font = Font(name="Calibri", size=10, bold=(col_idx == 8), color=SLATE_DARK if col_idx == 8 else SLATE_TEXT)
                cell.number_format = "#,##0.0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.fill = PatternFill("solid", fgColor=CYAN_CELL if not is_alt else "E0F7FA")
            elif col_idx == 9:
                cell.font = Font(name="Calibri", size=10, bold=True, color="0E7490")
                cell.number_format = '"Rp "#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.fill = PatternFill("solid", fgColor=CYAN_CELL if not is_alt else "E0F7FA")
            elif col_idx == 10:
                cell.font = Font(name="Calibri", size=10, bold=True, color=SLATE_DARK)
                cell.number_format = '"Rp "#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if is_alt:
                    cell.fill = PatternFill("solid", fgColor=ZEBRA_BG)
            else:
                cell.font = Font(name="Calibri", size=9, color=SLATE_MUTED)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if is_alt:
                    cell.fill = PatternFill("solid", fgColor=ZEBRA_BG)

    # TOTAL Row
    total_row = data_end_row + 1
    ws.row_dimensions[total_row].height = 23
    ws.cell(total_row, 1, "TOTAL").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(total_row, 2, "-").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(total_row, 3, "-").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(total_row, 4, f"=SUM(D{data_start_row}:D{data_end_row})").number_format = "#,##0.0"
    ws.cell(total_row, 5, f"=SUM(E{data_start_row}:E{data_end_row})").number_format = '"Rp "#,##0'
    ws.cell(total_row, 6, "-").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(total_row, 7, "-").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(total_row, 8, f"=SUM(H{data_start_row}:H{data_end_row})").number_format = "#,##0.0"
    ws.cell(total_row, 9, f"=SUM(I{data_start_row}:I{data_end_row})").number_format = '"Rp "#,##0'
    ws.cell(total_row, 10, f"=SUM(J{data_start_row}:J{data_end_row})").number_format = '"Rp "#,##0'
    ws.cell(total_row, 11, f"{len(rows)} Bulan").alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, 12):
        cell = ws.cell(total_row, col_idx)
        cell.font = Font(name="Calibri", size=10, bold=True, color=SLATE_DARK)
        cell.fill = PatternFill("solid", fgColor=TOTAL_BG)
        cell.border = medium_bottom
        if col_idx in (4, 5, 8, 9, 10):
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # RATA-RATA Row
    avg_row = total_row + 1
    ws.row_dimensions[avg_row].height = 20
    ws.cell(avg_row, 1, "RATA-RATA").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(avg_row, 2, "-").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(avg_row, 3, "-").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(avg_row, 4, f"=AVERAGE(D{data_start_row}:D{data_end_row})").number_format = "#,##0.0"
    ws.cell(avg_row, 5, f"=AVERAGE(E{data_start_row}:E{data_end_row})").number_format = '"Rp "#,##0'
    ws.cell(avg_row, 6, "-").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(avg_row, 7, "-").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(avg_row, 8, f"=AVERAGE(H{data_start_row}:H{data_end_row})").number_format = "#,##0.0"
    ws.cell(avg_row, 9, f"=AVERAGE(I{data_start_row}:I{data_end_row})").number_format = '"Rp "#,##0'
    ws.cell(avg_row, 10, f"=AVERAGE(J{data_start_row}:J{data_end_row})").number_format = '"Rp "#,##0'
    ws.cell(avg_row, 11, "Per Bulan").alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, 12):
        cell = ws.cell(avg_row, col_idx)
        cell.font = Font(name="Calibri", size=9, italic=True, bold=True, color=SLATE_MUTED)
        cell.fill = PatternFill("solid", fgColor=AVG_BG)
        cell.border = thin_border
        if col_idx in (4, 5, 8, 9, 10):
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # =========================================================================
    # CHARTS SECTION (Below table on Sheet 1)
    # =========================================================================
    chart_header_row = avg_row + 2
    ws.merge_cells(start_row=chart_header_row, start_column=1, end_row=chart_header_row, end_column=11)
    ws.cell(chart_header_row, 1, "GRAFIK TREN PEMAKAIAN & TAGIHAN UTILITAS BULANAN").font = Font(
        name="Calibri", size=11, bold=True, color=NAVY_DARK
    )
    ws.cell(chart_header_row, 1).alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[chart_header_row].height = 24

    chart_row_1 = chart_header_row + 1
    chart_row_2 = chart_row_1 + 19

    cats = Reference(ws, min_col=1, min_row=data_start_row, max_row=data_end_row)
    elec_usage_ref = Reference(ws, min_col=4, min_row=14, max_row=data_end_row)
    water_usage_ref = Reference(ws, min_col=8, min_row=14, max_row=data_end_row)
    elec_cost_ref = Reference(ws, min_col=5, min_row=14, max_row=data_end_row)
    water_cost_ref = Reference(ws, min_col=9, min_row=14, max_row=data_end_row)
    grand_cost_ref = Reference(ws, min_col=10, min_row=14, max_row=data_end_row)

    # Chart 1 (Top-Left): Dual-axis Combo Chart — Listrik (kWh) Bar + Air (m³) Line
    combo_bar = BarChart()
    combo_bar.type = "col"
    combo_bar.grouping = "clustered"
    _style_chart_common(
        combo_bar,
        title=f"Tren Pemakaian Listrik & Air — {unit}",
        y_title="Listrik (kWh)",
        width=16.5,
        height=9.2,
    )
    combo_bar.add_data(elec_usage_ref, titles_from_data=True)
    combo_bar.set_categories(cats)
    combo_bar.series[0].graphicalProperties.solidFill = "F59E0B"
    combo_bar.series[0].graphicalProperties.line.solidFill = "D97706"

    combo_line = LineChart()
    combo_line.add_data(water_usage_ref, titles_from_data=True)
    combo_line.set_categories(cats)
    combo_line.x_axis.axId = 20
    combo_line.x_axis.axPos = "t"
    combo_line.x_axis.delete = True
    combo_line.x_axis.crossAx = 200
    combo_line.x_axis.crosses = "max"
    combo_line.y_axis.axId = 200
    combo_line.y_axis.axPos = "r"
    combo_line.y_axis.title = "Air (m³)"
    combo_line.y_axis.delete = False
    combo_line.y_axis.crossAx = 20
    combo_line.y_axis.crosses = "autoZero"
    combo_line.y_axis.majorGridlines = None
    combo_line.series[0].graphicalProperties.line.solidFill = "06B6D4"
    combo_line.series[0].graphicalProperties.line.width = 28575  # 2.25 pt
    combo_line.series[0].marker.symbol = "circle"
    combo_line.series[0].marker.size = 7
    combo_line.series[0].marker.graphicalProperties.solidFill = "06B6D4"
    combo_line.series[0].marker.graphicalProperties.line.solidFill = "0891B2"
    combo_line.series[0].smooth = False

    combo_bar += combo_line
    ws.add_chart(combo_bar, f"A{chart_row_1}")

    # Chart 2 (Top-Right): Tren Tagihan Utilitas Bulanan (Rp)
    cost_chart = BarChart()
    cost_chart.type = "col"
    cost_chart.grouping = "clustered"
    _style_chart_common(
        cost_chart,
        title=f"Tren Tagihan Utilitas Bulanan (Rp) — {unit}",
        y_title="Tagihan (Rp)",
        width=16.5,
        height=9.2,
    )
    cost_chart.add_data(elec_cost_ref, titles_from_data=True)
    cost_chart.add_data(water_cost_ref, titles_from_data=True)
    cost_chart.add_data(grand_cost_ref, titles_from_data=True)
    cost_chart.set_categories(cats)
    cost_chart.y_axis.number_format = "#,##0"
    cost_chart.series[0].graphicalProperties.solidFill = "F59E0B"
    cost_chart.series[0].graphicalProperties.line.solidFill = "D97706"
    cost_chart.series[1].graphicalProperties.solidFill = "06B6D4"
    cost_chart.series[1].graphicalProperties.line.solidFill = "0891B2"
    cost_chart.series[2].graphicalProperties.solidFill = NAVY_DARK
    cost_chart.series[2].graphicalProperties.line.solidFill = SLATE_DARK
    ws.add_chart(cost_chart, f"G{chart_row_1}")

    # Chart 3 (Bottom-Left): Grafik Pemakaian Listrik (kWh)
    elec_chart = BarChart()
    elec_chart.type = "col"
    _style_chart_common(
        elec_chart,
        title=f"Pemakaian Listrik Bulanan (kWh) — {unit}",
        y_title="Listrik (kWh)",
        width=16.5,
        height=8.8,
    )
    elec_chart.add_data(elec_usage_ref, titles_from_data=True)
    elec_chart.set_categories(cats)
    elec_chart.legend = None
    elec_chart.series[0].graphicalProperties.solidFill = "F59E0B"
    elec_chart.series[0].graphicalProperties.line.solidFill = "D97706"
    ws.add_chart(elec_chart, f"A{chart_row_2}")

    # Chart 4 (Bottom-Right): Grafik Pemakaian Air (m³)
    water_chart = BarChart()
    water_chart.type = "col"
    _style_chart_common(
        water_chart,
        title=f"Pemakaian Air Bulanan (m³) — {unit}",
        y_title="Air (m³)",
        width=16.5,
        height=8.8,
    )
    water_chart.add_data(water_usage_ref, titles_from_data=True)
    water_chart.set_categories(cats)
    water_chart.legend = None
    water_chart.series[0].graphicalProperties.solidFill = "06B6D4"
    water_chart.series[0].graphicalProperties.line.solidFill = "0891B2"
    ws.add_chart(water_chart, f"G{chart_row_2}")

    # Column widths for Sheet 1
    col_widths_s1 = {
        1: 14,
        2: 14,
        3: 14,
        4: 15,
        5: 20,
        6: 13,
        7: 13,
        8: 14,
        9: 18,
        10: 22,
        11: 16,
    }
    for col_idx, width in col_widths_s1.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Print / Page setup for Sheet 1
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    # =========================================================================
    # SHEET 2: Detail Perhitungan (Complete DB columns & tariff breakdown)
    # =========================================================================
    ws2 = wb.create_sheet(title="Detail Perhitungan")
    ws2.views.sheetView[0].showGridLines = True

    detail_headers = [
        ("Periode", "text", 12),
        ("Tower", "text", 9),
        ("Unit", "text", 12),
        ("Pemilik", "text_left", 26),
        ("MCB (A)", "num1", 10),
        ("Daya (KVA)", "num2", 12),
        ("Catatan", "text", 12),
        ("Listrik Awal", "num1", 14),
        ("Listrik Akhir", "num1", 14),
        ("Pemakaian Listrik (kWh)", "num1", 16),
        ("kWh Min", "num1", 12),
        ("Jam Nyala", "num2", 12),
        ("Subtotal Listrik (Rp)", "money", 18),
        ("Tagihan Listrik + PPJ (Rp)", "money", 20),
        ("Air Awal", "num1", 12),
        ("Air Akhir", "num1", 12),
        ("Pemakaian Air (m³)", "num1", 14),
        ("Air Blok M1 (m³)", "num1", 13),
        ("Air Blok M2 (m³)", "num1", 13),
        ("Air Blok M3 (m³)", "num1", 13),
        ("Biaya Air M1 (Rp)", "money", 16),
        ("Biaya Air M2 (Rp)", "money", 16),
        ("Biaya Air M3 (Rp)", "money", 16),
        ("Biaya Tetap Air (Rp)", "money", 16),
        ("Total Air (Rp)", "money", 17),
        ("Grand Total Inc. PPN 11% (Rp)", "money", 22),
    ]

    last_col_letter = get_column_letter(len(detail_headers))
    ws2.merge_cells(f"A1:{last_col_letter}1")
    ws2["A1"] = f"TA RESIDENCE — DETAIL PERHITUNGAN TARIF & METER ({unit} - {owner})"
    ws2["A1"].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    ws2["A1"].fill = PatternFill("solid", fgColor=NAVY_DARK)
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws2.row_dimensions[1].height = 28

    ws2.merge_cells(f"A2:{last_col_letter}2")
    ws2["A2"] = f"Periode: {period_range}  |  Diekspor: {exported_at}"
    ws2["A2"].font = Font(name="Calibri", size=9, italic=True, color="DBEAFE")
    ws2["A2"].fill = PatternFill("solid", fgColor=NAVY_MED)
    ws2["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws2.row_dimensions[2].height = 18

    ws2.row_dimensions[3].height = 8
    ws2.row_dimensions[4].height = 30

    for col_idx, (title, _fmt, width) in enumerate(detail_headers, start=1):
        cell = ws2.cell(4, col_idx, title)
        if 8 <= col_idx <= 14:
            bg = AMBER_HDR
        elif 15 <= col_idx <= 25:
            bg = CYAN_HDR
        else:
            bg = NAVY_DARK
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws2.column_dimensions[get_column_letter(col_idx)].width = width

    s2_start = 5
    for idx, rec in enumerate(rows):
        r = s2_start + idx
        ws2.row_dimensions[r].height = 20
        is_alt = idx % 2 == 1

        vals = [
            rec.get("period") or "-",
            rec.get("tower") or "-",
            rec.get("unit") or "-",
            rec.get("owner_name") or "-",
            _num(rec.get("mcb"), 1),
            _num(rec.get("kva"), 2),
            rec.get("note") or "NORMAL",
            _num(rec.get("elec_meter_start"), 2),
            _num(rec.get("elec_meter_end"), 2),
            _num(rec.get("elec_usage"), 2),
            _num(rec.get("kwh_min"), 2),
            _num(rec.get("jam_nyala"), 2),
            _money(rec.get("elec_subtotal")),
            _money(rec.get("elec_total")),
            _num(rec.get("water_meter_start"), 2),
            _num(rec.get("water_meter_end"), 2),
            _num(rec.get("water_usage"), 2),
            _num(rec.get("water_m1"), 2),
            _num(rec.get("water_m2"), 2),
            _num(rec.get("water_m3"), 2),
            _money(rec.get("water_cost_m1")),
            _money(rec.get("water_cost_m2")),
            _money(rec.get("water_cost_m3")),
            _money(rec.get("water_fixed_cost")),
            _money(rec.get("water_total")),
            _money(rec.get("grand_total")),
        ]

        for col_idx, val in enumerate(vals, start=1):
            fmt_type = detail_headers[col_idx - 1][1]
            cell = ws2.cell(r, col_idx, val)
            cell.border = thin_border
            cell.font = Font(
                name="Calibri",
                size=9,
                bold=(col_idx in (1, 10, 14, 17, 25, 26)),
                color=SLATE_DARK if col_idx in (1, 10, 14, 17, 25, 26) else SLATE_TEXT,
            )
            if is_alt:
                cell.fill = PatternFill("solid", fgColor=ZEBRA_BG)

            if fmt_type == "text":
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif fmt_type == "text_left":
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif fmt_type == "num1":
                cell.number_format = "#,##0.0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif fmt_type == "num2":
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif fmt_type == "money":
                cell.number_format = '"Rp "#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")

    s2_end = s2_start + len(rows) - 1
    s2_total = s2_end + 1
    ws2.row_dimensions[s2_total].height = 22
    summable_cols = {10, 13, 14, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26}

    for col_idx in range(1, len(detail_headers) + 1):
        cell = ws2.cell(s2_total, col_idx)
        cell.font = Font(name="Calibri", size=9, bold=True, color=SLATE_DARK)
        cell.fill = PatternFill("solid", fgColor=TOTAL_BG)
        cell.border = medium_bottom
        if col_idx == 1:
            cell.value = "TOTAL"
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx in summable_cols:
            col_ltr = get_column_letter(col_idx)
            cell.value = f"=SUM({col_ltr}{s2_start}:{col_ltr}{s2_end})"
            fmt_type = detail_headers[col_idx - 1][1]
            cell.number_format = '"Rp "#,##0' if fmt_type == "money" else "#,##0.0"
            cell.alignment = Alignment(horizontal="right", vertical="center")
        else:
            cell.value = ""

    ws2.freeze_panes = "B5"
    ws2.auto_filter.ref = f"A4:{last_col_letter}{s2_end}"
    ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE
    ws2.page_setup.paperSize = ws2.PAPERSIZE_A4
    ws2.sheet_properties.pageSetUpPr.fitToPage = True
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 0

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
