const ExcelJS = require('exceljs');

function num(val, decimals = 1) {
  if (val === null || val === undefined || val === '') return 0;
  const n = Number(val);
  return isNaN(n) ? 0 : Number(n.toFixed(decimals));
}

function money(val) {
  if (val === null || val === undefined || val === '') return 0;
  const n = Number(val);
  return isNaN(n) ? 0 : Math.round(n);
}

const thinBorder = {
  top: { style: 'thin', color: { argb: 'CBD5E1' } },
  left: { style: 'thin', color: { argb: 'CBD5E1' } },
  bottom: { style: 'thin', color: { argb: 'CBD5E1' } },
  right: { style: 'thin', color: { argb: 'CBD5E1' } }
};

const mediumBottomBorder = {
  top: { style: 'medium', color: { argb: '1E3A8A' } },
  left: { style: 'thin', color: { argb: 'CBD5E1' } },
  bottom: { style: 'double', color: { argb: '1E3A8A' } },
  right: { style: 'thin', color: { argb: 'CBD5E1' } }
};

/**
 * Builds utility report Excel workbook (Buffer) with Rincian, KPI, Summary, Charts, and Detail Perhitungan
 * @param {Array} rows 
 * @param {string|null} chartImageBase64 Optional PNG base64 string
 * @returns {Promise<Buffer>}
 */
async function buildUtilityWorkbook(rows, chartImageBase64 = null) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'TA RESIDENCE Utility Dashboard';
  wb.created = new Date();

  const NAVY_DARK = '1E3A8A';
  const NAVY_MED = '1D4ED8';
  const SLATE_DARK = '0F172A';
  const SLATE_TEXT = '334155';
  const SLATE_MUTED = '64748B';
  const SLATE_BG = 'F8FAFC';
  const ZEBRA_BG = 'F1F5F9';
  const AMBER_HDR = 'D97706';
  const AMBER_SUB = 'FEF3C7';
  const AMBER_CELL = 'FFFBEB';
  const CYAN_HDR = '0891B2';
  const CYAN_SUB = 'CFFAFE';
  const CYAN_CELL = 'ECFEFF';
  const TOTAL_BG = 'E2E8F0';
  const AVG_BG = 'F8FAFC';

  const first = rows[0] || {};
  const unit = first.unit || '-';
  const owner = first.owner_name || '-';
  const tower = first.tower || '-';
  const kva = first.kva;
  const mcb = first.mcb;
  let dayaLabel = kva !== undefined && kva !== null ? `${kva} KVA` : '-';
  if (mcb !== undefined && mcb !== null) {
    dayaLabel += ` (MCB ${mcb}A)`;
  }

  const periods = rows.map(r => r.period || '-');
  const periodRange = periods.length === 1 
    ? periods[0] 
    : `${periods[0]} s/d ${periods[periods.length - 1]} (${periods.length} Bulan)`;

  const now = new Date();
  const exportedAt = `${String(now.getDate()).padStart(2, '0')}-${String(now.getMonth() + 1).padStart(2, '0')}-${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

  // =========================================================================
  // SHEET 1: Rincian & Grafik
  // =========================================================================
  const ws = wb.addWorksheet('Rincian & Grafik', {
    views: [{ showGridLines: true }],
    pageSetup: {
      orientation: 'landscape',
      paperSize: 9, // A4
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0
    }
  });

  // Title Banner (Rows 1-2)
  ws.mergeCells('A1:K1');
  const a1 = ws.getCell('A1');
  a1.value = 'TA RESIDENCE — LAPORAN RINCIAN PEMAKAIAN & TAGIHAN UTILITAS';
  a1.font = { name: 'Calibri', size: 15, bold: true, color: { argb: 'FFFFFF' } };
  a1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY_DARK } };
  a1.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  ws.getRow(1).height = 30;

  ws.mergeCells('A2:K2');
  const a2 = ws.getCell('A2');
  a2.value = 'Monitoring Pemakaian Listrik (kWh) & Air (m³) — Termasuk PPJ (2,4%) & PPN (11%)';
  a2.font = { name: 'Calibri', size: 10, italic: true, color: { argb: 'DBEAFE' } };
  a2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY_MED } };
  a2.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  ws.getRow(2).height = 20;

  ws.getRow(3).height = 8;

  // Unit Info Block (Rows 4-6)
  const infoLeft = [
    ['Kode Unit', unit],
    ['Nama Pemilik', owner],
    ['Tower', `Tower ${tower}`]
  ];
  const infoRight = [
    ['Daya Listrik', dayaLabel],
    ['Periode Laporan', periodRange],
    ['Tanggal Export', exportedAt]
  ];

  for (let idx = 0; idx < 3; idx++) {
    const r = 4 + idx;
    ws.getRow(r).height = 20;

    ws.mergeCells(`A${r}:B${r}`);
    const cA = ws.getCell(`A${r}`);
    cA.value = infoLeft[idx][0];
    cA.font = { name: 'Calibri', size: 10, bold: true, color: { argb: SLATE_MUTED } };
    cA.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

    ws.mergeCells(`C${r}:E${r}`);
    const cC = ws.getCell(`C${r}`);
    cC.value = infoLeft[idx][1];
    cC.font = { name: 'Calibri', size: 10, bold: true, color: { argb: SLATE_DARK } };
    cC.alignment = { horizontal: 'left', vertical: 'middle' };

    ws.mergeCells(`G${r}:H${r}`);
    const cG = ws.getCell(`G${r}`);
    cG.value = infoRight[idx][0];
    cG.font = { name: 'Calibri', size: 10, bold: true, color: { argb: SLATE_MUTED } };
    cG.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

    ws.mergeCells(`I${r}:K${r}`);
    const cI = ws.getCell(`I${r}`);
    cI.value = infoRight[idx][1];
    cI.font = { name: 'Calibri', size: 10, bold: true, color: { argb: SLATE_DARK } };
    cI.alignment = { horizontal: 'left', vertical: 'middle' };

    for (let c = 1; c <= 11; c++) {
      const cell = ws.getRow(r).getCell(c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SLATE_BG } };
      cell.border = thinBorder;
    }
  }

  ws.getRow(7).height = 10;

  // KPI Summary Cards (Rows 8-9)
  const dataStartRow = 15;
  const dataEndRow = dataStartRow + rows.length - 1;

  const kpis = [
    { start: 1, end: 2, title: 'TOTAL LISTRIK (kWh)', formula: `SUM(D${dataStartRow}:D${dataEndRow})`, fmt: '#,##0.0" kWh"', bg: AMBER_SUB, fg: '92400E' },
    { start: 3, end: 4, title: 'TOTAL TAGIHAN LISTRIK', formula: `SUM(E${dataStartRow}:E${dataEndRow})`, fmt: '"Rp "#,##0', bg: AMBER_CELL, fg: NAVY_DARK },
    { start: 5, end: 6, title: 'TOTAL AIR (m³)', formula: `SUM(H${dataStartRow}:H${dataEndRow})`, fmt: '#,##0.0" m³"', bg: CYAN_SUB, fg: '155E75' },
    { start: 7, end: 8, title: 'TOTAL TAGIHAN AIR', formula: `SUM(I${dataStartRow}:I${dataEndRow})`, fmt: '"Rp "#,##0', bg: CYAN_CELL, fg: '0E7490' },
    { start: 9, end: 11, title: 'GRAND TOTAL (Inc. PPN 11%)', formula: `SUM(J${dataStartRow}:J${dataEndRow})`, fmt: '"Rp "#,##0', bg: SLATE_DARK, fg: 'FBBF24' }
  ];

  ws.getRow(8).height = 18;
  ws.getRow(9).height = 28;

  kpis.forEach(k => {
    ws.mergeCells(8, k.start, 8, k.end);
    ws.mergeCells(9, k.start, 9, k.end);

    const isDark = k.bg === SLATE_DARK;
    const hCell = ws.getRow(8).getCell(k.start);
    hCell.value = k.title;
    hCell.font = { name: 'Calibri', size: 8, bold: true, color: { argb: isDark ? 'CBD5E1' : SLATE_MUTED } };
    hCell.alignment = { horizontal: 'center', vertical: 'middle' };

    const vCell = ws.getRow(9).getCell(k.start);
    vCell.value = { formula: k.formula };
    vCell.font = { name: 'Calibri', size: 13, bold: true, color: { argb: k.fg } };
    vCell.numFmt = k.fmt;
    vCell.alignment = { horizontal: 'center', vertical: 'middle' };

    for (let r = 8; r <= 9; r++) {
      for (let c = k.start; c <= k.end; c++) {
        const cell = ws.getRow(r).getCell(c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: k.bg } };
        cell.border = thinBorder;
      }
    }
  });

  ws.getRow(10).height = 8;
  ws.getRow(11).height = 6;

  // Table Section Title (Row 12)
  ws.mergeCells('A12:K12');
  const a12 = ws.getCell('A12');
  a12.value = 'RINCIAN PEMAKAIAN & BIAYA PER PERIODE';
  a12.font = { name: 'Calibri', size: 11, bold: true, color: { argb: NAVY_DARK } };
  a12.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(12).height = 22;

  // Table Group Header (Row 13) & Sub-header (Row 14)
  ws.getRow(13).height = 22;
  ws.getRow(14).height = 26;

  ws.mergeCells('A13:A14');
  ws.getCell('A13').value = 'Periode (Bulan)';
  ws.mergeCells('B13:E13');
  ws.getCell('B13').value = 'LISTRIK (kWh) — Inc. PPJ 2,4%';
  ws.mergeCells('F13:I13');
  ws.getCell('F13').value = 'AIR (m³)';
  ws.getCell('J13').value = 'Inc. PPN 11%';
  ws.mergeCells('K13:K14');
  ws.getCell('K13').value = 'Catatan';

  const groupRanges = [
    { start: 1, end: 1, bg: NAVY_DARK, fg: 'FFFFFF' },
    { start: 2, end: 5, bg: AMBER_HDR, fg: 'FFFFFF' },
    { start: 6, end: 9, bg: CYAN_HDR, fg: 'FFFFFF' },
    { start: 10, end: 10, bg: NAVY_DARK, fg: 'FFFFFF' },
    { start: 11, end: 11, bg: NAVY_DARK, fg: 'FFFFFF' }
  ];

  groupRanges.forEach(g => {
    for (let c = g.start; c <= g.end; c++) {
      const cell = ws.getRow(13).getCell(c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: g.bg } };
      cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: g.fg } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = thinBorder;
    }
  });

  ws.getRow(14).getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY_DARK } };
  ws.getRow(14).getCell(1).border = thinBorder;
  ws.getRow(14).getCell(11).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY_DARK } };
  ws.getRow(14).getCell(11).border = thinBorder;

  const subHeaders = [
    { col: 2, text: 'Meter Awal', bg: AMBER_SUB, fg: '78350F' },
    { col: 3, text: 'Meter Akhir', bg: AMBER_SUB, fg: '78350F' },
    { col: 4, text: 'Listrik (kWh)', bg: AMBER_SUB, fg: '78350F' },
    { col: 5, text: 'Tagihan Listrik (Rp)', bg: AMBER_SUB, fg: '78350F' },
    { col: 6, text: 'Meter Awal', bg: CYAN_SUB, fg: '164E63' },
    { col: 7, text: 'Meter Akhir', bg: CYAN_SUB, fg: '164E63' },
    { col: 8, text: 'Air (m³)', bg: CYAN_SUB, fg: '164E63' },
    { col: 9, text: 'Tagihan Air (Rp)', bg: CYAN_SUB, fg: '164E63' },
    { col: 10, text: 'Grand Total (Rp)', bg: 'DBEAFE', fg: NAVY_DARK }
  ];

  subHeaders.forEach(sh => {
    const cell = ws.getRow(14).getCell(sh.col);
    cell.value = sh.text;
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: sh.bg } };
    cell.font = { name: 'Calibri', size: 9, bold: true, color: { argb: sh.fg } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = thinBorder;
  });

  // Populate Data Rows
  rows.forEach((rec, idx) => {
    const r = dataStartRow + idx;
    ws.getRow(r).height = 20;
    const isAlt = idx % 2 === 1;

    const rowVals = [
      rec.period || '-',
      num(rec.elec_meter_start),
      num(rec.elec_meter_end),
      num(rec.elec_usage),
      money(rec.elec_total),
      num(rec.water_meter_start),
      num(rec.water_meter_end),
      num(rec.water_usage),
      money(rec.water_total),
      money(rec.grand_total),
      rec.note || 'NORMAL'
    ];

    rowVals.forEach((val, cIdx) => {
      const col = cIdx + 1;
      const cell = ws.getRow(r).getCell(col);
      cell.value = val;
      cell.border = thinBorder;

      if (col === 1) {
        cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: SLATE_DARK } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        if (isAlt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ZEBRA_BG } };
      } else if (col >= 2 && col <= 4) {
        cell.font = { name: 'Calibri', size: 10, bold: col === 4, color: { argb: col === 4 ? SLATE_DARK : SLATE_TEXT } };
        cell.numFmt = '#,##0.0';
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isAlt ? 'FEF9E7' : AMBER_CELL } };
      } else if (col === 5) {
        cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: NAVY_MED } };
        cell.numFmt = '"Rp "#,##0';
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isAlt ? 'FEF9E7' : AMBER_CELL } };
      } else if (col >= 6 && col <= 8) {
        cell.font = { name: 'Calibri', size: 10, bold: col === 8, color: { argb: col === 8 ? SLATE_DARK : SLATE_TEXT } };
        cell.numFmt = '#,##0.0';
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isAlt ? 'E0F7FA' : CYAN_CELL } };
      } else if (col === 9) {
        cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: '0E7490' } };
        cell.numFmt = '"Rp "#,##0';
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isAlt ? 'E0F7FA' : CYAN_CELL } };
      } else if (col === 10) {
        cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: SLATE_DARK } };
        cell.numFmt = '"Rp "#,##0';
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        if (isAlt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ZEBRA_BG } };
      } else {
        cell.font = { name: 'Calibri', size: 9, color: { argb: SLATE_MUTED } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        if (isAlt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ZEBRA_BG } };
      }
    });
  });

  // TOTAL Row
  const totalRow = dataEndRow + 1;
  ws.getRow(totalRow).height = 23;
  ws.getRow(totalRow).getCell(1).value = 'TOTAL';
  ws.getRow(totalRow).getCell(2).value = '-';
  ws.getRow(totalRow).getCell(3).value = '-';
  ws.getRow(totalRow).getCell(4).value = { formula: `SUM(D${dataStartRow}:D${dataEndRow})` };
  ws.getRow(totalRow).getCell(4).numFmt = '#,##0.0';
  ws.getRow(totalRow).getCell(5).value = { formula: `SUM(E${dataStartRow}:E${dataEndRow})` };
  ws.getRow(totalRow).getCell(5).numFmt = '"Rp "#,##0';
  ws.getRow(totalRow).getCell(6).value = '-';
  ws.getRow(totalRow).getCell(7).value = '-';
  ws.getRow(totalRow).getCell(8).value = { formula: `SUM(H${dataStartRow}:H${dataEndRow})` };
  ws.getRow(totalRow).getCell(8).numFmt = '#,##0.0';
  ws.getRow(totalRow).getCell(9).value = { formula: `SUM(I${dataStartRow}:I${dataEndRow})` };
  ws.getRow(totalRow).getCell(9).numFmt = '"Rp "#,##0';
  ws.getRow(totalRow).getCell(10).value = { formula: `SUM(J${dataStartRow}:J${dataEndRow})` };
  ws.getRow(totalRow).getCell(10).numFmt = '"Rp "#,##0';
  ws.getRow(totalRow).getCell(11).value = `${rows.length} Bulan`;

  for (let c = 1; c <= 11; c++) {
    const cell = ws.getRow(totalRow).getCell(c);
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: SLATE_DARK } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TOTAL_BG } };
    cell.border = mediumBottomBorder;
    if ([4, 5, 8, 9, 10].includes(c)) {
      cell.alignment = { horizontal: 'right', vertical: 'middle' };
    } else {
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    }
  }

  // RATA-RATA Row
  const avgRow = totalRow + 1;
  ws.getRow(avgRow).height = 20;
  ws.getRow(avgRow).getCell(1).value = 'RATA-RATA';
  ws.getRow(avgRow).getCell(2).value = '-';
  ws.getRow(avgRow).getCell(3).value = '-';
  ws.getRow(avgRow).getCell(4).value = { formula: `AVERAGE(D${dataStartRow}:D${dataEndRow})` };
  ws.getRow(avgRow).getCell(4).numFmt = '#,##0.0';
  ws.getRow(avgRow).getCell(5).value = { formula: `AVERAGE(E${dataStartRow}:E${dataEndRow})` };
  ws.getRow(avgRow).getCell(5).numFmt = '"Rp "#,##0';
  ws.getRow(avgRow).getCell(6).value = '-';
  ws.getRow(avgRow).getCell(7).value = '-';
  ws.getRow(avgRow).getCell(8).value = { formula: `AVERAGE(H${dataStartRow}:H${dataEndRow})` };
  ws.getRow(avgRow).getCell(8).numFmt = '#,##0.0';
  ws.getRow(avgRow).getCell(9).value = { formula: `AVERAGE(I${dataStartRow}:I${dataEndRow})` };
  ws.getRow(avgRow).getCell(9).numFmt = '"Rp "#,##0';
  ws.getRow(avgRow).getCell(10).value = { formula: `AVERAGE(J${dataStartRow}:J${dataEndRow})` };
  ws.getRow(avgRow).getCell(10).numFmt = '"Rp "#,##0';
  ws.getRow(avgRow).getCell(11).value = 'Per Bulan';

  for (let c = 1; c <= 11; c++) {
    const cell = ws.getRow(avgRow).getCell(c);
    cell.font = { name: 'Calibri', size: 9, italic: true, bold: true, color: { argb: SLATE_MUTED } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AVG_BG } };
    cell.border = thinBorder;
    if ([4, 5, 8, 9, 10].includes(c)) {
      cell.alignment = { horizontal: 'right', vertical: 'middle' };
    } else {
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    }
  }

  // Section Title for Charts
  const chartHeaderRow = avgRow + 2;
  ws.mergeCells(`A${chartHeaderRow}:K${chartHeaderRow}`);
  const chCell = ws.getCell(`A${chartHeaderRow}`);
  chCell.value = 'GRAFIK TREN PEMAKAIAN & TAGIHAN UTILITAS';
  chCell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: NAVY_DARK } };
  chCell.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(chartHeaderRow).height = 24;

  // Embed Chart Image if provided
  if (chartImageBase64) {
    const cleanBase64 = chartImageBase64.replace(/^data:image\/\w+;base64,/, '');
    const imgId = wb.addImage({
      base64: cleanBase64,
      extension: 'png'
    });
    ws.addImage(imgId, {
      tl: { col: 0, row: chartHeaderRow },
      ext: { width: 850, height: 380 }
    });
  }

  // Column Widths for Sheet 1
  const colWidths = [14, 14, 14, 15, 20, 13, 13, 14, 18, 22, 16];
  colWidths.forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });

  // =========================================================================
  // SHEET 2: Detail Perhitungan
  // =========================================================================
  const ws2 = wb.addWorksheet('Detail Perhitungan', {
    views: [{ showGridLines: true, xSplit: 1, ySplit: 4 }],
    pageSetup: {
      orientation: 'landscape',
      paperSize: 9,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0
    }
  });

  const detailHeaders = [
    { title: 'Periode', fmt: 'text', width: 12 },
    { title: 'Tower', fmt: 'text', width: 9 },
    { title: 'Unit', fmt: 'text', width: 12 },
    { title: 'Pemilik', fmt: 'text_left', width: 26 },
    { title: 'MCB (A)', fmt: 'num1', width: 10 },
    { title: 'Daya (KVA)', fmt: 'num2', width: 12 },
    { title: 'Catatan', fmt: 'text', width: 12 },
    { title: 'Listrik Awal', fmt: 'num1', width: 14 },
    { title: 'Listrik Akhir', fmt: 'num1', width: 14 },
    { title: 'Pemakaian Listrik (kWh)', fmt: 'num1', width: 16 },
    { title: 'kWh Min', fmt: 'num1', width: 12 },
    { title: 'Jam Nyala', fmt: 'num2', width: 12 },
    { title: 'Subtotal Listrik (Rp)', fmt: 'money', width: 18 },
    { title: 'Tagihan Listrik + PPJ (Rp)', fmt: 'money', width: 20 },
    { title: 'Air Awal', fmt: 'num1', width: 12 },
    { title: 'Air Akhir', fmt: 'num1', width: 12 },
    { title: 'Pemakaian Air (m³)', fmt: 'num1', width: 14 },
    { title: 'Air Blok M1 (m³)', fmt: 'num1', width: 13 },
    { title: 'Air Blok M2 (m³)', fmt: 'num1', width: 13 },
    { title: 'Air Blok M3 (m³)', fmt: 'num1', width: 13 },
    { title: 'Biaya Air M1 (Rp)', fmt: 'money', width: 16 },
    { title: 'Biaya Air M2 (Rp)', fmt: 'money', width: 16 },
    { title: 'Biaya Air M3 (Rp)', fmt: 'money', width: 16 },
    { title: 'Biaya Tetap Air (Rp)', fmt: 'money', width: 16 },
    { title: 'Total Air (Rp)', fmt: 'money', width: 17 },
    { title: 'Grand Total Inc. PPN 11% (Rp)', fmt: 'money', width: 22 }
  ];

  ws2.mergeCells(1, 1, 1, detailHeaders.length);
  const s2Title = ws2.getCell(1, 1);
  s2Title.value = `TA RESIDENCE — DETAIL PERHITUNGAN TARIF & METER (${unit} - ${owner})`;
  s2Title.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFF' } };
  s2Title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY_DARK } };
  s2Title.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  ws2.getRow(1).height = 28;

  ws2.mergeCells(2, 1, 2, detailHeaders.length);
  const s2Sub = ws2.getCell(2, 1);
  s2Sub.value = `Periode: ${periodRange}  |  Diekspor: ${exportedAt}`;
  s2Sub.font = { name: 'Calibri', size: 9, italic: true, color: { argb: 'DBEAFE' } };
  s2Sub.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY_MED } };
  s2Sub.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  ws2.getRow(2).height = 18;

  ws2.getRow(3).height = 8;
  ws2.getRow(4).height = 30;

  detailHeaders.forEach((dh, idx) => {
    const col = idx + 1;
    const cell = ws2.getRow(4).getCell(col);
    cell.value = dh.title;
    let bg = NAVY_DARK;
    if (col >= 8 && col <= 14) bg = AMBER_HDR;
    else if (col >= 15 && col <= 25) bg = CYAN_HDR;

    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    cell.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FFFFFF' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = thinBorder;
    ws2.getColumn(col).width = dh.width;
  });

  const s2Start = 5;
  rows.forEach((rec, idx) => {
    const r = s2Start + idx;
    ws2.getRow(r).height = 20;
    const isAlt = idx % 2 === 1;

    const vals = [
      rec.period || '-',
      rec.tower || '-',
      rec.unit || '-',
      rec.owner_name || '-',
      num(rec.mcb, 1),
      num(rec.kva, 2),
      rec.note || 'NORMAL',
      num(rec.elec_meter_start, 2),
      num(rec.elec_meter_end, 2),
      num(rec.elec_usage, 2),
      num(rec.kwh_min, 2),
      num(rec.jam_nyala, 2),
      money(rec.elec_subtotal),
      money(rec.elec_total),
      num(rec.water_meter_start, 2),
      num(rec.water_meter_end, 2),
      num(rec.water_usage, 2),
      num(rec.water_m1, 2),
      num(rec.water_m2, 2),
      num(rec.water_m3, 2),
      money(rec.water_cost_m1),
      money(rec.water_cost_m2),
      money(rec.water_cost_m3),
      money(rec.water_fixed_cost),
      money(rec.water_total),
      money(rec.grand_total)
    ];

    vals.forEach((val, cIdx) => {
      const col = cIdx + 1;
      const fmtType = detailHeaders[cIdx].fmt;
      const cell = ws2.getRow(r).getCell(col);
      cell.value = val;
      cell.border = thinBorder;
      const isBold = [1, 10, 14, 17, 25, 26].includes(col);
      cell.font = { name: 'Calibri', size: 9, bold: isBold, color: { argb: isBold ? SLATE_DARK : SLATE_TEXT } };
      if (isAlt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ZEBRA_BG } };

      if (fmtType === 'text') {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (fmtType === 'text_left') {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      } else if (fmtType === 'num1') {
        cell.numFmt = '#,##0.0';
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      } else if (fmtType === 'num2') {
        cell.numFmt = '#,##0.00';
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      } else if (fmtType === 'money') {
        cell.numFmt = '"Rp "#,##0';
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      }
    });
  });

  const s2End = s2Start + rows.length - 1;
  const s2Total = s2End + 1;
  ws2.getRow(s2Total).height = 22;

  const summableCols = new Set([10, 13, 14, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]);
  for (let c = 1; c <= detailHeaders.length; c++) {
    const cell = ws2.getRow(s2Total).getCell(c);
    cell.font = { name: 'Calibri', size: 9, bold: true, color: { argb: SLATE_DARK } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TOTAL_BG } };
    cell.border = mediumBottomBorder;

    if (c === 1) {
      cell.value = 'TOTAL';
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    } else if (summableCols.has(c)) {
      const colLetter = ws2.getColumn(c).letter;
      cell.value = { formula: `SUM(${colLetter}${s2Start}:${colLetter}${s2End})` };
      const fmtType = detailHeaders[c - 1].fmt;
      cell.numFmt = fmtType === 'money' ? '"Rp "#,##0' : '#,##0.0';
      cell.alignment = { horizontal: 'right', vertical: 'middle' };
    } else {
      cell.value = '';
    }
  }

  // AutoFilter
  ws2.autoFilter = {
    from: { row: 4, column: 1 },
    to: { row: s2End, column: detailHeaders.length }
  };

  const buffer = await wb.xlsx.writeBuffer();
  return buffer;
}

module.exports = {
  buildUtilityWorkbook
};
