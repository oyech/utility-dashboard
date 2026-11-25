const express = require('express');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const ExcelJS = require('exceljs');
const { getDb } = require('./db');
const { buildUtilityWorkbook } = require('./excel-exporter');

const app = express();

// Middleware
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use('/static', express.static(path.join(__dirname, 'static')));

// Upload configuration using in-memory storage (serverless read-only filesystem compatible)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 }
});

// ----------------------------------------------------------------------------
// Frontend Route (Modern Angular dist or fallback)
// ----------------------------------------------------------------------------
const distBrowserPath = path.join(__dirname, 'dist', 'utility-dashboard', 'browser');

if (fs.existsSync(distBrowserPath)) {
  app.use(express.static(distBrowserPath));
  app.get('/', (req, res) => {
    res.sendFile(path.join(distBrowserPath, 'index.html'));
  });
} else if (fs.existsSync(path.join(__dirname, 'templates', 'index.html'))) {
  app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'templates', 'index.html'));
  });
}

// ----------------------------------------------------------------------------
// API Endpoints
// ----------------------------------------------------------------------------

// 1. Get units with optional search query & tower filter
app.get('/api/units', async (req, res) => {
  try {
    const db = await getDb();
    const search = (req.query.q || '').trim();
    const tower = (req.query.tower || '').trim();

    let query = 'SELECT DISTINCT unit, tower, owner_name, kva FROM billing_data WHERE 1=1';
    const params = [];

    if (tower) {
      query += ' AND tower = ?';
      params.push(tower);
    }
    if (search) {
      query += ' AND (unit LIKE ? OR owner_name LIKE ?)';
      params.push(`%${search}%`, `%${search}%`);
    }

    query += ' ORDER BY unit LIMIT 50';
    const rows = db.all(query, params);
    res.json(rows);
  } catch (err) {
    console.error('Error /api/units:', err);
    res.status(500).json({ error: err.message });
  }
});

// 2. Get distinct periods
app.get('/api/periods', async (req, res) => {
  try {
    const db = await getDb();
    const rows = db.all('SELECT DISTINCT period FROM billing_data ORDER BY period DESC');
    const periods = rows.map(r => r.period);
    res.json(periods);
  } catch (err) {
    console.error('Error /api/periods:', err);
    res.status(500).json({ error: err.message });
  }
});

// 3. Get distinct towers
app.get('/api/towers', async (req, res) => {
  try {
    const db = await getDb();
    const rows = db.all("SELECT DISTINCT tower FROM billing_data WHERE tower IS NOT NULL AND tower != '' ORDER BY tower");
    const towers = rows.map(r => r.tower);
    res.json(towers);
  } catch (err) {
    console.error('Error /api/towers:', err);
    res.status(500).json({ error: err.message });
  }
});

// Helper to fetch usage records
async function fetchUnitUsage(unit, periods) {
  const db = await getDb();
  let periodList = [];
  if (Array.isArray(periods)) {
    periodList = periods.filter(p => p && p.trim());
  } else if (typeof periods === 'string' && periods.trim()) {
    periodList = [periods.trim()];
  }

  let query = 'SELECT * FROM billing_data WHERE unit = ?';
  const params = [unit];

  if (periodList.length > 0) {
    const placeholders = periodList.map(() => '?').join(',');
    query += ` AND period IN (${placeholders})`;
    params.push(...periodList);
  }

  query += ' ORDER BY period ASC';
  return db.all(query, params);
}

// 4. Get usage for a unit across periods
app.get('/api/usage', async (req, res) => {
  try {
    const unit = (req.query.unit || '').trim();
    if (!unit) {
      return res.status(400).json({ error: 'Unit is required' });
    }

    const rows = await fetchUnitUsage(unit, req.query.periods);
    res.json(rows);
  } catch (err) {
    console.error('Error /api/usage:', err);
    res.status(500).json({ error: err.message });
  }
});

// 5. Export to Excel (GET or POST)
async function handleExportExcel(req, res) {
  try {
    const q = req.query || {};
    const b = req.body || {};
    const unit = (q.unit || b.unit || '').trim();
    const periods = q.periods || b.periods;
    const chartImage = b.chartImage || null;

    if (!unit) {
      return res.status(400).json({ error: 'Unit is required' });
    }

    const rows = await fetchUnitUsage(unit, periods);
    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: 'Data tidak ditemukan untuk unit dan periode yang dipilih' });
    }

    const buffer = await buildUtilityWorkbook(rows, chartImage);

    const safeUnit = unit.replace(/[^A-Za-z0-9_-]+/g, '_').replace(/^_|_$/g, '') || 'UNIT';
    const rowPeriods = rows.map(r => r.period).filter(Boolean);
    let periodSlug = 'all';
    if (rowPeriods.length === 1) {
      periodSlug = rowPeriods[0].replace(/[^A-Za-z0-9_-]+/g, '_');
    } else if (rowPeriods.length > 1) {
      const pStart = rowPeriods[0].replace(/[^A-Za-z0-9_-]+/g, '_');
      const pEnd = rowPeriods[rowPeriods.length - 1].replace(/[^A-Za-z0-9_-]+/g, '_');
      periodSlug = `${pStart}_sd_${pEnd}`;
    }

    const filename = `Laporan_Utilitas_${safeUnit}_${periodSlug}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', buffer.length);
    res.send(buffer);
  } catch (err) {
    console.error('Error /api/export-excel:', err);
    res.status(500).json({ error: err.message });
  }
}

app.get('/api/export-excel', handleExportExcel);
app.post('/api/export-excel', handleExportExcel);

// Normalize and guard electricity minimum charge calculation during Excel import
function buildNormalizedBillingRow(period, unitVal, extractVal) {
  const no = extractVal(1);
  const tower = extractVal(2);
  const unit = String(unitVal).trim();
  const ownerName = extractVal(4);
  const mcb = extractVal(5);
  const kvaRaw = extractVal(6);
  let note = extractVal(7);

  const kva = Number(kvaRaw);
  // Skip footer/summary rows that lack a unit note or valid KVA rating
  if (!note && (!Number.isFinite(kva) || kva <= 0)) {
    return null;
  }

  const elecStart = extractVal(8);
  const elecEnd = extractVal(9);
  let elecUsage = extractVal(10);
  let kwhMin = extractVal(11);
  const jamNyala = extractVal(12);
  let elecSubtotal = extractVal(13);
  let elecTotal = extractVal(14);

  const waterStart = extractVal(15);
  const waterEnd = extractVal(16);
  const waterUsage = extractVal(17);
  const waterM1 = extractVal(18);
  const waterM2 = extractVal(19);
  const waterM3 = extractVal(20);
  const waterCostM1 = extractVal(21);
  const waterCostM2 = extractVal(22);
  const waterCostM3 = extractVal(23);
  const waterFixedCost = extractVal(24);
  const waterTotal = extractVal(25);
  let grandTotal = extractVal(26);

  if (Number.isFinite(kva) && kva > 0) {
    const expectedKwhMin = Math.round(40 * kva * 10000) / 10000;
    const numericUsage = Number(elecUsage);
    const kwhMinNum = Number(kwhMin);

    if (!Number.isFinite(kwhMinNum) || Math.abs(kwhMinNum - expectedKwhMin) > 0.05) {
      kwhMin = expectedKwhMin;
    }

    if (Number.isFinite(numericUsage)) {
      const isMinCharge = numericUsage < expectedKwhMin;
      note = isMinCharge ? 'MIN CHARGE' : 'NORMAL';
      const tariff = kva <= 2.2 ? 1444.7 : 1699.53;
      const billableKwh = isMinCharge ? expectedKwhMin : numericUsage;
      const expectedSubtotal = billableKwh * tariff;
      const expectedTotal = expectedSubtotal * 1.024;
      const subtotalNum = Number(elecSubtotal);
      const totalNum = Number(elecTotal);

      if (
        !Number.isFinite(subtotalNum) ||
        Math.abs(subtotalNum - expectedSubtotal) >= 1 ||
        !Number.isFinite(totalNum) ||
        Math.abs(totalNum - expectedTotal) >= 1
      ) {
        elecSubtotal = expectedSubtotal;
        elecTotal = expectedTotal;
        grandTotal = (elecTotal + (Number(waterTotal) || 0)) * 1.11;
      }
    }
  }

  return [
    period,
    no,
    tower,
    unit,
    ownerName,
    mcb,
    kvaRaw,
    note,
    elecStart,
    elecEnd,
    elecUsage,
    kwhMin,
    jamNyala,
    elecSubtotal,
    elecTotal,
    waterStart,
    waterEnd,
    waterUsage,
    waterM1,
    waterM2,
    waterM3,
    waterCostM1,
    waterCostM2,
    waterCostM3,
    waterFixedCost,
    waterTotal,
    grandTotal
  ];
}

// 6. Upload Excel file and import data
app.post('/api/upload', upload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }
  const period = (req.body.period || '').trim();
  if (!period) {
    return res.status(400).json({ error: 'Period is required (e.g. 2026-11)' });
  }
  if (!req.file.originalname.endsWith('.xlsx')) {
    return res.status(400).json({ error: 'Only .xlsx files are supported' });
  }

  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);
    const ws = wb.getWorksheet('ME4a') || wb.worksheets[0];

    const rowsToInsert = [];
    ws.eachRow((row, rowNumber) => {
      if (rowNumber < 7) return;
      const unitVal = row.getCell(3).value;
      if (!unitVal) return;

      const extractVal = (cellIdx) => {
        const val = row.getCell(cellIdx).value;
        if (val && typeof val === 'object' && val.result !== undefined) {
          return val.result;
        }
        return val !== undefined && val !== null ? val : null;
      };

      const normalizedRow = buildNormalizedBillingRow(period, unitVal, extractVal);
      if (normalizedRow) {
        rowsToInsert.push(normalizedRow);
      }
    });

    const db = await getDb();
    db.exec('BEGIN TRANSACTION');
    db.run('DELETE FROM billing_data WHERE period = ?', [period]);

    const insertSql = `
      INSERT INTO billing_data (
        period, no, tower, unit, owner_name, mcb, kva, note,
        elec_meter_start, elec_meter_end, elec_usage, kwh_min, jam_nyala, elec_subtotal, elec_total,
        water_meter_start, water_meter_end, water_usage, water_m1, water_m2, water_m3,
        water_cost_m1, water_cost_m2, water_cost_m3, water_fixed_cost, water_total, grand_total
      ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    `;

    for (const r of rowsToInsert) {
      db.run(insertSql, r);
    }
    db.exec('COMMIT');

    res.json({ message: `Berhasil mengimpor ${rowsToInsert.length} unit untuk periode ${period}!` });
  } catch (err) {
    try {
      const db = await getDb();
      db.exec('ROLLBACK');
    } catch (_) {}
    console.error('Error /api/upload:', err);
    res.status(500).json({ error: err.message });
  }
});

// Export app for Vercel serverless runtime
module.exports = app;

// Start server if run directly
if (require.main === module) {
  const PORT = process.env.PORT || 5005;
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`[Modern Angular + Node.js] Utility Dashboard listening on http://0.0.0.0:${PORT}`);
  });
}
