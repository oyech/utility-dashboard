const express = require('express');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const { DatabaseSync } = require('node:sqlite');
const ExcelJS = require('exceljs');
const { buildUtilityWorkbook } = require('./excel-exporter');

const app = express();
const DB_PATH = path.join(__dirname, 'utility.db');

// Initialize SQLite database
const db = new DatabaseSync(DB_PATH);

// Middleware
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use('/static', express.static(path.join(__dirname, 'static')));

// Upload configuration
const upload = multer({ dest: path.join(__dirname, 'uploads/') });

// Ensure uploads folder exists
if (!fs.existsSync(path.join(__dirname, 'uploads'))) {
  fs.mkdirSync(path.join(__dirname, 'uploads'), { recursive: true });
}

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
app.get('/api/units', (req, res) => {
  try {
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
    const stmt = db.prepare(query);
    const rows = stmt.all(...params);
    res.json(rows);
  } catch (err) {
    console.error('Error /api/units:', err);
    res.status(500).json({ error: err.message });
  }
});

// 2. Get distinct periods
app.get('/api/periods', (req, res) => {
  try {
    const rows = db.prepare('SELECT DISTINCT period FROM billing_data ORDER BY period DESC').all();
    const periods = rows.map(r => r.period);
    res.json(periods);
  } catch (err) {
    console.error('Error /api/periods:', err);
    res.status(500).json({ error: err.message });
  }
});

// 3. Get distinct towers
app.get('/api/towers', (req, res) => {
  try {
    const rows = db.prepare("SELECT DISTINCT tower FROM billing_data WHERE tower IS NOT NULL AND tower != '' ORDER BY tower").all();
    const towers = rows.map(r => r.tower);
    res.json(towers);
  } catch (err) {
    console.error('Error /api/towers:', err);
    res.status(500).json({ error: err.message });
  }
});

// Helper to fetch usage records
function fetchUnitUsage(unit, periods) {
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
  return db.prepare(query).all(...params);
}

// 4. Get usage for a unit across periods
app.get('/api/usage', (req, res) => {
  try {
    const unit = (req.query.unit || '').trim();
    if (!unit) {
      return res.status(400).json({ error: 'Unit is required' });
    }

    const rows = fetchUnitUsage(unit, req.query.periods);
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

    const rows = fetchUnitUsage(unit, periods);
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

// 6. Upload Excel file and import data
app.post('/api/upload', upload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }
  const period = (req.body.period || '').trim();
  if (!period) {
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    return res.status(400).json({ error: 'Period is required (e.g. 2026-11)' });
  }
  if (!req.file.originalname.endsWith('.xlsx')) {
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    return res.status(400).json({ error: 'Only .xlsx files are supported' });
  }

  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(req.file.path);
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

      rowsToInsert.push([
        period,
        extractVal(1),
        extractVal(2),
        String(unitVal).trim(),
        extractVal(4),
        extractVal(5),
        extractVal(6),
        extractVal(7),
        extractVal(8),
        extractVal(9),
        extractVal(10),
        extractVal(11),
        extractVal(12),
        extractVal(13),
        extractVal(14),
        extractVal(15),
        extractVal(16),
        extractVal(17),
        extractVal(18),
        extractVal(19),
        extractVal(20),
        extractVal(21),
        extractVal(22),
        extractVal(23),
        extractVal(24),
        extractVal(25),
        extractVal(26)
      ]);
    });

    // Execute SQLite transaction
    db.exec('BEGIN TRANSACTION');
    db.prepare('DELETE FROM billing_data WHERE period = ?').run(period);

    const insertStmt = db.prepare(`
      INSERT INTO billing_data (
        period, no, tower, unit, owner_name, mcb, kva, note,
        elec_meter_start, elec_meter_end, elec_usage, kwh_min, jam_nyala, elec_subtotal, elec_total,
        water_meter_start, water_meter_end, water_usage, water_m1, water_m2, water_m3,
        water_cost_m1, water_cost_m2, water_cost_m3, water_fixed_cost, water_total, grand_total
      ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    `);

    for (const r of rowsToInsert) {
      insertStmt.run(...r);
    }
    db.exec('COMMIT');

    if (fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }

    res.json({ message: `Berhasil mengimpor ${rowsToInsert.length} unit untuk periode ${period}!` });
  } catch (err) {
    db.exec('ROLLBACK');
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
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
    console.log(`[Node.js + AngularJS] Utility Dashboard listening on http://0.0.0.0:${PORT}`);
  });
}
