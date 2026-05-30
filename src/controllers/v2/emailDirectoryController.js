const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const EmailDirectory = require('../../models/v2/emailDirectory.model');
const logger = require('../../utils/logger');

const ALLOWED_EXTENSIONS = ['.csv', '.xlsx', '.xls', '.cdf', '.dlm'];

const normalizeHeader = (value) =>
  String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/_/g, ' ');

const cellToString = (val) => {
  if (val == null || val === '') return '';
  if (typeof val === 'object') {
    if (val.text) return String(val.text).trim();
    if (val.hyperlink) {
      return String(val.hyperlink).replace(/^mailto:/i, '').trim();
    }
    if (val.w) return String(val.w).trim();
  }
  return String(val).trim();
};

const isValidEmail = (email) => {
  const cleaned = email.replace(/\s/g, '');
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleaned);
};

const normalizeEmail = (email) => email.replace(/\s/g, '').toLowerCase();

const findHeaderRowAndColumns = (rows) => {
  const scanLimit = Math.min(rows.length, 10);

  for (let rowIndex = 0; rowIndex < scanLimit; rowIndex++) {
    const row = rows[rowIndex] || [];
    let displayCol = -1;
    let emailCol = -1;

    row.forEach((cell, colIndex) => {
      const h = normalizeHeader(cell);
      if (!h) return;

      if (
        h === 'e-mail display name' ||
        h === 'email display name' ||
        h === 'e mail display name' ||
        (h.includes('display') && h.includes('name') && !h.includes('address'))
      ) {
        displayCol = colIndex;
      }

      if (
        h === 'e-mail address' ||
        h === 'email address' ||
        h === 'e mail address' ||
        ((h.includes('mail') || h.includes('email')) &&
          h.includes('address') &&
          !h.includes('display'))
      ) {
        emailCol = colIndex;
      }
    });

    if (displayCol >= 0 && emailCol >= 0 && displayCol !== emailCol) {
      return { headerRowIndex: rowIndex, displayCol, emailCol };
    }
  }

  // Fallback: exactly 2 columns — col0 = name, col1 = email
  for (let rowIndex = 0; rowIndex < scanLimit; rowIndex++) {
    const row = rows[rowIndex] || [];
    const nonEmpty = row.filter((c) => cellToString(c) !== '');
    if (nonEmpty.length >= 2 && row.length >= 2) {
      const first = normalizeHeader(row[0]);
      const second = normalizeHeader(row[1]);
      const firstLooksHeader =
        first.includes('display') || first.includes('name') || first.includes('email');
      const secondLooksHeader =
        second.includes('address') || second.includes('email') || second.includes('mail');
      if (firstLooksHeader && secondLooksHeader) {
        return { headerRowIndex: rowIndex, displayCol: 0, emailCol: 1 };
      }
    }
  }

  // No header row — assume column 0 = name, column 1 = email from row 0
  if (rows.length > 0 && (rows[0] || []).length >= 2) {
    const r0 = rows[0] || [];
    const c0 = cellToString(r0[0]);
    const c1 = cellToString(r0[1]);
    if (c0.includes('@') && !c1.includes('@')) {
      return { headerRowIndex: -1, displayCol: 1, emailCol: 0 };
    }
    if (!c0.includes('@') && c1.includes('@')) {
      return { headerRowIndex: -1, displayCol: 0, emailCol: 1 };
    }
    return { headerRowIndex: 0, displayCol: 0, emailCol: 1 };
  }

  return null;
};

const parseRowsFromWorkbook = (filePath) => {
  const ext = path.extname(filePath).toLowerCase();
  if (!ALLOWED_EXTENSIONS.includes(ext)) {
    throw new Error('Invalid file type. Use CSV, CDF/DLM, or Excel (.xlsx, .xls)');
  }

  const readOptions = { cellDates: false, raw: false };
  const workbook = xlsx.readFile(filePath, readOptions);
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error('File has no sheets');

  const sheet = workbook.Sheets[sheetName];
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (!rows.length) throw new Error('File has no data rows');

  const layout = findHeaderRowAndColumns(rows);
  if (!layout) {
    throw new Error(
      'Required columns not found. File must have "E-mail Display Name" and "E-mail Address" columns.'
    );
  }

  const { headerRowIndex, displayCol, emailCol } = layout;
  const dataStartRow = headerRowIndex + 1;
  const entries = [];
  const skipped = { empty: 0, invalidEmail: 0, duplicate: 0 };

  for (let i = dataStartRow; i < rows.length; i++) {
    const row = rows[i] || [];
    const displayName = cellToString(row[displayCol]);
    const email = normalizeEmail(cellToString(row[emailCol]));

    if (!displayName && !email) {
      skipped.empty++;
      continue;
    }
    if (!displayName || !email) {
      skipped.empty++;
      continue;
    }

    // Skip repeated header rows inside sheet
    if (
      normalizeHeader(displayName) === 'e-mail display name' ||
      normalizeHeader(displayName) === 'email display name'
    ) {
      continue;
    }

    if (!isValidEmail(email)) {
      skipped.invalidEmail++;
      continue;
    }

    entries.push({ displayName, email });
  }

  if (!entries.length) {
    throw new Error(
      `No valid rows found. Empty/skipped: ${skipped.empty}, invalid email: ${skipped.invalidEmail}`
    );
  }

  logger.info(
    `Email directory parsed: ${entries.length} entries (skipped empty: ${skipped.empty}, invalid: ${skipped.invalidEmail})`
  );

  return entries;
};

/** Keep first occurrence per email (use after reversing file rows so last row in file is first). */
const dedupeEntries = (entries) => {
  const map = new Map();
  for (const entry of entries) {
    const key = entry.email.toLowerCase();
    if (!map.has(key)) map.set(key, entry);
  }
  return Array.from(map.values());
};

/** Newest / last-in-file contacts appear at the top of the list. */
const orderNewestFirst = (entries) => dedupeEntries([...entries].reverse());

exports.getDirectory = async (req, res) => {
  try {
    const doc = await EmailDirectory.findOne({ key: 'global' }).lean();
    res.status(200).json({
      entries: doc?.entries || [],
      fileName: doc?.fileName || null,
      count: doc?.entries?.length || 0,
      updatedAt: doc?.updatedAt || null,
    });
  } catch (error) {
    logger.error(`getDirectory: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

exports.searchDirectory = async (req, res) => {
  try {
    const q = String(req.query.q || '').trim().toLowerCase();
    const limit = Math.min(parseInt(req.query.limit, 10) || 50, 200);
    const doc = await EmailDirectory.findOne({ key: 'global' }).lean();
    let entries = doc?.entries || [];

    if (q) {
      entries = entries.filter(
        (e) =>
          e.displayName.toLowerCase().includes(q) ||
          e.email.toLowerCase().includes(q)
      );
    }

    res.status(200).json({ entries: entries.slice(0, limit), count: entries.length });
  } catch (error) {
    logger.error(`searchDirectory: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

exports.saveDirectoryEntries = async (req, res) => {
  try {
    const { entries } = req.body;
    if (!Array.isArray(entries)) {
      return res.status(400).json({ message: 'entries must be an array' });
    }

    const validated = [];
    for (const e of entries) {
      const displayName = String(e.displayName || '').trim();
      const email = normalizeEmail(String(e.email || ''));
      if (!displayName || !email) continue;
      if (!isValidEmail(email)) continue;
      validated.push({ displayName, email });
    }

    const unique = dedupeEntries(validated);

    const doc = await EmailDirectory.findOneAndUpdate(
      { key: 'global' },
      {
        entries: unique,
        uploadedBy: req.userId,
      },
      { upsert: true, new: true }
    );

    res.status(200).json({
      message: 'Email directory saved',
      count: doc.entries.length,
      entries: doc.entries,
      fileName: doc.fileName,
      updatedAt: doc.updatedAt,
    });
  } catch (error) {
    logger.error(`saveDirectoryEntries: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

exports.uploadDirectory = async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ message: 'File is required' });
  }

  try {
    const parsed = parseRowsFromWorkbook(req.file.path);
    const entries = orderNewestFirst(parsed);

    const doc = await EmailDirectory.findOneAndUpdate(
      { key: 'global' },
      {
        entries,
        fileName: req.file.originalname,
        uploadedBy: req.userId,
      },
      { upsert: true, new: true }
    );

    logger.info(`Email directory uploaded: ${entries.length} entries`);

    res.status(200).json({
      message: `Email directory uploaded successfully (${entries.length} contacts)`,
      count: entries.length,
      fileName: doc.fileName,
      entries: doc.entries,
    });
  } catch (error) {
    logger.error(`uploadDirectory: ${error.message}`);
    res.status(400).json({ message: error.message });
  } finally {
    if (req.file?.path && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
  }
};
