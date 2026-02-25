// calculationInvoiceController.js
const CalculationInvoice = require('../../models/v2/calculationInvoice.model');
const LossesCalculationData = require('../../models/v1/lossesCalculation.model');
const SubClient = require('../../models/v1/subClient.model');
const Policy = require('../../models/v2/policy.model');
const logger = require('../../utils/logger');
const mongoose = require('mongoose');
const ExcelJS = require('exceljs');

// Get losses calculation data for a sub-client, month, and year
exports.getLossesDataForSubClient = async (req, res) => {
  try {
    const { subClientId, month, year } = req.query;

    if (!subClientId || !month || !year) {
      return res.status(400).json({
        message: 'Sub client ID, month, and year are required',
      });
    }

    if (!mongoose.Types.ObjectId.isValid(subClientId)) {
      return res.status(400).json({ message: 'Invalid sub client ID format' });
    }

    const monthNum = parseInt(month);
    const yearNum = parseInt(year);

    if (monthNum < 1 || monthNum > 12) {
      return res.status(400).json({ message: 'Invalid month' });
    }

    // Find losses calculation data matching the month, year, and subClientId
    const lossesData = await LossesCalculationData.find({
      month: monthNum,
      year: yearNum,
      'subClient.subClientId': new mongoose.Types.ObjectId(subClientId),
    });

    if (!lossesData || lossesData.length === 0) {
      return res.status(404).json({
        message: 'No losses calculation data found for the specified sub-client, month, and year',
      });
    }

    // Find the specific sub-client data
    let subClientData = null;
    for (const lossDoc of lossesData) {
      const subClientMatch = lossDoc.subClient.find(
        (sc) => sc.subClientId && sc.subClientId.toString() === subClientId
      );
      if (subClientMatch && subClientMatch.subClientsData) {
        subClientData = subClientMatch.subClientsData;
        break;
      }
    }

    if (!subClientData) {
      return res.status(404).json({
        message: 'Sub-client data not found in losses calculation',
      });
    }

    // Calculate generation unit and drawl unit (after losses)
    const generationUnit = subClientData.grossInjectionMWHAfterLosses
      ? subClientData.grossInjectionMWHAfterLosses * 1000
      : null;
    const drawlUnit = subClientData.drawlMWHAfterLosses
      ? subClientData.drawlMWHAfterLosses * 1000 * (-1)
      : null;

    logger.info(`Retrieved losses data for sub-client: ${subClientId}, month: ${month}, year: ${year}`);
    res.status(200).json({
      generationUnit,
      drawlUnit,
      grossInjectionMWH: subClientData.grossInjectionMWH,
      drawlMWH: subClientData.drawlMWH,
      grossInjectionMWHAfterLosses: subClientData.grossInjectionMWHAfterLosses,
      drawlMWHAfterLosses: subClientData.drawlMWHAfterLosses,
    });
  } catch (error) {
    logger.error(`Error retrieving losses data: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Create a new calculation invoice
exports.createCalculationInvoice = async (req, res) => {
  try {
    const {
      subClientId,
      policyId,
      selectedPolicies,
      solarGenerationMonth, // Legacy field
      solarGenerationYear, // Legacy field
      solarGenerationMonths, // New field - array of months
      adjustmentBillingMonth,
      adjustmentBillingYear,
      dgvclCredit, // New field - shared DGVCL CREDIT
      setOffEntry, // Legacy field
      todEntry, // Legacy field
      roofTop, // Legacy field
      windFarm, // Legacy field
      anyOther, // Legacy field
      manualEntry,
      notes,
    } = req.body;

    if (!subClientId) {
      return res.status(400).json({
        message: 'Sub client ID is required',
      });
    }

    if (!mongoose.Types.ObjectId.isValid(subClientId)) {
      return res.status(400).json({ message: 'Invalid sub client ID format' });
    }

    // Check if using new structure (solarGenerationMonths) or legacy structure
    const isNewStructure = solarGenerationMonths && Array.isArray(solarGenerationMonths) && solarGenerationMonths.length > 0;
    
    if (isNewStructure) {
      // Validate new structure
      if (!solarGenerationMonths[0].policyId) {
        return res.status(400).json({
          message: 'Policy ID is required for at least one month',
        });
      }
      
      // Validate each month has required fields
      for (const month of solarGenerationMonths) {
        if (!month.month || !month.year) {
          return res.status(400).json({
            message: 'Each month must have month and year',
          });
        }
        if (month.policyId && !mongoose.Types.ObjectId.isValid(month.policyId)) {
          return res.status(400).json({ message: 'Invalid policy ID format in month entry' });
        }
      }
    } else {
      // Legacy structure validation
      if (!policyId) {
        return res.status(400).json({
          message: 'Policy ID is required',
        });
      }
      if (!mongoose.Types.ObjectId.isValid(policyId)) {
        return res.status(400).json({ message: 'Invalid policy ID format' });
      }
      if (!solarGenerationMonth || !solarGenerationYear) {
        return res.status(400).json({
          message: 'Solar generation month and year are required',
        });
      }
    }

    if (!adjustmentBillingMonth || !adjustmentBillingYear) {
      return res.status(400).json({
        message: 'Adjustment billing month and year are required',
      });
    }

    // Verify sub client exists
    const subClient = await SubClient.findById(subClientId);
    if (!subClient) {
      return res.status(404).json({ message: 'Sub client not found' });
    }

    // Verify policy exists (for legacy structure or first month in new structure)
    const policyToVerify = isNewStructure ? solarGenerationMonths[0].policyId : policyId;
    if (policyToVerify) {
      const policy = await Policy.findById(policyToVerify);
      if (!policy) {
        return res.status(404).json({ message: 'Policy not found' });
      }
    }

    // Check if invoice already exists for this combination
    // For new structure, check by subClientId, adjustment dates, and month combinations
    let existingInvoice;
    if (isNewStructure) {
      // Create a unique identifier from month combinations
      const monthKeys = solarGenerationMonths.map(m => `${m.month}-${m.year}`).sort().join(',');
      existingInvoice = await CalculationInvoice.findOne({
        subClientId,
        adjustmentBillingMonth,
        adjustmentBillingYear,
        isActive: true,
        $or: [
          { 'solarGenerationMonths': { $exists: true, $ne: [] } },
          { solarGenerationMonth: { $exists: true } }
        ]
      });
      
      // If found, check if month combinations match
      if (existingInvoice && existingInvoice.solarGenerationMonths) {
        const existingMonthKeys = existingInvoice.solarGenerationMonths
          .map(m => `${m.month}-${m.year}`)
          .sort()
          .join(',');
        if (existingMonthKeys !== monthKeys) {
          existingInvoice = null; // Different month combination, treat as new
        }
      }
    } else {
      // Legacy structure - use old query
      existingInvoice = await CalculationInvoice.findOne({
        subClientId,
        policyId,
        solarGenerationMonth,
        solarGenerationYear,
        adjustmentBillingMonth,
        adjustmentBillingYear,
        isActive: true,
      });
    }

    let calculationInvoice;
    if (existingInvoice) {
      // Update existing invoice
      if (isNewStructure) {
        // Update new structure fields
        existingInvoice.solarGenerationMonths = solarGenerationMonths || existingInvoice.solarGenerationMonths || [];
        existingInvoice.dgvclCredit = dgvclCredit !== undefined ? dgvclCredit : existingInvoice.dgvclCredit;
        // Update legacy fields for backward compatibility (use first month's data)
        if (solarGenerationMonths && solarGenerationMonths.length > 0) {
          existingInvoice.solarGenerationMonth = solarGenerationMonths[0].month;
          existingInvoice.solarGenerationYear = solarGenerationMonths[0].year;
          existingInvoice.policyId = solarGenerationMonths[0].policyId || existingInvoice.policyId;
          existingInvoice.selectedPolicies = solarGenerationMonths[0].selectedPolicies || existingInvoice.selectedPolicies || [];
          existingInvoice.setOffEntry = solarGenerationMonths[0].setOffEntry || existingInvoice.setOffEntry || {};
          existingInvoice.todEntry = solarGenerationMonths[0].todEntry || existingInvoice.todEntry || {};
          existingInvoice.roofTop = solarGenerationMonths[0].roofTop || existingInvoice.roofTop || {};
          existingInvoice.windFarm = solarGenerationMonths[0].windFarm || existingInvoice.windFarm || {};
          existingInvoice.anyOther = solarGenerationMonths[0].anyOther || existingInvoice.anyOther || {};
        }
      } else {
        // Update legacy structure fields
        existingInvoice.selectedPolicies = selectedPolicies || existingInvoice.selectedPolicies || [];
        existingInvoice.setOffEntry = setOffEntry || existingInvoice.setOffEntry || {};
        existingInvoice.todEntry = todEntry || existingInvoice.todEntry || {};
        existingInvoice.roofTop = roofTop || existingInvoice.roofTop || {};
        existingInvoice.windFarm = windFarm || existingInvoice.windFarm || {};
        existingInvoice.anyOther = anyOther || existingInvoice.anyOther || {};
      }
      existingInvoice.manualEntry = manualEntry || existingInvoice.manualEntry || {};
      existingInvoice.calculationTable = req.body.calculationTable || existingInvoice.calculationTable || {};
      if (notes !== undefined) {
        existingInvoice.notes = notes;
      }
      await existingInvoice.save();
      calculationInvoice = existingInvoice;
    } else {
      // Create new invoice
      const invoiceData = {
        subClientId,
        adjustmentBillingMonth,
        adjustmentBillingYear,
        manualEntry: manualEntry || {},
        calculationTable: req.body.calculationTable || {},
        notes,
        isActive: true,
      };
      
      if (isNewStructure) {
        // New structure
        invoiceData.solarGenerationMonths = solarGenerationMonths || [];
        invoiceData.dgvclCredit = dgvclCredit;
        // Set legacy fields from first month for backward compatibility
        if (solarGenerationMonths && solarGenerationMonths.length > 0) {
          invoiceData.policyId = solarGenerationMonths[0].policyId;
          invoiceData.solarGenerationMonth = solarGenerationMonths[0].month;
          invoiceData.solarGenerationYear = solarGenerationMonths[0].year;
          invoiceData.selectedPolicies = solarGenerationMonths[0].selectedPolicies || [];
          invoiceData.setOffEntry = solarGenerationMonths[0].setOffEntry || {};
          invoiceData.todEntry = solarGenerationMonths[0].todEntry || {};
          invoiceData.roofTop = solarGenerationMonths[0].roofTop || {};
          invoiceData.windFarm = solarGenerationMonths[0].windFarm || {};
          invoiceData.anyOther = solarGenerationMonths[0].anyOther || {};
        }
      } else {
        // Legacy structure
        invoiceData.policyId = policyId;
        invoiceData.selectedPolicies = selectedPolicies || [];
        invoiceData.solarGenerationMonth = solarGenerationMonth;
        invoiceData.solarGenerationYear = solarGenerationYear;
        invoiceData.setOffEntry = setOffEntry || {};
        invoiceData.todEntry = todEntry || {};
        invoiceData.roofTop = roofTop || {};
        invoiceData.windFarm = windFarm || {};
        invoiceData.anyOther = anyOther || {};
      }
      
      calculationInvoice = new CalculationInvoice(invoiceData);
      await calculationInvoice.save();
    }

    // Populate references for response
    await calculationInvoice.populate('subClientId', 'name consumerNo');
    await calculationInvoice.populate('policyId', 'name');

    logger.info(`Calculation invoice created: ${calculationInvoice._id}`);
    res.status(201).json({
      message: 'Calculation invoice created successfully',
      invoice: calculationInvoice,
    });
  } catch (error) {
    logger.error(`Error creating calculation invoice: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Get all calculation invoices
exports.getAllCalculationInvoices = async (req, res) => {
  try {
    const { subClientId, policyId, isActive } = req.query;
    const query = {};

    if (subClientId) {
      if (!mongoose.Types.ObjectId.isValid(subClientId)) {
        return res.status(400).json({ message: 'Invalid sub client ID format' });
      }
      query.subClientId = subClientId;
    }

    if (policyId) {
      if (!mongoose.Types.ObjectId.isValid(policyId)) {
        return res.status(400).json({ message: 'Invalid policy ID format' });
      }
      query.policyId = policyId;
    }

    if (isActive !== undefined) {
      query.isActive = isActive === 'true';
    }

    const invoices = await CalculationInvoice.find(query)
      .populate('subClientId', 'name consumerNo')
      .populate('policyId', 'name')
      .sort({ createdAt: -1 });

    logger.info(`Retrieved ${invoices.length} calculation invoices`);
    res.status(200).json({ invoices });
  } catch (error) {
    logger.error(`Error retrieving calculation invoices: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Get a single calculation invoice by ID
exports.getCalculationInvoiceById = async (req, res) => {
  try {
    const { invoiceId } = req.params;

    if (!mongoose.Types.ObjectId.isValid(invoiceId)) {
      return res.status(400).json({ message: 'Invalid invoice ID format' });
    }

    const invoice = await CalculationInvoice.findById(invoiceId)
      .populate('subClientId', 'name consumerNo')
      .populate('policyId', 'name policies');

    if (!invoice) {
      logger.warn(`Calculation invoice not found: ${invoiceId}`);
      return res.status(404).json({ message: 'Calculation invoice not found' });
    }

    logger.info(`Retrieved calculation invoice: ${invoiceId}`);
    res.status(200).json({ invoice });
  } catch (error) {
    logger.error(`Error retrieving calculation invoice: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Update a calculation invoice
exports.updateCalculationInvoice = async (req, res) => {
  try {
    const { invoiceId } = req.params;
    const updateData = req.body;

    if (!mongoose.Types.ObjectId.isValid(invoiceId)) {
      return res.status(400).json({ message: 'Invalid invoice ID format' });
    }

    const invoice = await CalculationInvoice.findById(invoiceId);

    if (!invoice) {
      logger.warn(`Calculation invoice not found: ${invoiceId}`);
      return res.status(404).json({ message: 'Calculation invoice not found' });
    }

    // Update fields
    Object.keys(updateData).forEach((key) => {
      if (updateData[key] !== undefined) {
        invoice[key] = updateData[key];
      }
    });

    await invoice.save();

    await invoice.populate('subClientId', 'name consumerNo');
    await invoice.populate('policyId', 'name');

    logger.info(`Calculation invoice updated: ${invoiceId}`);
    res.status(200).json({
      message: 'Calculation invoice updated successfully',
      invoice,
    });
  } catch (error) {
    logger.error(`Error updating calculation invoice: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Delete a calculation invoice
exports.deleteCalculationInvoice = async (req, res) => {
  try {
    const { invoiceId } = req.params;

    if (!mongoose.Types.ObjectId.isValid(invoiceId)) {
      return res.status(400).json({ message: 'Invalid invoice ID format' });
    }

    const invoice = await CalculationInvoice.findById(invoiceId);

    if (!invoice) {
      logger.warn(`Calculation invoice not found: ${invoiceId}`);
      return res.status(404).json({ message: 'Calculation invoice not found' });
    }

    await CalculationInvoice.findByIdAndDelete(invoiceId);

    logger.info(`Calculation invoice deleted: ${invoiceId}`);
    res.status(200).json({ message: 'Calculation invoice deleted successfully' });
  } catch (error) {
    logger.error(`Error deleting calculation invoice: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Get calculation invoice by criteria (subClientId, policyId, dates)
exports.getCalculationInvoiceByCriteria = async (req, res) => {
  try {
    const {
      subClientId,
      policyId,
      solarGenerationMonth,
      solarGenerationYear,
      adjustmentBillingMonth,
      adjustmentBillingYear,
    } = req.query;

    if (!subClientId || !policyId || !solarGenerationMonth || !solarGenerationYear || 
        !adjustmentBillingMonth || !adjustmentBillingYear) {
      return res.status(400).json({
        message: 'Sub client ID, policy ID, and all date fields are required',
      });
    }

    if (!mongoose.Types.ObjectId.isValid(subClientId)) {
      return res.status(400).json({ message: 'Invalid sub client ID format' });
    }

    if (!mongoose.Types.ObjectId.isValid(policyId)) {
      return res.status(400).json({ message: 'Invalid policy ID format' });
    }

    // Try to find invoice using both old and new structures
    const query = {
      subClientId: new mongoose.Types.ObjectId(subClientId),
      adjustmentBillingMonth: parseInt(adjustmentBillingMonth),
      adjustmentBillingYear: parseInt(adjustmentBillingYear),
      isActive: true,
      $or: [
        // Old structure
        {
          policyId: new mongoose.Types.ObjectId(policyId),
          solarGenerationMonth: parseInt(solarGenerationMonth),
          solarGenerationYear: parseInt(solarGenerationYear),
        },
        // New structure - check if solarGenerationMonths array contains matching month
        {
          'solarGenerationMonths': {
            $elemMatch: {
              month: parseInt(solarGenerationMonth),
              year: parseInt(solarGenerationYear),
              policyId: new mongoose.Types.ObjectId(policyId),
            }
          }
        }
      ]
    };

    const invoice = await CalculationInvoice.findOne(query)
      .populate('subClientId', 'name consumerNo')
      .populate('policyId', 'name policies')
      .populate('solarGenerationMonths.policyId', 'name policies');

    if (!invoice) {
      logger.info(`No calculation invoice found for the specified criteria`);
      return res.status(404).json({ 
        message: 'No calculation invoice found for the specified criteria',
        invoice: null 
      });
    }

    logger.info(`Retrieved calculation invoice by criteria: ${invoice._id}`);
    res.status(200).json({ invoice });
  } catch (error) {
    logger.error(`Error retrieving calculation invoice by criteria: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Month names for Excel display
const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

const thinBorder = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' },
};

function formatNum(val) {
  if (val === undefined || val === null || val === '') return '';
  const n = Number(val);
  return isNaN(n) ? String(val) : (Number.isInteger(n) ? n : parseFloat(n.toFixed(2)));
}

function applyCellStyle(cell, opts = {}) {
  const font = { name: 'Times New Roman', size: opts.size || 10, bold: !!opts.bold, italic: !!opts.italic };
  if (opts.fontColor) font.color = { argb: opts.fontColor };
  cell.font = font;
  cell.alignment = { horizontal: opts.horizontal || 'left', vertical: 'middle', wrapText: opts.wrapText !== false };
  cell.border = thinBorder;
  if (opts.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.fill } };
}

/**
 * Build Unit Credit Calculation Excel and return workbook.
 * Expects payload with: subClientId, solarGenerationMonths, adjustmentBillingMonth, adjustmentBillingYear, calculationTable.
 */
function buildUnitCreditExcel(clientName, solarLabel, adjustmentLabel, calculationTable) {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Final Summary', { views: [{ rightToLeft: false }] });

  ws.columns = [
    { width: 8 },  // A Sr.No.
    { width: 55 }, // B Particulars
    { width: 14 }, // C Units
    { width: 14 }, // D Rate
    { width: 16 }, // E Credit
    { width: 16 }, // F Debit
    { width: 22 }, // G Remark
  ];

  let rowNum = 1;

  // Row 1: empty
  ws.getRow(rowNum).height = 15;
  rowNum++;

  // Row 2: FINAL SUMMARY SHEET | Client name
  ws.mergeCells(`A${rowNum}:B${rowNum}`);
  const r2a = ws.getCell(`A${rowNum}`);
  r2a.value = 'FINAL SUMMARY SHEET';
  applyCellStyle(r2a, { bold: true, size: 16, horizontal: 'center', fill: 'FFFF00' });
  ws.mergeCells(`C${rowNum}:G${rowNum}`);
  const r2b = ws.getCell(`C${rowNum}`);
  r2b.value = (clientName || '').toUpperCase();
  applyCellStyle(r2b, { bold: true, size: 16, horizontal: 'center', fill: 'FFFF00' });
  ws.getRow(rowNum).height = 28;
  rowNum++;

  // Row 3: SOLAR GENRATION MONTH | value (light peach, value in red)
  ws.mergeCells(`A${rowNum}:B${rowNum}`);
  ws.getCell(`A${rowNum}`).value = 'SOLAR GENRATION MONTH';
  applyCellStyle(ws.getCell(`A${rowNum}`), { bold: true, size: 14, horizontal: 'center', fill: 'F8CBAD' });
  ws.mergeCells(`C${rowNum}:G${rowNum}`);
  ws.getCell(`C${rowNum}`).value = solarLabel || '';
  applyCellStyle(ws.getCell(`C${rowNum}`), { bold: true, size: 14, horizontal: 'center', fill: 'F8CBAD', fontColor: 'FF0000' });
  ws.getRow(rowNum).height = 28;
  rowNum++;

  // Row 4: ADJUSTMENT BILLING | value (light peach, value in red)
  ws.mergeCells(`A${rowNum}:B${rowNum}`);
  ws.getCell(`A${rowNum}`).value = 'ADJUSTMENT BILLING';
  applyCellStyle(ws.getCell(`A${rowNum}`), { bold: true, size: 14, horizontal: 'center', fill: 'F8CBAD' });
  ws.mergeCells(`C${rowNum}:G${rowNum}`);
  ws.getCell(`C${rowNum}`).value = adjustmentLabel || '';
  applyCellStyle(ws.getCell(`C${rowNum}`), { bold: true, size: 14, horizontal: 'center', fill: 'F8CBAD', fontColor: 'FF0000' });
  ws.getRow(rowNum).height = 28;
  rowNum++;

  // Row 5: Table headers (light gray)
  const headers = ['Sr. No.', 'Particulars', 'Units in kwh', 'Rate (Rs./kwh)', 'Credit Amount', 'Debit Amount', 'Remark'];
  const headerRow = ws.getRow(rowNum);
  headers.forEach((text, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = text;
    applyCellStyle(cell, { bold: true, size: 10, horizontal: 'center', fill: 'D9D9D9' });
  });
  ws.getRow(rowNum).height = 22;
  rowNum++;

  if (!calculationTable) {
    return workbook;
  }

  const section1Order = ['1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7', '1.8', '1.9', '1.10'];
  const section2Order = ['2.1', '2.2', '2.3', '2.4', '2.5'];

  function writeRow(srNo, particulars, units, rate, credit, debit, remark, options = {}) {
    const r = ws.getRow(rowNum);
    const cells = [srNo, particulars, formatNum(units), formatNum(rate), formatNum(credit), formatNum(debit), remark || ''];
    const fill = options.fill || 'FFFFFF';
    const bold = options.bold !== false;
    const italic = options.italic || false;
    const fontColor = options.fontColor;
    cells.forEach((val, i) => {
      const cell = r.getCell(i + 1);
      cell.value = val;
      const isNumericCol = i >= 2 && i <= 5;
      applyCellStyle(cell, {
        horizontal: i === 0 || i === 1 ? 'left' : 'center',
        fill,
        size: options.size || 10,
        bold: options.bold !== undefined ? options.bold : (isNumericCol && val !== ''),
        italic,
        fontColor,
      });
    });
    ws.getRow(rowNum).height = options.height || 22;
    rowNum++;
  }

  function writeSectionHeader(srNo, label, fillHex) {
    const r = ws.getRow(rowNum);
    r.getCell(1).value = srNo;
    applyCellStyle(r.getCell(1), { bold: true, size: 10, horizontal: 'center', fill: fillHex });
    ws.mergeCells(`B${rowNum}:G${rowNum}`);
    const cell = r.getCell(2);
    cell.value = label;
    applyCellStyle(cell, { bold: true, size: 10, horizontal: 'left', fill: fillHex });
    ws.getRow(rowNum).height = 22;
    rowNum++;
  }

  // Section 1 header
  writeSectionHeader('1', 'As Per RE Solar Policy-2023', 'D9D9D9');

  // Section 1 rows: Particulars = 1.1, 1.2 (Sr. No. and Particulars both show key); Electricity Duty only if enabled
  if (calculationTable.section1) {
    for (const key of section1Order) {
      const row = calculationTable.section1[key];
      if (!row) continue;
      writeRow(key, key, row.unitsInKwh, row.rate, row.creditAmount, row.debitAmount, row.remark, { bold: true });
      if (row.subRows && typeof row.subRows === 'object') {
        const subKeys = Object.keys(row.subRows).filter((sk) => {
          if (sk === 'electricityDuty') return row.showElectricityDuty === true;
          return true;
        });
        subKeys.forEach((sk, idx) => {
          const sub = row.subRows[sk];
          if (!sub || typeof sub !== 'object') return;
          const subSr = `${key}.${idx + 1}`;
          const subLabel = sk === 'drawlFromDiscom' ? 'Drawl from DISCOM by Solar Generator' : sk === 'electricityDuty' ? 'Electricity Duty' : sk === 'bankedEnergyPercent' ? 'Banked Energy in %' : (sub.remark || sk);
          writeRow(subSr, subLabel, sub.unitsInKwh, sub.rate, sub.creditAmount, sub.debitAmount, sub.remark, { italic: sk === 'electricityDuty', bold: sk !== 'electricityDuty' });
        });
      }
    }
  }

  // Section 2 header (Other Credit)
  writeSectionHeader('2', 'Other Credit (if any)', 'B4C6E7');

  if (calculationTable.section2) {
    for (const key of section2Order) {
      const row = calculationTable.section2[key];
      if (!row) continue;
      writeRow(key, key, row.unitsInKwh, row.rate, row.creditAmount, row.debitAmount, row.remark, { bold: true });
      if (row.subRows && typeof row.subRows === 'object') {
        const subKeys = Object.keys(row.subRows).filter((sk) => {
          if (sk === 'electricityDuty') return row.showElectricityDuty === true;
          return true;
        });
        subKeys.forEach((sk, idx) => {
          const sub = row.subRows[sk];
          if (!sub || typeof sub !== 'object') return;
          const subSr = `${key}.${idx + 1}`;
          const subLabel = sk === 'electricityDuty' ? 'Electricity Duty' : (sub.remark || sk);
          writeRow(subSr, subLabel, sub.unitsInKwh, sub.rate, sub.creditAmount, sub.debitAmount, sub.remark, { italic: sk === 'electricityDuty', bold: sk !== 'electricityDuty' });
        });
      }
    }
  }

  // Section 3 header (Other Debit)
  writeSectionHeader('3', 'Other Debit (if any)', 'B4C6E7');

  if (Array.isArray(calculationTable.section3)) {
    let sec3Idx = 1;
    for (const row of calculationTable.section3) {
      const sr = String(sec3Idx++);
      const part = row.particulars || row.remark || row.id || '';
      writeRow(sr, part, row.unitsInKwh, row.rate, row.creditAmount, row.debitAmount, row.remark, { bold: true });
    }
  }

  // Empty row before final
  rowNum++;

  // Final section (yellow background, red font for row2/row3 labels per image)
  const fs = calculationTable.finalSection;
  if (fs) {
    const r1 = fs.row1 || {};
    writeRow('', 'TOTAL AMOUNT', '', '', formatNum(r1.creditAmount), formatNum(r1.debitAmount), '', { fill: 'FFFF00', bold: true });
    const r2 = fs.row2 || {};
    writeRow('', 'TOTAL AMOUNT - CREDITABLE', '', '', formatNum(r2.mergedValue), '', '', { fill: 'FFFF00', bold: true, fontColor: 'FF0000' });
    const r3 = fs.row3 || {};
    writeRow('', 'AMOUNT IN BILL - CREDITED IN DISCOM', '', '', formatNum(r3.mergedValue), '', '', { fill: 'FFFF00', bold: true, fontColor: 'FF0000' });
    const r4 = fs.row4 || {};
    const status = r4.status || '';
    const lastRow = ws.getRow(rowNum);
    lastRow.getCell(1).value = '';
    lastRow.getCell(2).value = 'DIFF. IN AMOUNT';
    lastRow.getCell(3).value = '';
    lastRow.getCell(4).value = '';
    lastRow.getCell(5).value = formatNum(r4.mergedValue);
    lastRow.getCell(6).value = '';
    lastRow.getCell(7).value = status;
    [1, 2, 3, 4, 5, 6, 7].forEach((c) => {
      const cell = lastRow.getCell(c);
      applyCellStyle(cell, { fill: c === 7 && status ? '92D050' : 'FFFF00', bold: true });
    });
    ws.getRow(rowNum).height = 22;
    rowNum++;
  }

  return workbook;
}

/**
 * Export Unit Credit Calculation to Excel.
 * POST body: subClientId, solarGenerationMonths[], adjustmentBillingMonth, adjustmentBillingYear, calculationTable.
 */
exports.exportCalculationToExcel = async (req, res) => {
  try {
    const {
      subClientId,
      solarGenerationMonths,
      adjustmentBillingMonth,
      adjustmentBillingYear,
      calculationTable,
    } = req.body;

    if (!subClientId) {
      return res.status(400).json({ message: 'Sub client ID is required' });
    }
    if (!mongoose.Types.ObjectId.isValid(subClientId)) {
      return res.status(400).json({ message: 'Invalid sub client ID format' });
    }

    const subClient = await SubClient.findById(subClientId).select('name');
    const clientName = subClient ? (subClient.name || '').trim() : '';

    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    let solarLabel = '';
    if (solarGenerationMonths && Array.isArray(solarGenerationMonths) && solarGenerationMonths.length > 0) {
      solarLabel = solarGenerationMonths
        .map((m) => {
          const mo = m.month != null ? m.month : 0;
          const yr = m.year != null ? m.year : 0;
          return `${monthNames[mo - 1] || ''}-${String(yr).slice(-2)}`;
        })
        .filter(Boolean)
        .join(', ') || '';
    }

    const adjMonth = adjustmentBillingMonth != null ? Number(adjustmentBillingMonth) : 0;
    const adjYear = adjustmentBillingYear != null ? Number(adjustmentBillingYear) : 0;
    const adjustmentLabel = adjMonth && adjYear ? `${monthNames[adjMonth - 1] || ''}-${String(adjYear).slice(-2)}` : '';

    const workbook = buildUnitCreditExcel(clientName, solarLabel, adjustmentLabel, calculationTable || {});

    const sanitize = (str) => (str || '').replace(/[^a-zA-Z0-9\s-]/g, '').replace(/\s+/g, ' ').trim();
    const fileName = `Unit Credit Calculation - ${sanitize(clientName) || 'Client'} ${adjustmentLabel || 'Report'}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fileName)}"`);
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    logger.error(`Error exporting calculation to Excel: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};
