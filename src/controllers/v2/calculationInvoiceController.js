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

/**
 * Find a calculation invoice by sub-client, policy, and solar generation month/year only.
 * Ignores adjustment billing month/year so prior Step 5 total consumption can be loaded for row 1.5
 * even when the user saved the previous month under a different adjustment billing period.
 */
exports.getCalculationInvoiceBySolarPeriod = async (req, res) => {
  try {
    const {
      subClientId,
      policyId,
      solarGenerationMonth,
      solarGenerationYear,
    } = req.query;

    if (!subClientId || !policyId || !solarGenerationMonth || !solarGenerationYear) {
      return res.status(400).json({
        message: 'Sub client ID, policy ID, solar generation month, and solar generation year are required',
      });
    }

    if (!mongoose.Types.ObjectId.isValid(subClientId)) {
      return res.status(400).json({ message: 'Invalid sub client ID format' });
    }

    if (!mongoose.Types.ObjectId.isValid(policyId)) {
      return res.status(400).json({ message: 'Invalid policy ID format' });
    }

    const sm = parseInt(solarGenerationMonth, 10);
    const sy = parseInt(solarGenerationYear, 10);

    const query = {
      subClientId: new mongoose.Types.ObjectId(subClientId),
      isActive: true,
      $or: [
        {
          policyId: new mongoose.Types.ObjectId(policyId),
          solarGenerationMonth: sm,
          solarGenerationYear: sy,
        },
        {
          solarGenerationMonths: {
            $elemMatch: {
              month: sm,
              year: sy,
              policyId: new mongoose.Types.ObjectId(policyId),
            },
          },
        },
      ],
    };

    const invoice = await CalculationInvoice.findOne(query)
      .sort({ updatedAt: -1 })
      .populate('subClientId', 'name consumerNo')
      .populate('policyId', 'name policies')
      .populate('solarGenerationMonths.policyId', 'name policies');

    if (!invoice) {
      logger.info('No calculation invoice found for the specified solar period');
      return res.status(404).json({
        message: 'No calculation invoice found for the specified solar period',
        invoice: null,
      });
    }

    logger.info(`Retrieved calculation invoice by solar period: ${invoice._id}`);
    res.status(200).json({ invoice });
  } catch (error) {
    logger.error(`Error retrieving calculation invoice by solar period: ${error.message}`);
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

const mediumHeaderBorder = {
  top: { style: 'medium' },
  left: { style: 'thin' },
  bottom: { style: 'medium' },
  right: { style: 'thin' },
};

const mediumTopBorder = {
  top: { style: 'medium' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' },
};

const mediumBottomBorder = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'medium' },
  right: { style: 'thin' },
};

function formatNum(val) {
  if (val === undefined || val === null || val === '') return '';
  const n = Number(val);
  return isNaN(n) ? String(val) : n;
}

function applyCellStyle(cell, opts = {}) {
  const font = { name: 'Times New Roman', size: opts.size || 10, bold: !!opts.bold, italic: !!opts.italic };
  if (opts.fontColor) font.color = { argb: opts.fontColor };
  cell.font = font;
  cell.alignment = { horizontal: opts.horizontal || 'left', vertical: 'middle', wrapText: opts.wrapText !== false };
  cell.border = opts.border || thinBorder;
  if (opts.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.fill } };
  if (opts.numFmt) cell.numFmt = opts.numFmt;
}

/** Footer note for 2021-style export (frontend sets policy2021Style; static ₹2.25/kWh × row 1.10 units). */
function buildPolicy2021SurplusFooterText(calculationTable) {
  const raw = calculationTable?.section1?.['1.10']?.unitsInKwh;
  let n = typeof raw === 'number' ? raw : parseFloat(String(raw ?? '').replace(/,/g, ''));
  if (!Number.isFinite(n)) n = 0;
  const rate = 2.25;
  let unitsDisplay = n === 0 ? '0' : (Number.isInteger(n) ? String(n) : n.toFixed(4).replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, ''));
  const amt = n * rate;
  const parts = amt.toFixed(2).split('.');
  const intWithCommas = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const amountStr = `${intWithCommas}.${parts[1]}`;
  return `Surplus Energy after Set-Off ( Inadvertent Energy ) :- ${unitsDisplay} units × ₹${rate}→ approx. ₹${amountStr} payable by DISCOM. Note:- This is indicative only. For exact figures, please reach out to the relevant DISCOM  Division office`;
}

/**
 * Build Unit Credit Calculation Excel and return workbook.
 * Expects payload with: subClientId, solarGenerationMonths, adjustmentBillingMonth, adjustmentBillingYear, calculationTable.
 * @param {{ policy2021Style?: boolean }} [excelOptions]
 */
function buildUnitCreditExcel(clientName, solarLabel, adjustmentLabel, calculationTable, excelOptions = {}) {
  const policy2021Style = !!excelOptions.policy2021Style;
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Final Summary', { views: [{ rightToLeft: false }] });

  ws.columns = [
    { width: 4.5 },  // A Sr.No.
    { width: 38 }, // B Particulars
    { width: 11.5 }, // C Units
    { width: 10.5 }, // D Rate
    { width: 15.5 }, // E Credit
    { width: 14.5 }, // F Debit
    { width: 10.5 }, // G Remark
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
  ws.getRow(rowNum).height = 45;
  rowNum++;

  // Row 3: SOLAR GENRATION MONTH | value (light peach, value in red)
  ws.mergeCells(`A${rowNum}:B${rowNum}`);
  ws.getCell(`A${rowNum}`).value = 'SOLAR GENRATION MONTH';
  applyCellStyle(ws.getCell(`A${rowNum}`), { bold: true, size: 14, horizontal: 'center', fill: 'FFFF00', fontColor: 'FF0000' });
  ws.mergeCells(`C${rowNum}:G${rowNum}`);
  ws.getCell(`C${rowNum}`).value = solarLabel || '';
  applyCellStyle(ws.getCell(`C${rowNum}`), { bold: true, size: 14, horizontal: 'center', fill: 'FFFF00', fontColor: 'FF0000' });
  ws.getRow(rowNum).height = 19;
  rowNum++;

  // Row 4: ADJUSTMENT BILLING | value (light peach, value in red)
  ws.mergeCells(`A${rowNum}:B${rowNum}`);
  ws.getCell(`A${rowNum}`).value = 'ADJUSTMENT BILLING';
  applyCellStyle(ws.getCell(`A${rowNum}`), { bold: true, size: 14, horizontal: 'center', fill: 'FFFF00', fontColor: 'FF0000' });
  ws.mergeCells(`C${rowNum}:G${rowNum}`);
  ws.getCell(`C${rowNum}`).value = adjustmentLabel || '';
  applyCellStyle(ws.getCell(`C${rowNum}`), { bold: true, size: 14, horizontal: 'center', fill: 'FFFF00', fontColor: 'FF0000' });
  ws.getRow(rowNum).height = 19;
  rowNum++;

  // Row 5: Table headers (light gray)
  const headers = ['Sr. No.', 'Particulars', 'Units in kwh', 'Rate (Rs./kwh)', 'Credit Amount', 'Debit Amount', 'Remark'];
  const headerRow = ws.getRow(rowNum);
  headers.forEach((text, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = text;
    applyCellStyle(cell, { bold: true, size: 11, horizontal: 'center', fill: 'FFC000', border: mediumHeaderBorder });
  });
  ws.getRow(rowNum).height = 30;
  rowNum++;

  if (!calculationTable) {
    return workbook;
  }

  const section1Order = policy2021Style
    ? ['1.1', '1.2', '1.3', '1.4', '1.9', '1.10']
    : ['1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7', '1.8', '1.9', '1.10'];
  const section2Order = ['2.1', '2.2', '2.3', '2.4'];

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

      let horizontal = 'left';
      if (i === 0 || i === 3 || i === 4 || i === 5) horizontal = 'center';
      if (i === 2) horizontal = 'right';
      if (i === 1 && options.horizontal) horizontal = options.horizontal;

      let numFmt = undefined;
      const particularsStr = typeof particulars === 'string' ? particulars : '';
      const isPercentageRow = particularsStr.toLowerCase().includes('banked energy in') || particularsStr.toLowerCase().includes('electricity duty');

      if (typeof val === 'number') {
        if (i === 2) { // Units in kwh
          numFmt = isPercentageRow ? '0.00\"%\"' : '0';
        } else if (options.numFmt) {
          numFmt = options.numFmt;
        } else if (i === 3) { // Rate
          numFmt = isPercentageRow ? '0\"%\"' : (options.rateTwoDecimals ? '\"₹\" #,##0.00' : '\"₹\" #,##0.0000');
        } else if (i === 4 || i === 5) { // Credit / Debit
          numFmt = options.amountOneDecimal ? '\"₹\" #,##0.0' : '\"₹\" #,##0.00';
        }
      }

      let cellBold = options.bold !== undefined ? options.bold : (isNumericCol && val !== '');
      let cellItalic = italic;
      let cellFontColor = fontColor;
      let cellFill = fill;

      // New requirement: Rate, Credit, and Debit columns should be italic not bold
      // Units in kwh (index 2) should remain as is (bold).
      if (options.italicNumeric && (i >= 3 && i <= 5) && val !== '') {
        cellBold = false;
        cellItalic = (i !== 3); // Rate (i=3) is no longer italic, only Credit (4) and Debit (5) are
      }

      // Individual Column/Cell Overrides
      if (options.colStyles && options.colStyles[i]) {
        const cs = options.colStyles[i];
        if (cs.fill) cellFill = cs.fill;
        if (cs.fontColor) cellFontColor = cs.fontColor;
        if (cs.bold !== undefined) cellBold = cs.bold;
        if (cs.italic !== undefined) cellItalic = cs.italic;
        if (cs.numFmt !== undefined) numFmt = cs.numFmt;
      }

      // Explicitly remove italic for Rate column
      if (i === 3) {
        cellItalic = false;
      }

      // Explicitly remove italic for Debit Amount column, EXCEPT for F16 (item 1.8)
      if (i === 5) {
        const pStr = typeof particulars === 'string' ? particulars : (particulars && particulars.richText ? particulars.richText.map(rt => rt.text).join('') : '');
        if (!pStr.includes('Maxi.30% Eligible for Set-Off Banked Energy')) {
          cellItalic = false;
        }
      }

      applyCellStyle(cell, {
        horizontal,
        fill: cellFill,
        size: options.size || 10,
        bold: cellBold,
        italic: cellItalic,
        fontColor: cellFontColor,
        numFmt,
        border: options.border
      });
    });
    if (options.height) {
      ws.getRow(rowNum).height = options.height;
    } else {
      const pStrForHeight = typeof particulars === 'string' ? particulars : (particulars && particulars.richText ? particulars.richText.map(rt => rt.text).join('') : '');
      if (pStrForHeight.length > 110) {
        ws.getRow(rowNum).height = 60;
      } else if (pStrForHeight.length > 75) {
        ws.getRow(rowNum).height = 45;
      } else if (pStrForHeight.length > 40) {
        ws.getRow(rowNum).height = 30;
      } else {
        ws.getRow(rowNum).height = 15; // default minimum
      }
    }
    rowNum++;
  }

  function writeSectionHeader(srNo, label, fillHex) {
    const r = ws.getRow(rowNum);
    r.getCell(1).value = srNo;
    applyCellStyle(r.getCell(1), { bold: true, size: 11, horizontal: 'center', fill: fillHex });

    const cell = r.getCell(2);
    cell.value = label;
    applyCellStyle(cell, { bold: true, size: 11, horizontal: 'left', fill: fillHex });

    // Apply styling to empty cells instead of merging
    for (let c = 3; c <= 7; c++) {
      applyCellStyle(r.getCell(c), { fill: fillHex });
    }

    ws.getRow(rowNum).height = 15;
    rowNum++;
  }

  // Section 1 header
  writeSectionHeader(1, 'As Per RE Solar Policy-2023', 'BDD7EE');

  const srNoMap = {
    '1.1': 'A',
    '1.2': 'B',
    '1.3': 'C',
    '1.4': 'D',
    '1.9': 'F',
    '1.10': 'G',
  };

  const section2SrMap = {
    '2.1': 1,
    '2.2': 2,
    '2.3': 3,
    '2.4': 4,
  };

  const section1FallbackLabels = {
    '1.1': 'Gross Solar Generation as per SLDC- ( Policy-2023)',
    '1.2': 'Net Solar Generation after Wheeling loss = A * (1-7.25%)',
    '1.3': 'Generation Set-Off with Consumption 15-min basis',
    '1.4': 'Surplus Energy after Set-Off ( Banked Energy )',
    '1.5': `Total Consumption from DISCOM for the Month of ${adjustmentLabel}`,
    '1.6': 'Net Consumption from DISCOM after Generation Set-Off with Consumption 15-min basis',
    '1.7': 'Maxi.30% Eligible - This is indicative only. For exact figures, please reach out to the relevant DISCOM Division office',
    '1.8': 'As per RE Policy2023 Maxi.30% Eligible for Set-Off Banked Energy @ Net Consumption from DISCOM',
    '1.9': 'TOTAL Set-Off',
    '1.10': policy2021Style
      ? 'Surplus Energy after Set-Off ( Inadvertent Energy ) SELL Units to DISCOM'
      : 'Inadvertent Banked Energy = B - F if Any ( LAPSED )',
  };

  const section2FallbackLabels = {
    '2.1': 'TOD Solar Concession (11:00 to 15:00 Hrs / kwh )',
    '2.2': 'TCS / TDS',
    '2.3': 'Roof Top Solar',
    '2.4': 'Roof Top Solar-SELL Unit',

  };

  // Section 1 rows: Particulars from data (if unique) or fallback labels; Sr. No. = A, B, C, D, (blank for others), F, G
  if (calculationTable.section1) {
    for (const key of section1Order) {
      const row = calculationTable.section1[key];
      if (!row) continue;

      const mainSr = srNoMap[key] || '';
      // Use dynamic particulars if it's provided and not just the technical key
      let mainLabel = (row.particulars && row.particulars !== key) ? row.particulars : (section1FallbackLabels[key] || key);
      const mainOpts = { bold: true, italicNumeric: true };

      if (key === '1.1') {
        const baseTxt = typeof mainLabel === 'string' ? mainLabel : section1FallbackLabels['1.1'];
        if (baseTxt.includes('Policy-2023)')) {
          const parts = baseTxt.split('Policy-2023)');
          mainLabel = {
            richText: [
              { text: parts[0], font: { bold: true, size: 11, name: 'Times New Roman' } },
              { text: 'Policy-2023)', font: { bold: false, italic: true, size: 11, name: 'Times New Roman' } },
              { text: parts[1] || '', font: { bold: true, size: 11, name: 'Times New Roman' } }
            ]
          };
        }
      }

      if (key === '1.2') {
        const baseTxt = typeof mainLabel === 'string' ? mainLabel : section1FallbackLabels['1.2'];
        if (baseTxt.includes('A * (1-7.25%)')) {
          const parts = baseTxt.split('A * (1-7.25%)');
          mainLabel = {
            richText: [
              { text: parts[0], font: { bold: true, size: 11, name: 'Times New Roman' } },
              { text: 'A * (1-7.25%)', font: { bold: false, italic: true, size: 11, name: 'Times New Roman' } },
              { text: parts[1] || '', font: { bold: true, size: 11, name: 'Times New Roman' } }
            ]
          };
        }
      }

      if (key === '1.4') {
        const baseTxt = typeof mainLabel === 'string' ? mainLabel : section1FallbackLabels['1.4'];
        if (baseTxt.includes('( Banked Energy )')) {
          const parts = baseTxt.split('( Banked Energy )');
          mainLabel = {
            richText: [
              { text: parts[0], font: { bold: true, size: 11, name: 'Times New Roman' } },
              { text: '( Banked Energy )', font: { bold: false, italic: true, size: 11, name: 'Times New Roman' } },
              { text: parts[1] || '', font: { bold: true, size: 11, name: 'Times New Roman' } }
            ]
          };
        }
      }

      if (key === '1.5') {
        mainLabel = {
          richText: [
            { text: 'Total Consumption from DISCOM for the Month of ', font: { bold: true, size: 11, color: { argb: '000000' }, name: 'Times New Roman' } },
            { text: adjustmentLabel, font: { bold: true, size: 11, color: { argb: 'FF0031' }, name: 'Times New Roman' } }
          ]
        };
      }

      if (key === '1.7') {
        mainLabel = {
          richText: [
            { text: 'Maxi.30% Eligible', font: { italic: true, bold: true, size: 10, color: { argb: 'FF0000' }, name: 'Times New Roman' } },
            { text: ' - This is indicative only. For exact figures, please reach out to the relevant DISCOM Division office', font: { italic: true, size: 10, color: { argb: '000000' }, name: 'Times New Roman' } }
          ]
        };
      }

      if (key === '1.10' && !policy2021Style) {
        mainLabel = {
          richText: [
            { text: 'Inadvertent Banked Energy = B - F if Any ( ', font: { bold: true, size: 10, color: { argb: '000000' }, name: 'Times New Roman' } },
            { text: 'LAPSED', font: { bold: true, size: 10, color: { argb: 'FF0000' }, name: 'Times New Roman' } },
            { text: ' )', font: { bold: true, size: 10, color: { argb: '000000' }, name: 'Times New Roman' } }
          ]
        };
      }

      // Formatting based on UI/Excel requirements
      if (['1.5', '1.6', '1.7'].includes(key)) {
        mainOpts.fill = 'FCE4D6';
      }
      if (key === '1.5') mainOpts.border = mediumTopBorder;
      if (key === '1.7') mainOpts.border = mediumBottomBorder;

      if (['1.2', '1.4', '1.10'].includes(key)) {
        // Handled by default Units formatting now
      }
      if (['1.6', '1.7', '1.8', '1.9'].includes(key)) {
        mainOpts.numFmt = '\"₹\" #,##0.00';
      }

      if (['1.6', '1.7', '1.8'].includes(key)) {
        mainOpts.bold = false;
        mainOpts.italic = true;
        mainOpts.horizontal = 'right';
        if (['1.6', '1.7'].includes(key)) {
          mainOpts.colStyles = { 2: { bold: true, italic: false } };
        } else if (key === '1.8') {
          mainOpts.colStyles = { 2: { italic: false } };
        }
      }

      if (key === '1.9') {
        mainOpts.height = 30;
      }

      if (key === '1.10') {
        mainOpts.colStyles = {
          2: { fill: 'FFC7CE', fontColor: 'AD0006', bold: true, italic: false }
        };
      }

      writeRow(mainSr, mainLabel, row.unitsInKwh, row.rate, row.creditAmount, row.debitAmount, row.remark, mainOpts);

      if (row.subRows && typeof row.subRows === 'object') {
        const subKeys = Object.keys(row.subRows).filter((sk) => {
          if (sk === 'electricityDuty') return row.showElectricityDuty === true;
          return true;
        });
        subKeys.forEach((sk, idx) => {
          const sub = row.subRows[sk];
          if (!sub || typeof sub !== 'object') return;
          const subSr = ''; // Sub-rows have empty Sr. No. as per image
          const subLabel = sub.particulars || (sk === 'drawlFromDiscom' ? 'Drawl from DISCOM by Solar Generator' : sk === 'electricityDuty' ? 'Electricity Duty' : sk === 'bankedEnergyPercent' ? 'Banked Energy in %' : (sub.remark || sk));

          const subOpts = { italic: false, bold: false, horizontal: 'right', colStyles: { 1: { italic: true } } };
          if (sk === 'bankedEnergyPercent') {
            subOpts.colStyles[2] = { bold: true };
          }

          writeRow(subSr, subLabel, sub.unitsInKwh, sub.rate, sub.creditAmount, sub.debitAmount, sub.remark, subOpts);
        });
      }
    }
  }

  const emptyRowSec1 = rowNum;
  rowNum++; // Empty row after Section 1

  const sec2StartRow = rowNum;
  // Section 2 header (Other Credit)
  writeSectionHeader(2, 'Other Credit (if any)', 'BDD7EE');

  if (calculationTable.section2) {
    for (const key of section2Order) {
      const row = calculationTable.section2[key];
      if (!row) continue;
      const mainSr = section2SrMap[key] || '';
      const mainLabel = (row.particulars && row.particulars !== key) ? row.particulars : (section2FallbackLabels[key] || key);
      writeRow(mainSr, mainLabel, row.unitsInKwh, row.rate, row.creditAmount, row.debitAmount, row.remark, { bold: false, italicNumeric: true, rateTwoDecimals: true, colStyles: { 2: { bold: true, italic: false }, 4: { italic: key === '2.1', numFmt: '"₹" #,##0.0' }, 5: { numFmt: '"₹" #,##0.0' } } });
      if (row.subRows && typeof row.subRows === 'object') {
        const subKeys = Object.keys(row.subRows).filter((sk) => {
          if (sk === 'electricityDuty') return row.showElectricityDuty === true;
          return true;
        });
        subKeys.forEach((sk, idx) => {
          const sub = row.subRows[sk];
          if (!sub || typeof sub !== 'object') return;
          const subSr = ''; // Sub-rows have empty Sr. No.
          const subLabel = sub.particulars || (sk === 'electricityDuty' ? 'Electricity Duty' : (sub.remark || sk));
          writeRow(subSr, subLabel, sub.unitsInKwh, sub.rate, sub.creditAmount, sub.debitAmount, sub.remark, { italic: true, bold: false, horizontal: 'right', rateTwoDecimals: true, colStyles: { 4: { italic: false, numFmt: '"₹" #,##0.0' }, 5: { numFmt: '"₹" #,##0.0' } } });
        });
      }
    }
  }

  rowNum++; // Space after Other Credit (if any)
  const sec3StartRow = rowNum;

  // Section 3 header (Other Debit)
  writeSectionHeader(3, 'Other Debit (if any)', 'BDD7EE');

  if (Array.isArray(calculationTable.section3)) {
    let sec3Idx = 1;
    for (const row of calculationTable.section3) {
      const sr = sec3Idx++;
      const part = row.particulars || row.remark || row.id || '';
      writeRow(sr, part, row.unitsInKwh, row.rate, row.creditAmount, row.debitAmount, row.remark, { bold: false, italicNumeric: true, rateTwoDecimals: true, colStyles: { 2: { bold: true, italic: false } } });
    }
  }

  // Empty row before final
  rowNum++;

  // Final section logic redesigned to match Image 2
  const fs = calculationTable.finalSection;
  let footerStartRow = rowNum; // fallback
  let rowAfterFooter = rowNum; // fallback
  if (fs) {
    const startFooterRow = rowNum;
    footerStartRow = rowNum;
    const r1 = fs.row1 || {};
    const r2 = fs.row2 || {};
    const r3 = fs.row3 || {};
    const r4 = fs.row4 || {};
    const status = r4.status || '';

    const footerRows = [
      { label: 'TOTAL AMOUNT', credit: r1.creditAmount, debit: r1.debitAmount, color: '000000', align: 'center' },
      { label: 'TOTAL AMOUNT - CREDITABLE', credit: r2.mergedValue, color: 'FF0000', align: 'right' },
      { label: 'AMOUNT IN BILL - CREDITED IN DISCOM', credit: r3.mergedValue, color: 'FF0000', align: 'right' },
      { label: 'DIFF. IN AMOUNT', credit: r4.mergedValue, color: 'FF0000', align: 'right' }
    ];

    footerRows.forEach((fr, i) => {
      const r = ws.getRow(rowNum);
      r.height = 25;

      // Merge A-D for Labels
      ws.mergeCells(`A${rowNum}:D${rowNum}`);
      const labelCell = r.getCell(1);
      labelCell.value = fr.label;
      applyCellStyle(labelCell, { bold: true, horizontal: fr.align, fontColor: fr.color, size: 11 });

      // Columns E (Credit) and F (Debit) - Yellow background
      const creditCell = r.getCell(5);
      const debitCell = r.getCell(6);

      if (i > 0) {
        // Merge E and F for rows 2, 3, 4
        ws.mergeCells(`E${rowNum}:F${rowNum}`);
        creditCell.value = fr.credit != null ? Number(fr.credit) : '';
      } else {
        creditCell.value = fr.credit != null ? Number(fr.credit) : '';
        debitCell.value = fr.debit != null ? Number(fr.debit) : '';
      }

      // Numeric formats
      const numFmt = '\"₹\" #,##0.00';
      creditCell.numFmt = numFmt;
      debitCell.numFmt = numFmt;

      applyCellStyle(creditCell, { fill: 'FFFF00', bold: true, horizontal: 'center', fontColor: fr.color, size: 11 });
      applyCellStyle(debitCell, { fill: 'FFFF00', bold: true, horizontal: 'center', fontColor: fr.color, size: 11 });

      rowNum++;
    });

    // Merge Column G (Status) vertically
    const endFooterRow = rowNum - 1;
    ws.mergeCells(`G${startFooterRow}:G${endFooterRow}`);
    const statusCell = ws.getRow(startFooterRow).getCell(7);
    statusCell.value = status;

    let statusFill = '92D050'; // Default green
    const normalizedStatus = status ? status.trim() : '';
    if (normalizedStatus === 'Credit Settled OK' || normalizedStatus === 'Credit Settled 0K') {
      statusFill = 'C6E0B4';
    } else if (normalizedStatus === 'FOUND Credit Mismatch') {
      statusFill = 'F4B084';
    }

    applyCellStyle(statusCell, {
      fill: statusFill,
      bold: true,
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
      size: 11
    });

    // Apply borders to the footer block
    for (let r = startFooterRow; r <= endFooterRow; r++) {
      const currRow = ws.getRow(r);
      for (let c = 1; c <= 7; c++) {
        const cell = currRow.getCell(c);
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
    }
    rowAfterFooter = endFooterRow + 1;
  }

  if (policy2021Style) {
    const noteText = buildPolicy2021SurplusFooterText(calculationTable);
    ws.mergeCells(`A${rowNum}:G${rowNum}`);
    const noteCell = ws.getCell(`A${rowNum}`);
    noteCell.value = noteText;
    applyCellStyle(noteCell, {
      horizontal: 'left',
      vertical: 'top',
      wrapText: true,
      bold: false,
      size: 10,
      fill: 'FFFFFF',
    });
    ws.getRow(rowNum).height = 78;
    rowNum += 1;
  }

  // Apply requested Borders
  // 1. Top of Row 2
  ws.getRow(2).eachCell(cell => {
    cell.border = { ...cell.border, top: { style: 'medium' } };
  });

  // 2. Right of Column G (index 7)
  ws.getColumn(7).eachCell(cell => {
    cell.border = { ...cell.border, right: { style: 'medium' } };
  });

  // Ensure empty/spacer rows have right borders
  [`G${emptyRowSec1}`, `G${sec3StartRow - 1}`, `G${footerStartRow - 1}`].forEach(cellAddr => {
    const cell = ws.getCell(cellAddr);
    cell.border = { ...cell.border, right: { style: 'medium' } };
  });

  // 3. Bottom of Section 1 spacer (previously row 20)
  if (ws.getRow(emptyRowSec1)) {
    ws.getRow(emptyRowSec1).eachCell(cell => {
      cell.border = { ...cell.border, bottom: { style: 'medium' } };
    });
  }

  // 4. Top of dynamic section rows
  [sec2StartRow, sec3StartRow, footerStartRow, rowAfterFooter].forEach(rowIdx => {
    if (!rowIdx) return;
    const row = ws.getRow(rowIdx);
    for (let c = 1; c <= 7; c++) {
      const cell = row.getCell(c);
      cell.border = { ...cell.border, top: { style: 'medium' } };
    }
  });

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
      policy2021Style,
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

    const workbook = buildUnitCreditExcel(clientName, solarLabel, adjustmentLabel, calculationTable || {}, {
      policy2021Style: policy2021Style === true || policy2021Style === 'true',
    });

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
