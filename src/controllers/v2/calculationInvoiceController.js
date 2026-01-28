// calculationInvoiceController.js
const CalculationInvoice = require('../../models/v2/calculationInvoice.model');
const LossesCalculationData = require('../../models/v1/lossesCalculation.model');
const SubClient = require('../../models/v1/subClient.model');
const Policy = require('../../models/v2/policy.model');
const logger = require('../../utils/logger');
const mongoose = require('mongoose');

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

    // Calculate generation unit and drawl unit
    const generationUnit = subClientData.grossInjectionMWH ? subClientData.grossInjectionMWH * 1000 : null;
    const drawlUnit = subClientData.drawlMWH ? subClientData.drawlMWH * 1000 * (-1) : null;

    logger.info(`Retrieved losses data for sub-client: ${subClientId}, month: ${month}, year: ${year}`);
    res.status(200).json({
      generationUnit,
      drawlUnit,
      grossInjectionMWH: subClientData.grossInjectionMWH,
      drawlMWH: subClientData.drawlMWH,
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
