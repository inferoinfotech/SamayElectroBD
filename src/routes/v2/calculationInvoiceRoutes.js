// calculationInvoiceRoutes.js
const express = require('express');
const router = express.Router();
const calculationInvoiceController = require('../../controllers/v2/calculationInvoiceController');
const { verifyToken } = require('../../middleware/v1/authMiddleware');

// Get losses data for sub-client (for auto-filling generation and drawl units)
router.get('/losses-data', verifyToken, calculationInvoiceController.getLossesDataForSubClient);

// Create a new calculation invoice
router.post('/create', verifyToken, calculationInvoiceController.createCalculationInvoice);

// Export Unit Credit Calculation to Excel (POST with same payload as create)
router.post('/export-excel', verifyToken, calculationInvoiceController.exportCalculationToExcel);

// Get all calculation invoices
router.get('/', verifyToken, calculationInvoiceController.getAllCalculationInvoices);

// Get calculation invoice by criteria (subClientId, policyId, dates)
router.get('/by-criteria', verifyToken, calculationInvoiceController.getCalculationInvoiceByCriteria);

// Get calculation invoice by solar period only (ignores adjustment billing; latest match by updatedAt)
router.get('/by-solar-period', verifyToken, calculationInvoiceController.getCalculationInvoiceBySolarPeriod);

// Get a single calculation invoice by ID
router.get('/:invoiceId', verifyToken, calculationInvoiceController.getCalculationInvoiceById);

// Update a calculation invoice
router.put('/:invoiceId', verifyToken, calculationInvoiceController.updateCalculationInvoice);

// Delete a calculation invoice
router.delete('/:invoiceId', verifyToken, calculationInvoiceController.deleteCalculationInvoice);

module.exports = router;

