const express = require('express');
const router = express.Router();
const { verifyToken } = require('../../middleware/v1/authMiddleware');
const { generateTotalReport, downloadTotalReportExcel, downloadTotalReportPDF, getLatestTotalReports } = require('../../controllers/v1/totalReportController');

// Route for generating the total report for multiple Main Clients
router.post('/generate-total-report', verifyToken, generateTotalReport);

// New route for getting latest reports
router.get('/latest', verifyToken, getLatestTotalReports);
// Download Total Report as Excel
router.get('/download/:totalReportId', verifyToken, downloadTotalReportExcel);
router.get('/downloadPdf/:totalReportId', verifyToken, downloadTotalReportPDF);

module.exports = router;