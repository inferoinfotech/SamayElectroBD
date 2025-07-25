const express = require('express');
const router = express.Router();
const { generateDailyReport, downloadDailyReportExcel, getLatestDailyReports, downloadDailyReportPDF } = require('../../controllers/v1/dailyReportController');
const { verifyToken } = require('../../middleware/v1/authMiddleware');

// Route for generating the daily report for Main Client
router.post('/generate-daily-report', verifyToken, generateDailyReport);

router.get('/download/:dailyReportId', verifyToken, downloadDailyReportExcel);

router.get('/download/pdf/:dailyReportId', verifyToken, downloadDailyReportPDF);

router.get('/latest', verifyToken, getLatestDailyReports);

module.exports = router;
