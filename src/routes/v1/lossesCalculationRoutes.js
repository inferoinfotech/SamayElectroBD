const express = require('express');
const router = express.Router();
const lossesCalculationController = require('../../controllers/v1/lossesCalculationController');
const { verifyToken } = require('../../middleware/v1/authMiddleware');

// Route for generating Losses Calculation
router.post('/generate-losses-calculation', verifyToken, lossesCalculationController.generateLossesCalculation);

router.get('/latest',verifyToken, lossesCalculationController.getLatestLossesReports);

// Route for getting Losses Calculation by ID
router.get('/download-losses-calculation/:id', verifyToken, lossesCalculationController.downloadLossesCalculationExcel);

router.post('/get-losses-data-last-four-months', verifyToken, lossesCalculationController.getLossesDataLastFourMonths);

router.post('/get-sldc-data', verifyToken, lossesCalculationController.getSLDCData);
module.exports = router;