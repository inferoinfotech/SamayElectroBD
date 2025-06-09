const express = require('express');
const router = express.Router();
const { verifyToken } = require('../../middleware/v1/authMiddleware');
const { getLossesCalculationData } = require('../../controllers/v1/totalDataYearilyController');

router.post('/data',verifyToken,getLossesCalculationData);

module.exports = router;