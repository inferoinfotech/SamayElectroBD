const express = require('express');
const router = express.Router();
const meterDataController = require('../../controllers/v1/meterDataController');
const upload = require('../../utils/multerConfig');
const { verifyToken } = require('../../middleware/v1/authMiddleware');

// Route for uploading CSV files
router.post('/upload-csv', verifyToken, upload.array('csvFile', 50), meterDataController.uploadMeterCSV);

// Route for showing meter data
router.post('/show-meter-data', verifyToken, meterDataController.showMeterData);
// Route for deleting meter data by meterDataId
router.delete('/delete-meter-data/:meterDataId', verifyToken, meterDataController.deleteMeterData);


module.exports = router;
