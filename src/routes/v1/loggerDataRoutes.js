const express = require('express');
const router = express.Router();
const loggerDataController = require('../../controllers/v1/loggerDataController');
const upload = require('../../utils/multerConfig'); // Assuming this is where you're handling file uploads
const { verifyToken } = require('../../middleware/v1/authMiddleware');

// Route for uploading a single logger CSV file
router.post('/upload-logger-csv', verifyToken, upload.single('csvFile'), loggerDataController.uploadLoggerCSV);

// Route for getting logger data by month and year
router.get('/get-logger-data/:month/:year', verifyToken, loggerDataController.getLoggerDataByMonthYear);

// Route for getting a specific logger data entry by ID
router.get('/get-logger-data/:id', verifyToken, loggerDataController.getLoggerDataById);

// Route for deleting a logger data entry by ID
router.delete('/delete-logger-data/:loggerDataId', verifyToken, loggerDataController.deleteLoggerData);

module.exports = router;
