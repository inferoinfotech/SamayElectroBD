const express = require('express');
const router = express.Router();
const formatDataController = require('../../controllers/v2/formatDataController');

// Define routes
router.post('/upload', formatDataController.uploadFormatData);
router.post('/upload-auto', formatDataController.uploadFormatDataAuto); // Auto-detect format/month/year
router.post('/validate-bulk', formatDataController.validateBulkFiles);
router.post('/upload-bulk', formatDataController.uploadBulkFormatData);
router.post('/convert-bulk', formatDataController.convertBulkAndDownloadZip); // Bulk convert + zip download
router.get('/download', formatDataController.downloadFormatDataCSV);
router.get('/check-period', formatDataController.checkPeriodAvailability);
router.get('/download-zip', formatDataController.downloadFormatDataZip);
router.get('/all-uploads', formatDataController.getAllUploadedFormats);
router.get('/meters', formatDataController.getAvailableMeters);
router.get('/debug', formatDataController.debugFormatData);
router.delete('/delete', formatDataController.deleteFormatData);

module.exports = router;
