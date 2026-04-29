// emailSendRoutes.js
const express = require('express');
const router = express.Router();
const emailSendController = require('../../controllers/v2/emailSendController');
const { verifyToken } = require('../../middleware/v1/authMiddleware');
const multer = require('multer');
const path = require('path');

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/email-attachments/');
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, uniqueSuffix + '-' + file.originalname);
    }
});

const upload = multer({
    storage: storage,
    limits: { fileSize: 10 * 1024 * 1024 }, // 10MB limit
    fileFilter: (req, file, cb) => {
        const allowedTypes = /csv|cdf|xlsx|xls/;
        const extname = allowedTypes.test(path.extname(file.originalname).toLowerCase());
        const mimetype = allowedTypes.test(file.mimetype) || 
                        file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                        file.mimetype === 'application/vnd.ms-excel';

        if (mimetype && extname) {
            return cb(null, true);
        } else {
            cb(new Error('Only CSV, CDF, and Excel files are allowed for weekly/monthly emails'));
        }
    }
});

// Multer for general email attachments (accepts ALL file types)
const uploadAny = multer({
    storage: storage,
    limits: { fileSize: 10 * 1024 * 1024 }, // 10MB limit
    // No fileFilter - accepts all file types
});

// All routes require authentication
router.use(verifyToken);

// Process uploaded CSV/CDF files and create batch
router.post('/process-files', upload.array('files', 20), emailSendController.processEmailFiles);

// Upload files for general email attachments (accepts ALL file types)
router.post('/upload-files', uploadAny.array('files', 10), emailSendController.uploadEmailFiles);

// Create new email batch (legacy)
router.post('/batch', emailSendController.createEmailBatch);

// Get all email batches
router.get('/batches', emailSendController.getAllEmailBatches);

// Get single email batch
router.get('/batch/:batchId', emailSendController.getEmailBatchById);

// Get batch recipients for table display
router.get('/batch/:batchId/recipients', emailSendController.getBatchRecipients);

// Update recipient status (CSV check, CDF check)
router.patch('/batch/:batchId/recipient/:recipientId/status', emailSendController.updateRecipientStatus);

// Send email to single recipient
router.post('/batch/:batchId/recipient/:recipientId/send', emailSendController.sendEmailToRecipient);

// Send email to all recipients
router.post('/batch/:batchId/send-all', emailSendController.sendEmailToAll);

// Delete email batch
router.delete('/batch/:batchId', emailSendController.deleteEmailBatch);

// Send general email (for General Monthly Email tab)
router.post('/send-general', emailSendController.sendGeneralEmail);

// General Email Files Management
router.post('/general-files/save', uploadAny.array('files', 10), emailSendController.saveGeneralEmailFiles);
router.get('/general-files', emailSendController.getGeneralEmailFiles);
router.get('/general-files/counts', emailSendController.getAllClientsFileCounts);
router.get('/general-files/download', emailSendController.downloadGeneralEmailFile);
router.post('/general-files/email-sent', emailSendController.updateGeneralEmailSent);
router.delete('/general-files/delete', emailSendController.deleteGeneralEmailFile);

module.exports = router;
