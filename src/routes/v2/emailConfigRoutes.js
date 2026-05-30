// emailConfigRoutes.js
const express = require('express');
const path = require('path');
const fs = require('fs');
const router = express.Router();
const multer = require('multer');
const emailConfigController = require('../../controllers/v2/emailConfigController');
const emailDirectoryController = require('../../controllers/v2/emailDirectoryController');
const { verifyToken } = require('../../middleware/v1/authMiddleware');

const directoryUploadDir = 'uploads/email-directory';
if (!fs.existsSync(directoryUploadDir)) {
  fs.mkdirSync(directoryUploadDir, { recursive: true });
}

const directoryStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, directoryUploadDir),
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, uniqueSuffix + path.extname(file.originalname));
  },
});

const directoryUpload = multer({
  storage: directoryStorage,
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    const allowed = ['.csv', '.xlsx', '.xls', '.cdf', '.dlm'];
    if (allowed.includes(ext)) return cb(null, true);
    cb(new Error('Only CSV, CDF/DLM, and Excel files are allowed'));
  },
}).single('file');

// All routes require authentication
router.use(verifyToken);

// Get all clients for dropdown
router.get('/clients', emailConfigController.getAllClients);

// Get email configuration by type (weekly/monthly)
router.get('/config/:configType', emailConfigController.getEmailConfig);

// Update email configuration
router.put('/config/:configType', emailConfigController.updateEmailConfig);

// Add client to configuration
router.post('/config/:configType/client', emailConfigController.addClientToConfig);

// Remove client from configuration
router.delete('/config/:configType/client/:clientId', emailConfigController.removeClientFromConfig);

// Client-specific CC email management
router.post('/config/:configType/client/:clientId/cc', emailConfigController.addCCEmailToClient);
router.delete('/config/:configType/client/:clientId/cc/:email', emailConfigController.removeCCEmailFromClient);
router.put('/config/:configType/client/:clientId/cc/:oldEmail', emailConfigController.updateCCEmailForClient);

// Update client email
router.put('/config/:configType/client/:clientId/email', emailConfigController.updateClientEmail);

// Per-client template (general)
router.put('/config/:configType/client/:clientId/template', emailConfigController.updateClientTemplate);
router.post('/config/:configType/client/:clientId/template/reset', emailConfigController.resetClientTemplate);

// Global CC email management (deprecated but kept for backward compatibility)
router.post('/config/:configType/cc', emailConfigController.addCCEmail);
router.delete('/config/:configType/cc/:email', emailConfigController.removeCCEmail);

// Update template
router.put('/config/:configType/template', emailConfigController.updateTemplate);

// Reset configuration to default template
router.post('/config/:configType/reset', emailConfigController.resetConfigToDefault);

// Email directory (global contact list for To/CC suggestions)
router.get('/directory', emailDirectoryController.getDirectory);
router.get('/directory/search', emailDirectoryController.searchDirectory);
router.put('/directory/entries', emailDirectoryController.saveDirectoryEntries);
router.post('/directory/upload', directoryUpload, emailDirectoryController.uploadDirectory);

module.exports = router;
