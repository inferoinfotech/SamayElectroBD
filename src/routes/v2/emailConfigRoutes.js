// emailConfigRoutes.js
const express = require('express');
const router = express.Router();
const emailConfigController = require('../../controllers/v2/emailConfigController');
const { verifyToken } = require('../../middleware/v1/authMiddleware');

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

// Global CC email management (deprecated but kept for backward compatibility)
router.post('/config/:configType/cc', emailConfigController.addCCEmail);
router.delete('/config/:configType/cc/:email', emailConfigController.removeCCEmail);

// Update template
router.put('/config/:configType/template', emailConfigController.updateTemplate);

// Reset configuration to default template
router.post('/config/:configType/reset', emailConfigController.resetConfigToDefault);

module.exports = router;
