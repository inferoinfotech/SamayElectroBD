// clientPolicyRoutes.js
const express = require('express');
const router = express.Router();
const clientPolicyController = require('../../controllers/v2/clientPolicyController');
const { verifyToken } = require('../../middleware/v1/authMiddleware');

// Assign a policy to a client
router.post('/assign', verifyToken, clientPolicyController.assignPolicyToClient);

// Get all client-policy mappings
router.get('/', verifyToken, clientPolicyController.getAllClientPolicies);

// Get a single client-policy mapping
router.get('/:clientPolicyId', verifyToken, clientPolicyController.getClientPolicyById);

// Get policies for a specific client
router.get('/client/:clientId', verifyToken, clientPolicyController.getClientPolicies);

// Update client-policy mapping
router.put('/:clientPolicyId', verifyToken, clientPolicyController.updateClientPolicy);

// Update a specific sub-policy apply status for a client
router.put('/:clientPolicyId/sub-policy/:subPolicyId', verifyToken, clientPolicyController.updateClientSubPolicy);

// Remove policy assignment from a client
router.delete('/:clientPolicyId', verifyToken, clientPolicyController.removeClientPolicy);

module.exports = router;

