// policyRoutes.js
const express = require('express');
const router = express.Router();
const policyController = require('../../controllers/v2/policyController');
const { verifyToken } = require('../../middleware/v1/authMiddleware');

// Create a new policy
router.post('/create', verifyToken, policyController.createPolicy);

// Get all policies
router.get('/', verifyToken, policyController.getAllPolicies);

// Get a single policy by ID
router.get('/:policyId', verifyToken, policyController.getPolicyById);

// Update a policy
router.put('/:policyId', verifyToken, policyController.updatePolicy);

// Delete a policy
router.delete('/:policyId', verifyToken, policyController.deletePolicy);

// Add a sub-policy to an existing policy
router.post('/:policyId/sub-policy', verifyToken, policyController.addSubPolicy);

// Update a sub-policy
router.put('/:policyId/sub-policy/:subPolicyId', verifyToken, policyController.updateSubPolicy);

// Delete a sub-policy
router.delete('/:policyId/sub-policy/:subPolicyId', verifyToken, policyController.deleteSubPolicy);

module.exports = router;

