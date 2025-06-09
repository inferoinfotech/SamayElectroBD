// partClientRoutes.js

const express = require('express');
const router = express.Router();
const partClientController = require('../../controllers/v1/partClientController');
const { verifyToken } = require('../../middleware/v1/authMiddleware');

// Add PartClient
router.post('/add', verifyToken, partClientController.addPartClient);

// Edit PartClient Field
router.put('/edit', verifyToken, partClientController.editPartClientField);

// View PartClient
router.get('/:clientId', verifyToken, partClientController.viewPartClient);

// Get All PartClients
router.get('/', verifyToken, partClientController.getAllPartClients);

// Delete PartClient
router.delete('/:clientId', verifyToken, partClientController.deletePartClient);

module.exports = router;
