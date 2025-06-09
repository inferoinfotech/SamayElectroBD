// subClientRoutes.js
const express = require('express');
const router = express.Router();
const subClientController = require('../../controllers/v1/subClientController');
const { verifyToken } = require('../../middleware/v1/authMiddleware');

// Add SubClient
router.post('/add', verifyToken, subClientController.addSubClient);

// PUT request to edit a field in SubClient
router.put('/edit', subClientController.editSubClientField);

// GET request to view a SubClient and its history
router.get('/:subClientId', verifyToken, subClientController.viewSubClient);

// GET all SubClients
router.get('/', verifyToken, subClientController.getAllSubClients);

// DELETE SubClient
router.delete('/:subClientId', verifyToken, subClientController.deleteSubClient);

module.exports = router;
