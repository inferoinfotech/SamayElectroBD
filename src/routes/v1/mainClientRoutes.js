// mainClientRoutes.js
const express = require('express');
const router = express.Router();
const mainClientController = require('../../controllers/v1/mainClientController');
const { verifyToken } = require('../../middleware/v1/authMiddleware');

// Add Main Client (only one user should access this)
router.post('/add', verifyToken, mainClientController.addMainClient);

// Edit Main Client Field (only one user should access this)
router.put('/edit', verifyToken, mainClientController.editMainClientField);

// Get All Main Clients (only one user should access this)
router.get('/', verifyToken, mainClientController.getAllMainClients);

// Get Single Main Client (only one user should access this)
router.get('/:clientId', verifyToken, mainClientController.viewMainClient);

// Get complete hierarchy (main client + sub clients + part clients)
router.get('/hierarchy/:mainClientId', verifyToken, mainClientController.getMainClientHierarchy);

// Delete Main Client (only one user should access this)
router.delete('/:clientId', verifyToken, mainClientController.deleteMainClient);

module.exports = router;
