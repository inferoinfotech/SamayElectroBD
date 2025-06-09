// /src/routes/v1/clientProgressRoutes.js
const express = require('express');
const router = express.Router();
const clientProgressController = require('../../controllers/v1/clientProgressController');

// Create a new client progress entry
router.post('/', clientProgressController.createClientProgress);

// Update a client progress entry by ID
router.put('/:id', clientProgressController.updateClientProgress);

// Get client progress entries by month and year
router.get('/:month/:year', clientProgressController.getClientProgressByMonthYear);

router.get('/download-excel/:month/:year', clientProgressController.downloadClientProgressExcel);

module.exports = router;
