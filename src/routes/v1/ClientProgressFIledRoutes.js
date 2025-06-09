const express = require('express');
const router = express.Router();
const clientProgressController = require('../../controllers/v1/ClientProgressFIledController');

// Create client progress
router.post('/', clientProgressController.createClientProgress);

// Get all client progress records
router.get('/', clientProgressController.getAllClientProgress);

// Get client progress by id
router.get('/:id', clientProgressController.getClientProgressById);

// Update client progress by id
router.put('/:id', clientProgressController.updateClientProgress);

// Delete client progress by id
router.delete('/:id', clientProgressController.deleteClientProgress);


// Delete client from progress by clientId
router.patch('/delete-client/:clientId', clientProgressController.deleteClientFromProgress);


module.exports = router;
