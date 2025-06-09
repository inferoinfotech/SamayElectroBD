const ClientProgressField = require('../../models/v1/ClientProgressFIled.model');
const logger = require('../../utils/logger');
const mongoose = require('mongoose');

// Create a new client progress record
exports.createClientProgress = async (req, res) => {
  try {
    const { clients } = req.body;
    if (!clients || !Array.isArray(clients) || clients.length === 0) {
      return res.status(400).json({ message: 'Clients array is required.' });
    }

    const newRecord = new ClientProgressField({ clients });
    await newRecord.save();

    res.status(201).json({ message: 'Client progress data created successfully', data: newRecord });
  } catch (error) {
    logger.error(`Error creating client progress: ${error.message}`);
    res.status(500).json({ message: 'Error creating client progress', error: error.message });
  }
};

// Get client progress by id
exports.getClientProgressById = async (req, res) => {
  try {
    const { id } = req.params;
    const record = await ClientProgressField.findById(id).populate('clients.clientId', 'name');

    if (!record) {
      return res.status(404).json({ message: 'Client progress record not found' });
    }

    res.status(200).json({ data: record });
  } catch (error) {
    logger.error(`Error fetching client progress: ${error.message}`);
    res.status(500).json({ message: 'Error fetching client progress', error: error.message });
  }
};

// Update client progress by id
exports.updateClientProgress = async (req, res) => {
  try {
    const { id } = req.params;
    const { clients } = req.body;

    if (!clients || !Array.isArray(clients) || clients.length === 0) {
      return res.status(400).json({ message: 'Clients array is required for update.' });
    }

    const updated = await ClientProgressField.findByIdAndUpdate(id, { clients }, { new: true });

    if (!updated) {
      return res.status(404).json({ message: 'Client progress record not found' });
    }

    res.status(200).json({ message: 'Client progress updated successfully', data: updated });
  } catch (error) {
    logger.error(`Error updating client progress: ${error.message}`);
    res.status(500).json({ message: 'Error updating client progress', error: error.message });
  }
};

// Delete client progress by id
exports.deleteClientProgress = async (req, res) => {
  try {
    const { id } = req.params;
    const deleted = await ClientProgressField.findByIdAndDelete(id);

    if (!deleted) {
      return res.status(404).json({ message: 'Client progress record not found' });
    }

    res.status(200).json({ message: 'Client progress deleted successfully' });
  } catch (error) {
    logger.error(`Error deleting client progress: ${error.message}`);
    res.status(500).json({ message: 'Error deleting client progress', error: error.message });
  }
};

// Delete a client from the progress record
exports.deleteClientFromProgress = async (req, res) => {
  try {
    const clientId = req.params.clientId;

    if (!mongoose.Types.ObjectId.isValid(clientId)) {
      return res.status(400).json({ message: 'Invalid clientId' });
    }

    // Find the document containing the client and pull the client from the array
    const updatedDoc = await ClientProgressField.findOneAndUpdate(
      { 'clients.clientId': clientId },
      { $pull: { clients: { clientId: clientId } } },
      { new: true }
    );

    if (!updatedDoc) {
      return res.status(404).json({ message: 'Client not found in any progress record' });
    }

    res.status(200).json({
      message: `Client with ID ${clientId} removed successfully from progress record`,
      data: updatedDoc
    });
  } catch (error) {
    console.error('Error deleting client from progress:', error);
    res.status(500).json({
      message: 'Error deleting client from progress',
      error: error.message
    });
  }
};
// Get all client progress records (without filtering by id)
exports.getAllClientProgress = async (req, res) => {
  try {
    const records = await ClientProgressField.find().populate('clients.clientId', 'name');

    if (!records || records.length === 0) {
      return res.status(404).json({ message: 'No client progress records found' });
    }

    res.status(200).json({ data: records });
  } catch (error) {
    logger.error(`Error fetching all client progress records: ${error.message}`);
    res.status(500).json({ message: 'Error fetching client progress records', error: error.message });
  }
};
