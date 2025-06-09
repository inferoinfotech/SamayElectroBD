// partClientController.js
const historyModel = require('../../models/v1/history.model');
const PartClient = require('../../models/v1/partClient.model');
const SubClient = require('../../models/v1/subClient.model');
const logger = require('../../utils/logger');
const LossesCalculationData = require('../../models/v1/lossesCalculation.model');
const History = require('../../models/v1/history.model');
// Add a new Part Client
// Add a new PartClient with authentication middleware
exports.addPartClient = async (req, res) => {
    try {
        const { subClientId, sharingPercentage, divisionName, consumerNo, discom } = req.body;

        // Verify SubClient existence
        const subClient = await SubClient.findById(subClientId);
        if (!subClient) {
            logger.warn(`SubClient not found: ${subClientId}`);
            return res.status(404).json({ message: "SubClient not found" });
        }

        // Create PartClient
        const newPartClient = new PartClient({
            sharingPercentage,
            discom,
            divisionName,
            consumerNo,
            subClient: subClient._id,
        });

        await newPartClient.save();

        // Create History entry for addition of PartClient
        const historyEntry = new historyModel({
            clientId: subClient._id,
            clientType: 'part',
            fieldName: 'part client add',
            oldValue: 'not exist',
            newValue: 'create part client',
            changedBy: req.user?._id,
            updatedAt: new Date()
        });

        await historyEntry.save();

        // Attach this history to PartClient history array
        newPartClient.history.push(historyEntry._id);
        await newPartClient.save();

        logger.info(`New PartClient added: ${newPartClient.consumerNo}`);

        res.status(201).json({
            message: "PartClient added successfully",
            partClient: newPartClient,
            history: historyEntry
        });
    } catch (error) {
        logger.error(`Error adding PartClient: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};




// Edit Part Client Field
exports.editPartClientField = async (req, res) => {
    try {
        const { clientId, fieldName, newValue } = req.body;

        // Find the PartClient by ID
        const partClient = await PartClient.findById(clientId);
        if (!partClient) {
            logger.warn(`PartClient not found: ${clientId}`);
            return res.status(404).json({ message: "PartClient not found" });
        }

        // Check if the field exists
        if (!partClient[fieldName] && !fieldName.includes('.')) {
            logger.warn(`Field ${fieldName} does not exist on PartClient: ${clientId}`);
            return res.status(400).json({ message: `Field ${fieldName} does not exist` });
        }

        // Get old value
        const oldValue = fieldName.includes('.')
            ? fieldName.split('.').reduce((o, i) => (o ? o[i] : undefined), partClient.toObject())
            : partClient[fieldName];

        // Create history entry
        const historyEntry = new historyModel({
            clientId: partClient._id,
            clientType: 'part',
            fieldName: fieldName,
            oldValue: oldValue,
            newValue: newValue,
            changedBy: req.user?._id,
            updatedAt: new Date()
        });

        await historyEntry.save();

        // Update using dot notation for nested fields
        const updateObj = { $set: { [fieldName]: newValue } };
        await PartClient.updateOne({ _id: clientId }, updateObj);

        logger.info(`Updated ${fieldName} for PartClient: ${clientId}`);
        res.status(200).json({
            message: `${fieldName} updated successfully`,
            history: historyEntry
        });
    } catch (error) {
        logger.error(`Error updating field on PartClient: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// View PartClient
exports.viewPartClient = async (req, res) => {
    try {
        const partClientId = req.params.clientId;

        // Find PartClient by ID and populate the history field
        const partClient = await PartClient.findById(partClientId)
            .populate('history');  // Populate the history field

        if (!partClient) {
            logger.warn(`Part Client not found: ${partClientId}`);
            return res.status(404).json({ message: "Part Client not found" });
        }
        // Attach history to each PartClient
        const histories = await historyModel.find({ clientId: partClientId });
        // Attach history to each PartClient
        partClient.history = histories.filter(history => history.clientId.toString() === partClientId.toString());


        logger.info(`Retrieved Part Client with history: ${partClientId}`);
        res.status(200).json({ partClient });

    } catch (error) {
        logger.error(`Error retrieving Part Client: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};


// Get All PartClients
exports.getAllPartClients = async (req, res) => {
    try {
        // Find all PartClients and populate the 'history' field
        const partClients = await PartClient.find().populate('history');

        if (partClients.length === 0) {
            logger.warn("No PartClients found.");
            return res.status(404).json({ message: "No PartClients found" });
        }

        // this is partclient data so update this history code for partclients, 
        // Attach history to each PartClient
        const histories = await historyModel.find({ clientId: { $in: partClients.map(client => client._id) } });

        partClients.forEach(client => {
            client.history = histories.filter(history => history.clientId.toString() === client._id.toString());
        });


        logger.info(`Retrieved ${partClients.length} PartClients`);
        res.status(200).json({ partClients });
    } catch (error) {
        logger.error(`Error retrieving all PartClients: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};


// Delete PartClient
exports.deletePartClient = async (req, res) => {
    try {
        const { clientId } = req.params;

        // Find the PartClient by ID
        const partClient = await PartClient.findById(clientId);
        if (!partClient) {
            logger.warn(`PartClient not found: ${clientId}`);
            return res.status(404).json({ message: "PartClient not found" });
        }

        // Step 1: Delete related History entries for the PartClient
        await History.deleteMany({ clientId: partClient._id });
        logger.info(`Deleted History records for PartClient: ${clientId}`);

        // Step 2: Delete related LossesCalculationData for PartClient
        const lossesCalculationData = await LossesCalculationData.updateMany(
            { 'subClient.partclient.subClientId': partClient._id },
            { $pull: { 'subClient.partclient': { subClientId: partClient._id } } }
        );
        if (lossesCalculationData.nModified > 0) {
            logger.info(`Removed PartClient from LossesCalculationData: ${clientId}`);
        }

        // Step 3: Now delete the PartClient
        await PartClient.findByIdAndDelete(clientId);
        logger.info(`Deleted PartClient: ${clientId}`);

        res.status(200).json({ message: "PartClient and related data (History, LossesCalculationData) deleted successfully" });
    } catch (error) {
        logger.error(`Error deleting PartClient: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

