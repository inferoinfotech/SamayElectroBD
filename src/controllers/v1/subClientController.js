// subClientController.js
const SubClient = require('../../models/v1/subClient.model');
const MainClient = require('../../models/v1/mainClient.model'); // Import MainClient to check the connection
const History = require('../../models/v1/history.model');
const logger = require('../../utils/logger');
const historyModel = require('../../models/v1/history.model');
const MeterData = require('../../models/v1/meterData.model');
const LossesCalculationData = require('../../models/v1/lossesCalculation.model');
const LoggerData = require('../../models/v1/loggerData.model');
const Dailyreport = require('../../models/v1/dailyReport.model');
const PartClient = require('../../models/v1/partClient.model');

// Add SubClient
// Sub Client Add API
exports.addSubClient = async (req, res) => {
    const requiredFields = [
        'name', 'abtMainMeter', 'mainClient'
    ];

    // Check if required fields exist in the request body
    const missingFields = requiredFields.filter(field => !req.body[field]);
    if (missingFields.length) {
        return res.status(400).json({ message: `Missing required fields: ${missingFields.join(', ')}` });
    }

    const {
        name,
        divisionName,
        consumerNo,
        // modemSrNo,
        abtMainMeter,
        abtCheckMeter,
        voltageLevel,
        ctptSrNo,
        ctRatio,
        ptRatio,
        mf,
        pn,
        acCapacityKw,
        dcCapacityKwp,
        dcAcRatio,
        noOfModules,
        moduleCapacityWp,
        inverterCapacityKw,
        numberOfInverters,
        makeOfInverter,
        sharingPercentage,
        contactNo,
        email,
        discom,
        REtype,
        mainClient // mainClient ID passed here
    } = req.body;

    try {
        // Validate the existence of the Main Client reference
        const client = await MainClient.findById(mainClient);
        if (!client) {
            return res.status(404).json({ message: "Main Client not found" });
        }

        // Create and populate the new SubClient object
        const newSubClient = new SubClient({
            name,
            divisionName,
            consumerNo,
            // modemSrNo,
            abtMainMeter,
            abtCheckMeter,
            voltageLevel,
            ctptSrNo,
            ctRatio,
            ptRatio,
            mf,
            pn,
            acCapacityKw,
            dcCapacityKwp,
            dcAcRatio,
            noOfModules,
            moduleCapacityWp,
            inverterCapacityKw,
            numberOfInverters,
            makeOfInverter,
            sharingPercentage,
            discom,
            contactNo,
            email,
            REtype,
            mainClient: client._id // Ensure referencing the correct Main Client
        });

        // Save the new SubClient to the database
        await newSubClient.save();

        // Create History entry for addition of SubClient
        const historyEntry = new History({
            clientId: newSubClient._id,
            fieldName: 'sub client add',
            oldValue: 'not exist',
            newValue: 'create sub client',
            updatedAt: new Date()
        });
        await historyEntry.save();

        // Attach this history to SubClient's history array and save again
        newSubClient.history.push(historyEntry._id);
        await newSubClient.save();

        return res.status(201).json({
            message: 'Sub Client added successfully',
            subClient: newSubClient,
            history: historyEntry
        });

    } catch (error) {
        // Improved error handling for any unexpected issues
        console.error('Error in adding SubClient:', error);
        return res.status(500).json({ message: 'An error occurred while adding the sub-client', error: error.message });
    }
};



// Update SubClient Field
// In your subClientController.js
exports.editSubClientField = async (req, res) => {
    try {
        const { clientId, fieldName, newValue } = req.body;

        // Find the SubClient by ID
        const subClient = await SubClient.findById(clientId);
        if (!subClient) {
            logger.warn(`SubClient not found: ${clientId}`);
            return res.status(404).json({ message: "SubClient not found" });
        }

        // Convert to plain object to check nested paths
        const subClientObj = subClient.toObject();

        // Helper function to check if path exists
        const pathExists = (obj, path) => {
            return path.split('.').reduce((o, p) => (o && o[p] !== undefined) ? o[p] : undefined, obj) !== undefined;
        };

        // Check if the field exists (including nested)
        if (!pathExists(subClientObj, fieldName)) {
            logger.warn(`Field ${fieldName} does not exist on SubClient: ${clientId}`);
            return res.status(400).json({ message: `Field ${fieldName} does not exist` });
        }

        // Get old value
        const oldValue = fieldName.split('.').reduce((o, i) => (o ? o[i] : undefined), subClientObj);

        // Create history entry
        const historyEntry = new historyModel({
            clientId: subClient._id,
            clientType: 'sub',
            fieldName: fieldName,
            oldValue: oldValue,
            newValue: newValue,
            changedBy: req.user?._id,
            updatedAt: new Date()
        });

        await historyEntry.save();

        // Update using dot notation for nested fields
        const updateObj = { $set: { [fieldName]: newValue } };
        await SubClient.updateOne({ _id: clientId }, updateObj);

        logger.info(`Updated ${fieldName} for SubClient: ${clientId}`);
        res.status(200).json({
            message: `${fieldName} updated successfully`,
            history: historyEntry
        });
    } catch (error) {
        logger.error(`Error updating field on SubClient: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};


// View SubClient (with History)
exports.viewSubClient = async (req, res) => {
    try {
        const subClientId = req.params.subClientId;

        const subClient = await SubClient.findById(subClientId).lean();
        if (!subClient) {
            logger.warn(`SubClient not found: ${subClientId}`);
            return res.status(404).json({ message: "SubClient not found" });
        }

        // Get history
        const history = await History.find({ clientId: subClientId });

        // Include the history in the response
        subClient.history = history;

        logger.info(`Retrieved SubClient with history: ${subClientId}`);
        res.status(200).json({ subClient });
    } catch (error) {
        logger.error(`Error retrieving SubClient: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Get All SubClients
exports.getAllSubClients = async (req, res) => {
    try {
        const subClients = await SubClient.find().populate('mainClient'); // Optionally, populate MainClient details

        if (!subClients.length) {
            logger.warn('No SubClients found');
            return res.status(404).json({ message: "No SubClients found" });
        }

        // i want history of all subclients
        const histories = await History.find({ clientId: { $in: subClients.map(client => client._id) } });
        // Attach history to each subClient
        subClients.forEach(client => {
            client.history = histories.filter(history => history.clientId.toString() === client._id.toString());
        });
        
        logger.info('Retrieved all SubClients');
        res.status(200).json({ subClients });
    } catch (error) {
        logger.error(`Error retrieving SubClients: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Delete SubClient
exports.deleteSubClient = async (req, res) => {
    try {
        const subClientId = req.params.subClientId;

        // Find the SubClient by ID
        const subClient = await SubClient.findById(subClientId);
        if (!subClient) {
            logger.warn(`SubClient not found: ${subClientId}`);
            return res.status(404).json({ message: "SubClient not found" });
        }

        // Step 1: Delete related History entries for SubClient
        await History.deleteMany({ clientId: subClientId }); // Deleting history for SubClient
        logger.info(`Deleted History records for SubClient: ${subClientId}`);

        // Step 2: Delete related Meter Data for SubClient
        const subClientMeterData = await MeterData.deleteMany({ client: subClientId, clientType: 'SubClient' });
        if (subClientMeterData.deletedCount > 0) {
            logger.info(`Deleted ${subClientMeterData.deletedCount} MeterData related to SubClient: ${subClientId}`);
        }

        // Step 3: Delete related LossesCalculationData for SubClient
        const lossesCalculationData = await LossesCalculationData.deleteMany({ 'subClient.subClientId': subClientId });
        if (lossesCalculationData.deletedCount > 0) {
            logger.info(`Deleted ${lossesCalculationData.deletedCount} LossesCalculationData related to SubClient: ${subClientId}`);
        }

        // Step 4: Delete related LoggerData for SubClient
        const loggerData = await LoggerData.deleteMany({ 'subClient.subClientId': subClientId });
        if (loggerData.deletedCount > 0) {
            logger.info(`Deleted ${loggerData.deletedCount} LoggerData related to SubClient: ${subClientId}`);
        }

        // Step 5: Delete related DailyReport for SubClient
        const dailyReport = await Dailyreport.deleteMany({ 'subClient.subClientId': subClientId });
        if (dailyReport.deletedCount > 0) {
            logger.info(`Deleted DailyReport related to SubClient: ${subClientId}`);
        }

        // Step 6: Delete related PartClients for SubClient
        const partClients = await PartClient.find({ subClient: subClientId });
        if (partClients.length > 0) {
            // Delete related PartClients
            await PartClient.deleteMany({ subClient: subClientId });
            logger.info(`Deleted ${partClients.length} PartClients related to SubClient: ${subClientId}`);
        }

        // Step 7: Now delete the SubClient
        await SubClient.findByIdAndDelete(subClientId);
        logger.info(`SubClient deleted: ${subClientId}`);

        res.status(200).json({ message: "SubClient and related data (PartClients, MeterData, LossesCalculationData, LoggerData, DailyReports, History) deleted successfully" });
    } catch (error) {
        logger.error(`Error deleting SubClient: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

