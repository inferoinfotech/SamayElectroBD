const { default: mongoose } = require("mongoose");
const historyModel = require("../../models/v1/history.model.js");
const MainClient = require("../../models/v1/mainClient.model.js");
const SubClient = require('../../models/v1/subClient.model');
const PartClient = require('../../models/v1/partClient.model');
const logger = require("../../utils/logger");
const MeterData = require("../../models/v1/meterData.model.js");
const LossesCalculationData = require("../../models/v1/lossesCalculation.model.js");
const LoggerData = require("../../models/v1/loggerData.model.js");
const Dailyreport = require("../../models/v1/dailyReport.model.js");
const Totalreport = require("../../models/v1/totalReport.model.js");
const History = require("../../models/v1/history.model.js");

// Add Main Client
exports.addMainClient = async (req, res) => {
  try {
    const newClientData = req.body;
    const newClient = new MainClient(newClientData);
    await newClient.save();

    logger.info(`New Main Client added: ${newClient.name}`);
    res
      .status(201)
      .json({ message: "Main Client added successfully", client: newClient });
  } catch (error) {
    logger.error(`Error adding Main Client: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Edit Main Client Field (with enhanced nested field support)
exports.editMainClientField = async (req, res) => {
    try {
      const { clientId, fieldName, newValue } = req.body;
  
      // Enhanced validation
      if (!mongoose.Types.ObjectId.isValid(clientId)) {
        return res.status(400).json({ message: "Invalid client ID format" });
      }
  
      if (!fieldName || newValue === undefined) {
        return res.status(400).json({
          message: "Missing required fields",
          required: ["fieldName", "newValue"],
          received: Object.keys(req.body),
        });
      }
  
      // Find and validate the client exists
      const client = await MainClient.findById(clientId);
      if (!client) {
        return res.status(404).json({
          message: "Main Client not found",
          clientId,
          suggestion: "Verify the client ID and try again",
        });
      }
  
      // Special handling for unique meter numbers
      if (
        fieldName === "abtMainMeter.meterNumber" ||
        fieldName === "abtCheckMeter.meterNumber"
      ) {
        const existingClient = await MainClient.findOne({
          $or: [
            { "abtMainMeter.meterNumber": newValue },
            { "abtCheckMeter.meterNumber": newValue },
          ],
          _id: { $ne: clientId },
        });
  
        if (existingClient) {
          return res.status(409).json({
            message: "Meter number must be unique",
            conflictWith: existingClient._id,
            existingValue: newValue,
            suggestion: "Choose a different meter number",
          });
        }
      }
  
      // Get old value for history before updating
      const oldValue = fieldName
        .split(".")
        .reduce(
          (obj, key) => (obj && obj[key] !== undefined ? obj[key] : null),
          client.toObject()
        );
  
      // Create history record first (in case update fails)
      const historyEntry = new historyModel({
        clientId: client._id,
        fieldName,
        oldValue: oldValue !== null ? oldValue.toString() : "null",
        newValue: newValue.toString(),
        updatedAt: new Date(),
      });
      await historyEntry.save();
  
      // Update using findByIdAndUpdate for better atomicity
      const updatedClient = await MainClient.findByIdAndUpdate(
        clientId,
        { $set: { [fieldName]: newValue } },
        { new: true, runValidators: true }
      );
  
      if (!updatedClient) {
        throw new Error("Failed to update client");
      }
  
      return res.status(200).json({
        success: true,
        message: `${fieldName} updated successfully`,
        client: updatedClient,
        changes: {
          field: fieldName,
          oldValue,
          newValue,
        },
      });
    } catch (error) {
      console.error(`Error updating field ${req.body.fieldName}:`, error);
      return res.status(500).json({
        success: false,
        message: error.message,
        field: req.body.fieldName, // Fixed reference error here
        suggestion: "Check server logs for more details",
        ...(process.env.NODE_ENV === "development" && { stack: error.stack }),
      });
    }
  };

// View Main Client (with History)
exports.viewMainClient = async (req, res) => {
  try {
    const clientId = req.params.clientId;

    // Find the main client and populate its history collection
    const client = await MainClient.findById(clientId).lean();

    if (!client) {
      logger.warn(`Main Client not found: ${clientId}`);
      return res.status(404).json({ message: "Main Client not found" });
    }

    // Get the history data from the History collection
    const history = await historyModel.find({ clientId: clientId });

    // Add the history to the response object
    client.history = history;

    logger.info(`Retrieved Main Client with history: ${clientId}`);
    res.status(200).json({ client });
  } catch (error) {
    logger.error(`Error retrieving Main Client: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Get All Main Clients with History
exports.getAllMainClients = async (req, res) => {
  try {
    // Find all main clients in the database
    const clients = await MainClient.find().lean(); // Using .lean() to get plain JavaScript objects

    if (!clients || clients.length === 0) {
      logger.warn("No Main Clients found");
      return res.status(404).json({ message: "No RE Generator added yet..." });
    }

    // Retrieve the history for each client and attach it to the client
    const clientsWithHistory = await Promise.all(
      clients.map(async (client) => {
        const history = await historyModel.find({ clientId: client._id }).lean();
        client.history = history; // Add history to the client object
        return client;
      })
    );

    logger.info(`Retrieved all Main Clients with history`);
    res.status(200).json({ clients: clientsWithHistory });
  } catch (error) {
    logger.error(`Error retrieving all Main Clients with history: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Delete Main Client
exports.deleteMainClient = async (req, res) => {
  try {
      const { clientId } = req.params;

      // Find the main client by ID
      const client = await MainClient.findById(clientId);
      if (!client) {
          logger.warn(`Main Client not found: ${clientId}`);
          return res.status(404).json({ message: "Main Client not found" });
      }

      // Step 1: Delete related History entries for MainClient, SubClient, and PartClient
      await History.deleteMany({ clientId }); // Deleting history for MainClient
      logger.info(`Deleted History records for Main Client: ${clientId}`);

      // Delete history related to SubClients
      const subClients = await SubClient.find({ mainClient: clientId });
      for (const subClient of subClients) {
          await History.deleteMany({ clientId: subClient._id }); // Deleting history for SubClient
          logger.info(`Deleted History records for SubClient: ${subClient._id}`);
      }

      // Delete history related to PartClients
      for (const subClient of subClients) {
          const partClients = await PartClient.find({ subClient: subClient._id });
          for (const partClient of partClients) {
              await History.deleteMany({ clientId: partClient._id }); // Deleting history for PartClient
              logger.info(`Deleted History records for PartClient: ${partClient._id}`);
          }
      }

      // Step 2: Delete related Meter Data for Main Client
      const mainClientMeterData = await MeterData.deleteMany({ client: clientId, clientType: 'MainClient' });
      if (mainClientMeterData.deletedCount > 0) {
          logger.info(`Deleted ${mainClientMeterData.deletedCount} MeterData related to Main Client: ${clientId}`);
      }

      // Step 3: Find and delete all sub-clients related to this main client
      if (subClients.length > 0) {
          // Loop through each sub-client and delete related PartClients and Meter Data
          for (const subClient of subClients) {
              // Delete related PartClients
              const partClients = await PartClient.find({ subClient: subClient._id });
              if (partClients.length > 0) {
                  await PartClient.deleteMany({ subClient: subClient._id });
                  logger.info(`Deleted ${partClients.length} PartClients related to SubClient: ${subClient._id}`);
              }

              // Delete Meter Data for SubClient
              const subClientMeterData = await MeterData.deleteMany({ client: subClient._id, clientType: 'SubClient' });
              if (subClientMeterData.deletedCount > 0) {
                  logger.info(`Deleted ${subClientMeterData.deletedCount} MeterData related to SubClient: ${subClient._id}`);
              }

              // Step 4: Delete LoggerData related to SubClient
              const loggerData = await LoggerData.deleteMany({ 'subClient.subClientId': subClient._id });
              if (loggerData.deletedCount > 0) {
                  logger.info(`Deleted ${loggerData.deletedCount} LoggerData related to SubClient: ${subClient._id}`);
              }

              // Step 5: Delete DailyReport related to SubClient
              await Dailyreport.deleteMany({ 'subClient.subClientId': subClient._id });
              logger.info(`Deleted DailyReport related to SubClient: ${subClient._id}`);
          }

          // Delete all related sub-clients
          await SubClient.deleteMany({ mainClient: clientId });
          logger.info(`Deleted ${subClients.length} SubClients related to Main Client: ${clientId}`);
      }

      // Step 6: Delete related LossesCalculationData for Main Client
      const lossesCalculationData = await LossesCalculationData.deleteMany({ mainClientId: clientId });
      if (lossesCalculationData.deletedCount > 0) {
          logger.info(`Deleted ${lossesCalculationData.deletedCount} LossesCalculationData related to Main Client: ${clientId}`);
      }

      // Step 7: Delete related TotalReport for Main Client
      const totalReport = await Totalreport.deleteMany({ 'clients.mainClientId': clientId });
      if (totalReport.deletedCount > 0) {
          logger.info(`Deleted TotalReport related to Main Client: ${clientId}`);
      }

      // Step 8: Delete related LoggerData for Main Client (subClient data in LoggerData)
      const mainClientLoggerData = await LoggerData.deleteMany({ 'subClient.mainClientId': clientId });
      if (mainClientLoggerData.deletedCount > 0) {
          logger.info(`Deleted ${mainClientLoggerData.deletedCount} LoggerData related to Main Client: ${clientId}`);
      }

      // Step 9: Now delete the main client
      await MainClient.findByIdAndDelete(clientId);
      logger.info(`Main Client deleted: ${clientId}`);

      res.status(200).json({ message: "Main Client, related SubClients, PartClients, MeterData, LossesCalculationData, TotalReport, LoggerData, and History deleted successfully" });
  } catch (error) {
      logger.error(`Error deleting Main Client: ${error.message}`);
      res.status(500).json({ message: error.message });
  }
};



// Get all sub-clients and part-clients for a main client
exports.getMainClientHierarchy = async (req, res) => {
  try {
    const mainClientId = req.params.mainClientId;

    // Validate the ID format
    if (!mongoose.Types.ObjectId.isValid(mainClientId)) {
      return res.status(400).json({ message: "Invalid client ID format" });
    }

    const result = await MainClient.aggregate([
      { $match: { _id: new mongoose.Types.ObjectId(mainClientId) } },
      {
        $lookup: {
          from: "subclients",
          localField: "_id",
          foreignField: "mainClient",
          as: "subClients",
        },
      },
      {
        $lookup: {
          from: "partclients",
          localField: "subClients._id",
          foreignField: "subClient",
          as: "partClients",
        },
      },
    ]);

    if (!result.length) {
      logger.warn(`Main Client not found: ${mainClientId}`);
      return res.status(404).json({ message: "Main Client not found" });
    }

    logger.info(`Retrieved hierarchy for Main Client: ${mainClientId}`);
    res.status(200).json(result[0]);
  } catch (error) {
    logger.error(`Error retrieving client hierarchy: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};
