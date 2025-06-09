const fs = require('fs');
const path = require('path');
const csv = require('csv-parser');
const logger = require('../../utils/logger');
const MeterData = require('../../models/v1/meterData.model');
const MainClient = require('../../models/v1/mainClient.model');
const SubClient = require('../../models/v1/subClient.model');
const moment = require('moment');
const LossesCalculationData = require('../../models/v1/lossesCalculation.model');
const DailyReport = require('../../models/v1/dailyReport.model');
const TotalReport = require('../../models/v1/totalReport.model');

// Helper function: Extract Meter Number
const extractMeterNumber = (fileName) => {
    const match = fileName.match(/Load Survey - (\w+)/);
    return match ? match[1].trim() : null;
};

// Helper function: Determine client and meter type
const findClientByMeterNumber = async (meterNumber) => {
    let client, meterType;

    if (meterNumber.startsWith('GJ')) {
        client = await MainClient.findOne({ 'abtMainMeter.meterNumber': meterNumber });
        meterType = client ? 'abtMainMeter' : null;

        if (!client) {
            client = await MainClient.findOne({ 'abtCheckMeter.meterNumber': meterNumber });
            meterType = client ? 'abtCheckMeter' : null;
        }

        return client ? { client, clientType: 'MainClient', meterType } : null;
    } else if (meterNumber.startsWith('DG')) {
        client = await SubClient.findOne({ 'abtMainMeter.meterNumber': meterNumber });
        meterType = client ? 'abtMainMeter' : null;

        if (!client) {
            client = await SubClient.findOne({ 'abtCheckMeter.meterNumber': meterNumber });
            meterType = client ? 'abtCheckMeter' : null;
        }

        return client ? { client, clientType: 'SubClient', meterType } : null;
    }

    return null;
};

exports.uploadMeterCSV = async (req, res) => {
    try {
        const { month, year } = req.body;
        const files = req.files;

        if (!files || files.length === 0) {
            return res.status(400).json({ message: "No files uploaded." });
        }

        const successfulFiles = [];
        const invalidFiles = [];

        for (const file of files) {
            const filePath = path.resolve(file.path);
            const meterNumber = extractMeterNumber(file.originalname);

            if (!meterNumber) {
                fs.unlinkSync(filePath);  // Delete invalid file immediately
                invalidFiles.push({
                    fileName: file.originalname,
                    reason: "Invalid file name format - could not extract meter number"
                });
                continue;
            }

            const clientData = await findClientByMeterNumber(meterNumber);
            if (!clientData) {
                fs.unlinkSync(filePath);  // Delete invalid file immediately
                invalidFiles.push({
                    fileName: file.originalname,
                    reason: `Client data not found for meter number ${meterNumber}`
                });
                continue;
            }

            const existingData = await MeterData.findOne({ meterNumber, month, year });
            if (existingData) {
                fs.unlinkSync(filePath);  // Delete invalid file immediately
                invalidFiles.push({
                    fileName: file.originalname,
                    reason: `Data already exists for meter ${meterNumber} for ${month}/${year}`
                });
                continue;
            }

            const startDate = moment(`${year}-${month}-01`, "YYYY-MM-DD");
            const endDate = startDate.clone().add(1, 'month');

            const dataEntries = [];
            let lineNumber = 0;

            try {
                await new Promise((resolve, reject) => {
                    fs.createReadStream(filePath)
                        .pipe(csv({ skipLines: 7 }))
                        .on('data', (data) => {
                            lineNumber++;
                            if (lineNumber === 1 && !data.Date) return;  // Skip any extra header lines

                            const entryDate = moment(data['Date'], "DD/MM/YYYY");
                            if (entryDate.isValid() && entryDate.isSameOrAfter(startDate) && entryDate.isBefore(endDate)) {
                                dataEntries.push({
                                    date: entryDate.toDate(),
                                    intervalStart: data['Interval Start'],
                                    intervalEnd: data['Interval End'],
                                    parameters: data
                                });
                            }
                        })
                        .on('end', () => {
                            fs.unlinkSync(filePath); // Delete processed file
                            resolve();
                        })
                        .on('error', (error) => {
                            fs.unlinkSync(filePath); // Delete invalid file immediately
                            logger.error(`CSV parse error in file ${file.originalname}: ${error.message}`);
                            reject(error);
                        });
                });

                if (dataEntries.length > 0) {
                    const meterData = new MeterData({
                        meterNumber,
                        clientType: clientData.clientType,
                        meterType: clientData.meterType,
                        client: clientData.client._id,
                        month,
                        year,
                        dataEntries
                    });

                    await meterData.save();
                    successfulFiles.push({
                        fileName: file.originalname,
                        meterNumber,
                        message: `Successfully processed ${dataEntries.length} records`,
                        meterDataId: meterData._id
                    });
                    logger.info(`CSV data for ${meterNumber} uploaded successfully for ${month}/${year}.`);
                } else {
                    invalidFiles.push({
                        fileName: file.originalname,
                        reason: "No valid data entries found within the specified month/year"
                    });
                }
            } catch (error) {
                invalidFiles.push({
                    fileName: file.originalname,
                    reason: `CSV processing error: ${error.message}`
                });
                continue;
            }
        }

        const response = {
            message: "File processing completed",
            summary: {
                totalFiles: files.length,
                successful: successfulFiles.length,
                failed: invalidFiles.length
            },
            successfulFiles: successfulFiles.map(f => ({
                fileName: f.fileName,
                meterNumber: f.meterNumber,
                message: f.message
            })),
            invalidFiles: invalidFiles.map(f => ({
                fileName: f.fileName,
                reason: f.reason
            }))
        };

        if (successfulFiles.length === 0 && invalidFiles.length > 0) {
            return res.status(400).json(response);
        }

        if (invalidFiles.length > 0) {
            response.message = "Some files were not processed successfully";
            return res.status(207).json(response); // 207 Multi-Status
        }

        res.status(201).json(response);
    } catch (error) {
        logger.error(`Error uploading CSV files: ${error.message}`);
        res.status(500).json({
            message: "Internal server error",
            error: error.message
        });
    }
};


// Controller method to show meter data based on month, year, and main client ID
exports.showMeterData = async (req, res) => {
    try {
        const { month, year, mainClientId } = req.body;

        // Validate the input
        if (!month || !year || !mainClientId) {
            return res.status(400).json({ message: "Month, Year, and Main Client ID are required." });
        }

        // Find the main client by ID
        const mainClient = await MainClient.findById(mainClientId);
        if (!mainClient) {
            return res.status(404).json({ message: "Main Client not found." });
        }

        // Find all sub-clients related to the main client
        const subClients = await SubClient.find({ mainClient: mainClientId });
        if (!subClients.length) {
            return res.status(404).json({ message: "No sub-clients found for this main client." });
        }

        // Build the query for meter data based on month and year
        const startDate = moment(`${year}-${month}-01`, "YYYY-MM-DD");
        const endDate = startDate.clone().add(1, 'month');

        // Find meter data for the main client
        const mainClientMeterData = await MeterData.find({
            client: mainClientId,
            month,
            year,
        })
        .select('client meterNumber meterType month year') // Include necessary fields only
        .populate({
            path: 'client',
            select: 'name', // Only select the name field of the client
            model: 'MainClient',
        })
        .populate('meterType clientType')
        .lean();

        // Find meter data for all sub-clients for the given month and year
        const subClientMeterData = await MeterData.find({
            client: { $in: subClients.map(client => client._id) },
            month,
            year,
        })
        .select('client meterNumber meterType month year') // Include necessary fields only
        .populate({
            path: 'client',
            select: 'name',
            model: 'SubClient',
        })
        .populate('meterType clientType')
        .lean();

        // Combine the main client and sub-client meter data
        const allMeterData = [...mainClientMeterData, ...subClientMeterData];

        if (!allMeterData.length) {
            return res.status(404).json({ message: "No meter data found for the specified period." });
        }

        // Return the combined meter data (both main client and sub-client data)
        res.status(200).json({ message: "Meter data fetched successfully.", data: allMeterData });
    } catch (error) {
        logger.error(`Error fetching meter data: ${error.message}`);
        res.status(500).json({ message: "An error occurred while fetching the meter data." });
    }
};


// Controller method to delete meter data by ID
exports.deleteMeterData = async (req, res) => {
    try {
        const { meterDataId } = req.params;

        // Validate the input
        if (!meterDataId) {
            return res.status(400).json({ message: "Meter Data ID is required." });
        }

        // Find the meter data entry by ID
        const meterData = await MeterData.findById(meterDataId);
        if (!meterData) {
            return res.status(404).json({ message: "Meter data not found." });
        }

        const { meterNumber, month, year } = meterData;

        // Delete related LossesCalculationData entries
        await LossesCalculationData.deleteMany({
            $or: [
                { 'mainClient.meterNumber': meterNumber, month, year },
                { 'subClient.meterNumber': meterNumber, month, year }
            ]
        });
        logger.info(`Deleted LossesCalculationData related to MeterData: ${meterNumber} for ${month}/${year}`);

        // Delete related DailyReport entries
        await DailyReport.deleteMany({
            $or: [
                { 'mainClient.meterNumber': meterNumber, month, year },
                { 'subClient.meterNumber': meterNumber, month, year }
            ]
        });
        logger.info(`Deleted DailyReport related to MeterData: ${meterNumber} for ${month}/${year}`);

        // Delete related TotalReport entries
        await TotalReport.deleteMany({
            'clients.abtMainMeter.meterNumber': meterNumber,
            month,
            year
        });
        logger.info(`Deleted TotalReport related to MeterData: ${meterNumber} for ${month}/${year}`);

        // Delete associated files if they exist
        if (meterData.dataEntries?.length) {
            meterData.dataEntries.forEach(entry => {
                if (entry.filePath && fs.existsSync(entry.filePath)) {
                    fs.unlinkSync(entry.filePath);
                    logger.info(`Deleted file: ${entry.filePath}`);
                }
            });
        }

        // Delete the meter data entry from the database
        await MeterData.deleteOne({ _id: meterDataId });
        logger.info(`Meter data with ID ${meterDataId} deleted successfully.`);

        // Send response
        res.status(200).json({
            message: "Meter data and all related reports (LossesCalculationData, DailyReport, TotalReport) deleted successfully.",
            deleted: { meterDataId, meterNumber, month, year }
        });
    } catch (error) {
        logger.error(`Error deleting meter data: ${error.message}`);
        res.status(500).json({
            message: "An error occurred while deleting the meter data and related reports.",
            error: error.message
        });
    }
};
