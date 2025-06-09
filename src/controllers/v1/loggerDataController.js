const fs = require('fs');
const path = require('path');
const Papa = require('papaparse');
const logger = require('../../utils/logger');
const LoggerData = require('../../models/v1/loggerData.model');
const SubClient = require('../../models/v1/subClient.model');
const moment = require('moment');
const DailyReport = require('../../models/v1/dailyReport.model');

const extractMeterNumbers = (csvData) => {
    const meterNumbers = csvData[1].slice(1).map(meter => meter.trim());
    logger.info("Extracted Meter Numbers:", meterNumbers);
    return meterNumbers;
};

const findSubClientByMeterNumber = async (meterNumber) => {
    try {
        return await SubClient.findOne({
            $or: [
                { 'abtMainMeter.meterNumber': meterNumber },
                { 'abtCheckMeter.meterNumber': meterNumber }
            ]
        });
    } catch (error) {
        logger.error(`Error finding subclient by meter number ${meterNumber}: ${error.message}`);
        throw new Error("Error finding subclient");
    }
};

exports.uploadLoggerCSV = async (req, res) => {
    try {
        const { month, year } = req.body;
        const file = req.file;

        if (!file) {
            logger.warn("No file uploaded.");
            return res.status(400).json({ message: 'CSV file is required.' });
        }

        const filePath = path.resolve(file.path);
        const csvData = [];

        fs.createReadStream(filePath)
            .pipe(Papa.parse(Papa.NODE_STREAM_INPUT, {
                header: false,
                skipEmptyLines: true
            }))
            .on('data', (row) => {
                csvData.push(row);
            })
            .on('end', async () => {
                try {
                    if (csvData.length < 3) {
                        fs.unlinkSync(filePath);
                        return res.status(400).json({ message: 'CSV file must have at least 3 rows (header, meter numbers, and data).' });
                    }

                    let dateValidationFailed = false;
                    let firstInvalidDate = null;

                    const inputMonthNum = parseInt(month, 10);
                    const inputYearNum = parseInt(year, 10);

                    for (let i = 2; i < csvData.length; i++) {
                        const dateStr = csvData[i][0];

                        // Skip empty date rows instead of failing
                        if (!dateStr || dateStr.trim() === '') {
                            continue;
                        }

                        const dateParts = dateStr.split('-');
                        if (dateParts.length !== 3) {
                            dateValidationFailed = true;
                            firstInvalidDate = dateStr;
                            break;
                        }

                        const [day, csvMonth, csvYear] = dateParts.map(part => parseInt(part, 10));

                        if (isNaN(day) || isNaN(csvMonth) || isNaN(csvYear)) {
                            dateValidationFailed = true;
                            firstInvalidDate = dateStr;
                            break;
                        }

                        if (csvMonth !== inputMonthNum || csvYear !== inputYearNum) {
                            dateValidationFailed = true;
                            firstInvalidDate = `${day}-${csvMonth}-${csvYear}`;
                            break;
                        }
                    }

                    if (dateValidationFailed) {
                        fs.unlinkSync(filePath);
                        return res.status(400).json({
                            message: 'Selected month/year does not match dates in CSV file.',
                            details: {
                                selectedMonth: month,
                                selectedYear: year,
                                firstInvalidDateFound: firstInvalidDate,
                                expectedFormat: 'dd-mm-yyyy'
                            }
                        });
                    }

                    const meterNumbers = extractMeterNumbers(csvData);
                    if (!meterNumbers || meterNumbers.length === 0) {
                        fs.unlinkSync(filePath);
                        return res.status(400).json({ message: 'No meter numbers found in CSV file.' });
                    }

                    const existingData = await LoggerData.findOne({ month: inputMonthNum, year: inputYearNum });
                    if (existingData) {
                        const existingMeterNumbers = existingData.subClient.map(sc => sc.meterNumber);
                        const allMetersExist = meterNumbers.every(mn => existingMeterNumbers.includes(mn));

                        if (allMetersExist) {
                            fs.unlinkSync(filePath);
                            return res.status(409).json({ message: 'Logger data already exists for all meter numbers for this month and year.', existingData });
                        }

                        const someMetersExist = meterNumbers.some(mn => existingMeterNumbers.includes(mn));
                        if (someMetersExist) {
                            fs.unlinkSync(filePath);
                            return res.status(409).json({
                                message: 'Partial logger data already exists for some meter numbers for this month and year.',
                                existingMeters: existingMeterNumbers.filter(mn => meterNumbers.includes(mn)),
                                newMeters: meterNumbers.filter(mn => !existingMeterNumbers.includes(mn))
                            });
                        }
                    }

                    const loggerDataEntries = [];
                    for (let j = 0; j < meterNumbers.length; j++) {
                        const meterNumber = meterNumbers[j];
                        const subClient = await findSubClientByMeterNumber(meterNumber);

                        if (!subClient) {
                            logger.warn(`SubClient not found for meter number: ${meterNumber}`);
                            continue;
                        }

                        const existingMeterData = existingData?.subClient.find(sc => sc.meterNumber === meterNumber);
                        if (existingMeterData) {
                            logger.warn(`Data already exists for meter ${meterNumber} for ${month}/${year}`);
                            continue;
                        }

                        const subClientData = {
                            subClientId: subClient._id,
                            subClientName: subClient.name,
                            meterNumber,
                            meterType: subClient.abtMainMeter.meterNumber === meterNumber ? 'abtMainMeter' : 'abtCheckMeter',
                            loggerEntries: []
                        };

                        for (let i = 2; i < csvData.length; i++) {
                            const dateStr = csvData[i][0];
                            const data = parseFloat(csvData[i][j + 1]);

                            if (!dateStr || dateStr.trim() === '' || isNaN(data)) {
                                logger.warn(`Invalid date or data at row ${i + 1} for meter ${meterNumber}`);
                                continue;
                            }

                            subClientData.loggerEntries.push({ date: dateStr, data });
                        }

                        if (subClientData.loggerEntries.length > 0) {
                            loggerDataEntries.push(subClientData);
                        }
                    }

                    if (loggerDataEntries.length === 0) {
                        fs.unlinkSync(filePath);
                        return res.status(400).json({ message: 'No new valid logger data found in CSV file (all data may already exist).' });
                    }

                    if (existingData) {
                        existingData.subClient.push(...loggerDataEntries);
                        await existingData.save();
                        fs.unlinkSync(filePath);
                        return res.status(200).json({ message: 'New logger data added to existing month/year record.', updatedData: existingData });
                    }

                    const loggerData = new LoggerData({ month: inputMonthNum, year: inputYearNum, subClient: loggerDataEntries });
                    await loggerData.save();
                    fs.unlinkSync(filePath);

                    res.status(201).json({ message: 'Logger data uploaded and saved successfully.', data: loggerData });

                } catch (error) {
                    logger.error(`Error processing CSV data: ${error.message}`);
                    fs.unlinkSync(filePath);
                    res.status(500).json({ message: 'Error processing CSV data', error: error.message });
                }
            })
            .on('error', (err) => {
                logger.error(`Error parsing CSV file: ${err.message}`);
                fs.unlinkSync(filePath);
                res.status(500).json({ message: 'Error processing CSV file.' });
            });
    } catch (error) {
        logger.error(`Error in uploading logger CSV: ${error.message}`);
        if (req.file?.path) fs.unlinkSync(req.file.path);
        res.status(500).json({ message: 'Error in uploading logger CSV', error: error.message });
    }
};

// API to get logger data by month and year
exports.getLoggerDataByMonthYear = async (req, res) => {
    try {
        const { month, year } = req.params;

        // Validate input
        if (!month || !year) {
            logger.warn("Month or Year is missing from request parameters.");
            return res.status(400).json({ message: 'Month and Year are required.' });
        }

        // Fetch logger data for the given month and year
        const loggerData = await LoggerData.find({ month, year }).populate('subClient.subClientId', 'name');

        if (!loggerData.length) {
            logger.warn(`No logger data found for ${month}/${year}`);
            return res.status(404).json({ message: 'No logger data found for this month and year' });
        }

        logger.info(`Successfully retrieved logger data for ${month}/${year}`);
        res.status(200).json({ message: 'Logger data retrieved successfully', data: loggerData });
    } catch (error) {
        // Enhanced error logging
        logger.error(`Error retrieving logger data for ${month}/${year}: ${error.message}`, { error: error.stack });
        res.status(500).json({ message: 'An error occurred while retrieving logger data', error: error.message });
    }
};

// API to get specific logger data entry by ID
exports.getLoggerDataById = async (req, res) => {
    try {
        const { id } = req.params;

        if (!id) {
            logger.warn("Logger data ID is missing from request parameters.");
            return res.status(400).json({ message: 'Logger data ID is required.' });
        }

        // Fetch logger data by ID
        const loggerData = await LoggerData.findById(id).populate('subClient.subClientId', 'name');

        if (!loggerData) {
            logger.warn(`Logger data not found for ID: ${id}`);
            return res.status(404).json({ message: 'Logger data not found' });
        }

        logger.info(`Successfully retrieved logger data for ID: ${id}`);
        res.status(200).json({ message: 'Logger data retrieved successfully', data: loggerData });
    } catch (error) {
        // Enhanced error logging
        logger.error(`Error retrieving logger data by ID: ${req.params.id}: ${error.message}`, { error: error.stack });
        res.status(500).json({ message: 'An error occurred while retrieving logger data', error: error.message });
    }
};


exports.deleteLoggerData = async (req, res) => {
    try {
      const { loggerDataId } = req.params;
  
      // Validate input
      if (!loggerDataId) {
        logger.warn("Logger Data ID is missing from request parameters.");
        return res.status(400).json({ message: "Logger Data ID is required." });
      }
  
      // Find the logger data entry by ID
      const loggerData = await LoggerData.findById(loggerDataId);
      if (!loggerData) {
        logger.warn(`Logger data not found for ID: ${loggerDataId}`);
        return res.status(404).json({ message: "Logger data not found." });
      }
  
      // Step 1: Delete related DailyReport entries
      try {
        await DailyReport.deleteMany(
          { 'subClient.subClientId': { $in: loggerData.subClient.map(sub => sub.subClientId) } }
        );
        logger.info(`Deleted related DailyReports for LoggerData ID: ${loggerDataId}`);
      } catch (err) {
        logger.error(`Error deleting related DailyReports: ${err.message}`);
        return res.status(500).json({ message: "Error deleting related DailyReports." });
      }
  
      // Step 2: Check if the logger data has associated files to delete (assuming files are stored in a directory)
      try {
        if (loggerData.subClient && loggerData.subClient.length > 0) {
          for (const subClientData of loggerData.subClient) {
            for (const entry of subClientData.loggerEntries) {
              if (entry.filePath && fs.existsSync(entry.filePath)) {
                fs.unlinkSync(entry.filePath); // Delete the file from the server
                logger.info(`Deleted file: ${entry.filePath}`);
              }
            }
          }
        }
      } catch (err) {
        logger.error(`Error deleting files associated with LoggerData ID: ${loggerDataId}: ${err.message}`);
        return res.status(500).json({ message: "Error deleting associated files." });
      }
  
      // Step 3: Delete the logger data entry from the database
      try {
        await LoggerData.deleteOne({ _id: loggerDataId });
        logger.info(`Logger data with ID ${loggerDataId} deleted successfully.`);
      } catch (err) {
        logger.error(`Error deleting logger data entry from database: ${err.message}`);
        return res.status(500).json({ message: "Error deleting logger data from database." });
      }
  
      // Return success response
      res.status(200).json({ message: "Logger data and related DailyReports deleted successfully." });
    } catch (error) {
      logger.error(`Unexpected error during logger data deletion: ${error.message}`, { stack: error.stack });
      res.status(500).json({ message: "An unexpected error occurred while deleting the logger data.", error: error.message });
    }
  };



