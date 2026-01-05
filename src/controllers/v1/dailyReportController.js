const MainClient = require('../../models/v1/mainClient.model');
const MeterData = require('../../models/v1/meterData.model');
const SubClient = require('../../models/v1/subClient.model');
const DailyReport = require('../../models/v1/dailyReport.model');
const logger = require('../../utils/logger');
const LoggerData = require('../../models/v1/loggerData.model');
const { log } = require('winston');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');

// Helper function to calculate export and import for each entry
const calculateExportImport = (meterData, mf, pn) => {
    let exportSum = 0;
    let importSum = 0;
    // Logic based on the value of pn
    if (pn === -1) {
        meterData.forEach(entry => {
            exportSum += parseFloat(entry.parameters['Active(E) Total']);
            importSum += parseFloat(entry.parameters['Active(I) Total']);
        });
    } else if (pn === 1) {
        meterData.forEach(entry => {
            importSum += parseFloat(entry.parameters['Active(E) Total']);
            exportSum += parseFloat(entry.parameters['Active(I) Total']);
        });
    }

    exportSum *= mf; // Multiply by the main client's mf
    importSum *= mf;

    return { export: exportSum, import: importSum };
};
// const calculateExportImport = (meterData, mf) => {
//     let exportSum = 0;
//     let importSum = 0;

//     meterData.forEach(entry => {
//         exportSum += parseFloat(entry.parameters['Active(E) Total']);
//         importSum += parseFloat(entry.parameters['Active(I) Total']);
//     });

//     exportSum *= mf; // Multiply by the main client's mf
//     importSum *= mf;

//     return { export: exportSum, import: importSum };


// };

exports.generateDailyReport = async (req, res) => {
    try {
        const { mainClientId, month, year, showLoggerColumns = false, showAvgColumn = false, showAvgAcColumn = false } = req.body;

        // Track clients using check meters
        const clientsUsingCheckMeter = [];
        const clientsWithoutMeters = [];

        // Step 0: Check if report already exists
        // logger.info(`Checking for existing daily report for ${mainClientId}, ${month}-${year}`);
        // const existingReport = await DailyReport.findOne({ mainClientId, month, year });

        // if (existingReport) {
        //     logger.info("Existing report found. Updating timestamp.");
        //     existingReport.updatedAt = new Date();
        //     await existingReport.save();
        //     return res.status(200).json({
        //         message: 'Existing report retrieved successfully.',
        //         data: existingReport
        //     });
        // }

        // Step 1: Get Main Client Data
        const mainClientData = await MainClient.findById(mainClientId);
        if (!mainClientData) {
            logger.error(`Main Client not found: ${mainClientId}`);
            return res.status(404).json({ message: 'Main Client not found' });
        }

        // Step 2: Get Main Client Meter Data (with fallback)
        let meterData = await MeterData.find({
            meterNumber: mainClientData.abtMainMeter?.meterNumber,
            month,
            year
        });

        if (!meterData.length && mainClientData.abtCheckMeter?.meterNumber) {
            logger.info("Main Client abtMainMeter data not found, trying abtCheckMeter...");
            meterData = await MeterData.find({
                meterNumber: mainClientData.abtCheckMeter.meterNumber,
                month,
                year
            });
            if (meterData.length) {
                clientsUsingCheckMeter.push({
                    name: mainClientData.name,
                    type: 'Main Client'
                });
            }
        }

        if (!meterData.length) {
            logger.error(`No meter data found for Main Client: ${mainClientId}. Both abtMainMeter and abtCheckMeter files missing.`);
            clientsWithoutMeters.push({
                name: mainClientData.name,
                type: 'Main Client'
            });
            return res.status(400).json({
                message: `Meter data missing for Main Client (ID: ${mainClientId}). Both abtMainMeter and abtCheckMeter files are missing. Cannot generate report.`,
                clientsWithoutMeters,
                clientsUsingCheckMeter
            });
        }

        // Step 3: Get Sub Clients
        const subClients = await SubClient.find({ mainClient: mainClientId });
        if (!subClients.length) {
            logger.warn(`No subclients found for ${mainClientId}`);
        }

        // Step 4: Process Main Client Data
        let loggerData = [];
        let mainClientTotalExport = 0, mainClientTotalImport = 0, mainClientTotalAvgGeneration = 0, mainClientTotalAvgGenerationAc = 0;
        let currentDate = '', dailyExport = 0, dailyImport = 0;

        // Get main client DC capacity for avg generation calculation
        let mainClientDcCapacityKwp = mainClientData.dcCapacityKwp || mainClientData.acCapacityKw || 0;
        if (typeof mainClientDcCapacityKwp === 'string') {
            mainClientDcCapacityKwp = parseFloat(mainClientDcCapacityKwp.replace(/,/g, '').trim()) || 0;
        }
        if (!Number.isFinite(mainClientDcCapacityKwp) || mainClientDcCapacityKwp <= 0) {
            mainClientDcCapacityKwp = 0;
        }

        // Get main client AC capacity for avg generation AC calculation
        let mainClientAcCapacityKw = mainClientData.acCapacityKw || 0;
        if (typeof mainClientAcCapacityKw === 'string') {
            mainClientAcCapacityKw = parseFloat(mainClientAcCapacityKw.replace(/,/g, '').trim()) || 0;
        }
        if (!Number.isFinite(mainClientAcCapacityKw) || mainClientAcCapacityKw <= 0) {
            mainClientAcCapacityKw = 0;
        }

        for (let data of meterData) {
            for (let entry of data.dataEntries) {
                const entryDate = entry.parameters.Date;
                if (entryDate !== currentDate) {
                    if (currentDate !== '') {
                        // Calculate avg generation: export / dcCapacityKwp
                        const avgGeneration = mainClientDcCapacityKwp > 0 ? (dailyExport / mainClientDcCapacityKwp) : 0;
                        // Calculate avg generation AC: export / acCapacityKw
                        const avgGenerationAc = mainClientAcCapacityKw > 0 ? (dailyExport / mainClientAcCapacityKw) : 0;
                        loggerData.push({
                            date: currentDate,
                            export: dailyExport,
                            import: dailyImport,
                            avgGeneration: parseFloat(avgGeneration.toFixed(2)),
                            avgGenerationAc: parseFloat(avgGenerationAc.toFixed(2))
                        });
                        mainClientTotalExport += dailyExport;
                        mainClientTotalImport += dailyImport;
                        mainClientTotalAvgGeneration += avgGeneration;
                        mainClientTotalAvgGenerationAc += avgGenerationAc;
                    }
                    currentDate = entryDate;
                    dailyExport = 0;
                    dailyImport = 0;
                }

                const { export: dayExport, import: dayImport } = calculateExportImport([entry], mainClientData.mf, mainClientData.pn);
                dailyExport += dayExport;
                dailyImport += dayImport;
            }
        }

        if (currentDate !== '') {
            // Calculate avg generation for last day: export / dcCapacityKwp
            const avgGeneration = mainClientDcCapacityKwp > 0 ? (dailyExport / mainClientDcCapacityKwp) : 0;
            // Calculate avg generation AC for last day: export / acCapacityKw
            const avgGenerationAc = mainClientAcCapacityKw > 0 ? (dailyExport / mainClientAcCapacityKw) : 0;
            loggerData.push({
                date: currentDate,
                export: dailyExport,
                import: dailyImport,
                avgGeneration: parseFloat(avgGeneration.toFixed(2)),
                avgGenerationAc: parseFloat(avgGenerationAc.toFixed(2))
            });
            mainClientTotalExport += dailyExport;
            mainClientTotalImport += dailyImport;
            mainClientTotalAvgGeneration += avgGeneration;
            mainClientTotalAvgGenerationAc += avgGenerationAc;
        }

        // Calculate average of daily avg generation values (matching Excel =AVERAGE formula)
        const numberOfDays = loggerData.length;
        const avgOfAvgGeneration = numberOfDays > 0 ? (mainClientTotalAvgGeneration / numberOfDays) : 0;
        const avgOfAvgGenerationAc = numberOfDays > 0 ? (mainClientTotalAvgGenerationAc / numberOfDays) : 0;

        // Determine which capacity type was used for DC-based avg generation
        let capacityBasisDC = '';
        if (mainClientData.dcCapacityKwp && mainClientDcCapacityKwp > 0) {
            capacityBasisDC = 'DC';
        } else if (mainClientData.acCapacityKw && mainClientDcCapacityKwp > 0) {
            capacityBasisDC = 'AC';
        }

        // AC-based avg generation always uses AC capacity
        let capacityBasisAC = mainClientAcCapacityKw > 0 ? 'AC' : '';

        // Initialize daily report
        const dailyReport = new DailyReport({
            mainClientId,
            month,
            year,
            showLoggerColumns: showLoggerColumns === true, // Default to false if not provided
            showAvgColumn: showAvgColumn === true, // Default to false if not provided
            showAvgAcColumn: showAvgAcColumn === true, // Default to false if not provided
            capacityBasisDC: capacityBasisDC, // Track which capacity was used for DC avg gen
            capacityBasisAC: capacityBasisAC, // Track which capacity was used for AC avg gen
            mainClient: {
                meterNumber: meterData[0].meterNumber,
                meterType: meterData[0].meterType,
                mainClientDetail: mainClientData.toObject(),
                totalexport: mainClientTotalExport,
                totalimport: mainClientTotalImport,
                totalAvgGeneration: parseFloat(avgOfAvgGeneration.toFixed(2)),
                totalAvgGenerationAc: parseFloat(avgOfAvgGenerationAc.toFixed(2)),
                loggerdatas: loggerData
            },
            subClient: [],
            aclinelossdiffrence: {}
        });

        // Step 5: Process Sub Clients with logger data validation
        for (let subClient of subClients) {
            // Debug: Log subClient fields to verify dcCapacityKwp is available
            logger.info(`Processing subclient: ${subClient.name}, dcCapacityKwp: ${subClient.dcCapacityKwp}, type: ${typeof subClient.dcCapacityKwp}`);
            // Get subclient meter data
            // Try abtMainMeter meter data first
            let subClientMeterData = await MeterData.find({
                meterNumber: subClient.abtMainMeter?.meterNumber,
                month,
                year
            });

            // If none found, try abtCheckMeter
            if (!subClientMeterData.length && subClient.abtCheckMeter?.meterNumber) {
                logger.info(`Sub Client ${subClient.name} abtMainMeter data not found, trying abtCheckMeter...`);
                subClientMeterData = await MeterData.find({
                    meterNumber: subClient.abtCheckMeter.meterNumber,
                    month,
                    year
                });
                if (subClientMeterData.length) {
                    clientsUsingCheckMeter.push({
                        name: subClient.name,
                        type: 'Sub Client'
                    });
                }
            }

            // If still none, add to missing list and skip
            if (!subClientMeterData.length) {
                logger.error(`No meter data found for subclient ${subClient.name} (ID: ${subClient._id}). Both abtMainMeter and abtCheckMeter files missing.`);
                clientsWithoutMeters.push({
                    name: subClient.name,
                    type: 'Sub Client'
                });
                continue;
            }

            // Get logger data for this subclient
            const loggerDataDocs = await LoggerData.find({
                month: month.toString(),
                year: year.toString(),
                'subClient.subClientId': subClient._id
            });

            // if (!loggerDataDocs.length) {
            //     logger.error(`No logger data for subclient ${subClient._id}`);
            //     return res.status(404).json({
            //         message: `Logger data not available for subclient ${subClient.name}`,
            //         subClientId: subClient._id,
            //         subClientName: subClient.name,
            //         clientsWithoutMeters,
            //         clientsUsingCheckMeter
            //     });
            // }

            if (!loggerDataDocs.length) {
                logger.warn(`No logger data found for subclient ${subClient.name}, using 00 as default values`);
            }

            let subClientLoggerData = [];
            let subClientTotalExport = 0, subClientTotalImport = 0;
            let subClientTotalLoggerData = 0, subClientTotalInternalLosses = 0;
            let subClientTotalAvgGeneration = 0, subClientTotalAvgGenerationAc = 0;
            let subClientCurrentDate = '', dailyExport = 0, dailyImport = 0;

            // Get DC Capacity for avg generation calculation
            // Convert to number and handle string values, null, undefined
            let dcCapacityKwp = 0;

            // Try to get dcCapacityKwp from subClient object
            let rawDcCapacity = subClient.dcCapacityKwp;

            // If not found, try to fetch it directly from database
            if (!rawDcCapacity || rawDcCapacity === null || rawDcCapacity === undefined || rawDcCapacity === '') {
                try {
                    const freshSubClient = await SubClient.findById(subClient._id).select('dcCapacityKwp');
                    if (freshSubClient && freshSubClient.dcCapacityKwp != null) {
                        rawDcCapacity = freshSubClient.dcCapacityKwp;
                        logger.info(`Retrieved dcCapacityKwp from database for subclient ${subClient.name}: ${rawDcCapacity}`);
                    }
                } catch (err) {
                    logger.error(`Error fetching dcCapacityKwp from database for subclient ${subClient.name}: ${err.message}`);
                }
            }

            // Check if field exists and has a value
            if (rawDcCapacity != null && rawDcCapacity !== undefined && rawDcCapacity !== '') {
                // Try to convert to number, handling strings with commas or other formatting
                const cleanedValue = String(rawDcCapacity).replace(/,/g, '').trim();
                dcCapacityKwp = Number(cleanedValue);

                if (isNaN(dcCapacityKwp) || !Number.isFinite(dcCapacityKwp) || dcCapacityKwp <= 0) {
                    dcCapacityKwp = 0;
                    logger.warn(`DC Capacity (dcCapacityKwp) is invalid for subclient ${subClient.name} (ID: ${subClient._id}). Raw value: ${rawDcCapacity}, Cleaned: ${cleanedValue}, Converted: ${dcCapacityKwp}. Avg Generation will be 0.`);
                } else {
                    logger.info(`Using DC Capacity ${dcCapacityKwp} kWp for subclient ${subClient.name} (ID: ${subClient._id})`);
                }
            } else {
                logger.warn(`DC Capacity (dcCapacityKwp) is missing/null/undefined/empty for subclient ${subClient.name} (ID: ${subClient._id}). Original value: ${subClient.dcCapacityKwp}, Fetched value: ${rawDcCapacity}. Avg Generation will be 0.`);
            }

            // Get AC Capacity for avg generation AC calculation
            let acCapacityKw = 0;
            let rawAcCapacity = subClient.acCapacityKw;

            if (rawAcCapacity != null && rawAcCapacity !== undefined && rawAcCapacity !== '') {
                const cleanedAcValue = String(rawAcCapacity).replace(/,/g, '').trim();
                acCapacityKw = Number(cleanedAcValue);

                if (isNaN(acCapacityKw) || !Number.isFinite(acCapacityKw) || acCapacityKw <= 0) {
                    acCapacityKw = 0;
                    logger.warn(`AC Capacity (acCapacityKw) is invalid for subclient ${subClient.name} (ID: ${subClient._id}). Avg Generation AC will be 0.`);
                } else {
                    logger.info(`Using AC Capacity ${acCapacityKw} kW for subclient ${subClient.name} (ID: ${subClient._id})`);
                }
            } else {
                logger.warn(`AC Capacity (acCapacityKw) is missing for subclient ${subClient.name} (ID: ${subClient._id}). Avg Generation AC will be 0.`);
            }

            for (let data of subClientMeterData) {
                for (let entry of data.dataEntries) {
                    const entryDate = entry.parameters.Date;
                    if (entryDate !== subClientCurrentDate) {
                        if (subClientCurrentDate !== '') {
                            // Find matching logger data or use '00' if not available
                            let loggerDataValue = 0;
                            let foundMatch = false;

                            if (loggerDataDocs.length > 0) {
                                for (const doc of loggerDataDocs) {
                                    const matchingSubClient = doc.subClient.find(
                                        sc => sc.subClientId.toString() === subClient._id.toString()
                                    );

                                    if (matchingSubClient) {
                                        const matchingEntry = matchingSubClient.loggerEntries.find(
                                            e => e.date === subClientCurrentDate
                                        );
                                        if (matchingEntry) {
                                            loggerDataValue = matchingEntry.data;
                                            foundMatch = true;
                                            break;
                                        }
                                    }
                                }
                            }

                            // Calculate losses - MODIFIED PART
                            let internallosse = 0; // Default to 0
                            let lossinparsantege = 0; // Default to 0


                            // If no logger data found, use 00 as default
                            if (!foundMatch) {
                                logger.warn(`No logger data for ${subClientCurrentDate}, using 00 as default`);
                                loggerDataValue = 0;
                            }

                            // Only calculate loss if we found matching logger data
                            if (foundMatch) {
                                internallosse = dailyExport - loggerDataValue;
                                lossinparsantege = dailyExport !== 0 ? (internallosse / dailyExport) * 100 : 0;
                            }

                            // Calculate avg generation: export / dcCapacityKwp
                            const avgGeneration = dcCapacityKwp > 0 ? (dailyExport / dcCapacityKwp) : 0;
                            // Calculate avg generation AC: export / acCapacityKw
                            const avgGenerationAc = acCapacityKw > 0 ? (dailyExport / acCapacityKw) : 0;

                            subClientLoggerData.push({
                                date: subClientCurrentDate,
                                export: parseFloat(dailyExport.toFixed(2)),
                                import: parseFloat(dailyImport.toFixed(2)),
                                loggerdata: loggerDataValue,
                                internallosse: parseFloat(internallosse.toFixed(2)),
                                lossinparsantege: parseFloat(lossinparsantege.toFixed(2)),
                                avgGeneration: parseFloat(avgGeneration.toFixed(2)),
                                avgGenerationAc: parseFloat(avgGenerationAc.toFixed(2))
                            });

                            subClientTotalExport += dailyExport;
                            subClientTotalImport += dailyImport;
                            subClientTotalLoggerData += loggerDataValue;
                            subClientTotalInternalLosses += internallosse;
                            subClientTotalAvgGeneration += avgGeneration;
                            subClientTotalAvgGenerationAc += avgGenerationAc;
                        }

                        subClientCurrentDate = entryDate;
                        dailyExport = 0;
                        dailyImport = 0;
                    }

                    const { export: dayExport, import: dayImport } = calculateExportImport([entry], subClient.mf, subClient.pn);
                    dailyExport += dayExport;
                    dailyImport += dayImport;
                }
            }

            // Process last day
            if (subClientCurrentDate !== '') {
                let loggerDataValue = 0;
                let foundMatch = false;

                if (loggerDataDocs.length > 0) {
                    for (const doc of loggerDataDocs) {
                        const matchingSubClient = doc.subClient.find(
                            sc => sc.subClientId.toString() === subClient._id.toString()
                        );

                        if (matchingSubClient) {
                            const matchingEntry = matchingSubClient.loggerEntries.find(
                                e => e.date === subClientCurrentDate
                            );
                            if (matchingEntry) {
                                loggerDataValue = matchingEntry.data;
                                foundMatch = true;
                                break;
                            }
                        }
                    }
                }

                // Calculate losses - MODIFIED PART
                let internallosse = 0; // Default to 0
                let lossinparsantege = 0; // Default to 0

                // Only calculate loss if we found matching logger data
                if (foundMatch) {
                    internallosse = dailyExport - loggerDataValue;
                    lossinparsantege = dailyExport !== 0 ? (internallosse / dailyExport) * 100 : 0;
                }

                // Calculate avg generation: export / dcCapacityKwp
                const avgGeneration = dcCapacityKwp > 0 ? (dailyExport / dcCapacityKwp) : 0;
                // Calculate avg generation AC: export / acCapacityKw
                const avgGenerationAc = acCapacityKw > 0 ? (dailyExport / acCapacityKw) : 0;

                subClientLoggerData.push({
                    date: subClientCurrentDate,
                    export: parseFloat(dailyExport.toFixed(2)),
                    import: parseFloat(dailyImport.toFixed(2)),
                    loggerdata: loggerDataValue,
                    internallosse: parseFloat(internallosse.toFixed(2)),
                    lossinparsantege: parseFloat(lossinparsantege.toFixed(2)),
                    avgGeneration: parseFloat(avgGeneration.toFixed(2)),
                    avgGenerationAc: parseFloat(avgGenerationAc.toFixed(2))
                });

                subClientTotalExport += dailyExport;
                subClientTotalImport += dailyImport;
                subClientTotalLoggerData += loggerDataValue;
                subClientTotalInternalLosses += internallosse;
                subClientTotalAvgGeneration += avgGeneration;
                subClientTotalAvgGenerationAc += avgGenerationAc;
            }

            // Calculate average of daily avg generation values (matching Excel =AVERAGE formula)
            const numberOfDays = subClientLoggerData.length;
            const totalAvgGeneration = numberOfDays > 0 ? parseFloat((subClientTotalAvgGeneration / numberOfDays).toFixed(2)) : 0;
            const totalAvgGenerationAc = numberOfDays > 0 ? parseFloat((subClientTotalAvgGenerationAc / numberOfDays).toFixed(2)) : 0;

            // Add subclient to report
            dailyReport.subClient.push({
                name: subClient.name,
                divisionName: subClient.divisionName,
                consumerNo: subClient.consumerNo,
                contactNo: subClient.contactNo,
                email: subClient.email,
                subClientId: subClient._id,
                meterNumber: subClientMeterData[0].meterNumber,
                meterType: subClientMeterData[0].meterType,
                totalexport: parseFloat(subClientTotalExport.toFixed(2)),
                totalimport: parseFloat(subClientTotalImport.toFixed(2)),
                totalloggerdata: parseFloat(subClientTotalLoggerData.toFixed(2)),
                totalinternallosse: parseFloat(subClientTotalInternalLosses.toFixed(2)),
                totallossinparsantege: subClientTotalExport !== 0 ?
                    parseFloat(((subClientTotalInternalLosses / subClientTotalExport) * 100).toFixed(2)) : 0,
                totalAvgGeneration: totalAvgGeneration,
                totalAvgGenerationAc: totalAvgGenerationAc,
                loggerdatas: subClientLoggerData
            });
        }

        // Calculate ACLineLossDifference
        const aclineLossData = {
            loggerdatas: [],
            totalexport: 0,
            totalimport: 0,
            totallossinparsantegeexport: 0,
            totallossinparsantegeimport: 0
        };

        dailyReport.mainClient.loggerdatas.forEach(mainEntry => {
            const subEntries = dailyReport.subClient.flatMap(sub =>
                sub.loggerdatas.filter(e => e.date === mainEntry.date)
            );

            const subExportSum = subEntries.reduce((sum, e) => sum + e.export, 0);
            const subImportSum = subEntries.reduce((sum, e) => sum + e.import, 0);

            const exportDiff = mainEntry.export - subExportSum;
            const importDiff = mainEntry.import - subImportSum;

            const exportPercentage = mainEntry.export !== 0 ?
                (exportDiff / mainEntry.export) * 100 : 0;
            const importPercentage = mainEntry.import !== 0 ?
                (importDiff / mainEntry.import) * 100 : 0;

            aclineLossData.loggerdatas.push({
                date: mainEntry.date,
                export: parseFloat(exportDiff.toFixed(2)),
                import: parseFloat(importDiff.toFixed(2)),
                lossinparsantegeexport: parseFloat(exportPercentage.toFixed(2)),
                lossinparsantegeimport: parseFloat(importPercentage.toFixed(2))
            });

            aclineLossData.totalexport += exportDiff;
            aclineLossData.totalimport += importDiff;
        });

        aclineLossData.totallossinparsantegeexport = mainClientTotalExport !== 0 ?
            parseFloat(((aclineLossData.totalexport / mainClientTotalExport) * 100).toFixed(2)) : 0;
        aclineLossData.totallossinparsantegeimport = mainClientTotalImport !== 0 ?
            parseFloat(((aclineLossData.totalimport / mainClientTotalImport) * 100).toFixed(2)) : 0;

        dailyReport.aclinelossdiffrence = aclineLossData;

        // Save and return
        await dailyReport.save();

        // Include meter status information in response
        const responseData = {
            message: 'Daily report generated successfully',
            data: dailyReport
        };

        if (clientsUsingCheckMeter.length > 0) {
            responseData.clientsUsingCheckMeter = clientsUsingCheckMeter;
        }

        if (clientsWithoutMeters.length > 0) {
            responseData.clientsWithoutMeters = clientsWithoutMeters;
        }

        res.status(201).json(responseData);

    } catch (error) {
        logger.error(`Error generating report: ${error.message}`);
        res.status(500).json({
            message: 'Error generating daily report',
            error: error.message
        });
    }
};

// Get latest 10 daily reports
exports.getLatestDailyReports = async (req, res) => {
    try {
        // Get latest 10 reports sorted by update date (newest first)
        const latestReports = await DailyReport.find({})
            .sort({ updatedAt: -1 }) // Changed to sort by updatedAt
            .limit(10)
            .lean();

        if (!latestReports || latestReports.length === 0) {
            return res.status(404).json({
                message: 'No daily reports found'
            });
        }

        // Transform the data
        const simplifiedReports = latestReports.map(report => {
            return {
                id: report._id,
                month: report.month,
                year: report.year,
                clientName: report.mainClient?.mainClientDetail?.name || 'N/A',
                lastUpdated: report.updatedAt, // Added last updated timestamp
                generatedAt: report.createdAt, // Keep original creation timestamp
                totalExport: report.mainClient?.totalexport || 0,
                totalImport: report.mainClient?.totalimport || 0,
                reportType: 'Daily'
            };
        });

        res.status(200).json({
            message: 'Latest 10 daily reports retrieved successfully (sorted by last update)',
            data: simplifiedReports
        });

    } catch (error) {
        logger.error(`Error fetching daily reports: ${error.message}`);
        res.status(500).json({
            message: 'Error fetching daily reports',
            error: error.message
        });
    }
};

exports.downloadDailyReportExcel = async (req, res) => {
    try {
        const { dailyReportId } = req.params;
        const dailyReport = await DailyReport.findById(dailyReportId)
            .populate('mainClientId')
            .populate('subClient.subClientId');

        if (!dailyReport) {
            return res.status(404).json({ message: 'Daily Report not found' });
        }

        const workbook = new ExcelJS.Workbook();
        // Change sheet name to "Master Sheet"
        const worksheet = workbook.addWorksheet('Master Sheet');
        worksheet.pageSetup.orientation = 'landscape';
        const worksheetSetup = {
            margins: {
                left: 0.2,
                right: 0.2,
                top: 0.2,
                bottom: 0.2,
                header: 0.2,
                footer: 0.2
            },
            horizontalCentered: true,
            verticalCentered: false,
            fitToPage: true,
            fitToWidth: 1,
            fitToHeight: 1,
            paperSize: 9, // A4
            orientation: 'landscape'
        };
        worksheet.pageSetup = worksheetSetup;

        // Get month name and format for filename
        const filemonthName = getMonthName(dailyReport.month).toUpperCase();
        const year = dailyReport.year;
        const mainClientName = dailyReport.mainClient.mainClientDetail.name
            .replace(/\s+/g, '_')       // Replace spaces with underscores
            .replace(/\//g, '_')        // Replace forward slashes with underscores
            .toUpperCase();

        // Calculate dynamic column counts
        const subClientCount = dailyReport.subClient.length;
        const showLoggerColumns = dailyReport.showLoggerColumns === true; // Default to false
        const showAvgColumn = dailyReport.showAvgColumn === true; // Default to false (DC Capacity)
        const showAvgAcColumn = dailyReport.showAvgAcColumn === true; // Default to false (AC Capacity)
        // Column structure: Export + (Avg DC?) + (Avg AC?) + (Logger Data + Internal Loss + Loss in%?) + Import
        const colsPerSubClient = 1 + (showAvgColumn ? 1 : 0) + (showAvgAcColumn ? 1 : 0) + (showLoggerColumns ? 3 : 0) + 1; // Export + Avg DC + Avg AC + Logger cols + Import
        const colsPerMainClient = 1 + (showAvgColumn ? 1 : 0) + (showAvgAcColumn ? 1 : 0) + 1; // Export + Avg DC + Avg AC + Import
        const subClientCols = subClientCount * colsPerSubClient;
        const totalCols = 1 + colsPerMainClient + subClientCols + 4; // Date(1) + Main + Subs + AC Line(4)
        // Determine how many columns to cover in row 3 (minimum K, maximum same as row 2)
        const row3CoverCols = Math.max(11, totalCols); // At least up to column K (11), but more if needed

        // ====================
        // SET COLUMN WIDTHS
        // ====================

        // Main client columns
        worksheet.getColumn('A').width = 12;  // A cell width is 12
        worksheet.getColumn('B').width = 10;   // B cell width (Export)
        let mainColIndex = 3; // Start from column C
        if (showAvgColumn) {
            worksheet.getColumn(mainColIndex).width = 10;   // Avg Generation DC
            mainColIndex++;
        }
        if (showAvgAcColumn) {
            worksheet.getColumn(mainColIndex).width = 10;   // Avg Generation AC
            mainColIndex++;
        }
        worksheet.getColumn(mainColIndex).width = 10;   // Import

        // Sub-client columns (repeated for each sub-client)
        dailyReport.subClient.forEach((subClient, index) => {
            // Main client starts at column 2 (B), so subclients start after main client ends
            // Main client end column = 1 (Date) + colsPerMainClient, so subclients start at 2 + colsPerMainClient
            const baseCol = 2 + colsPerMainClient + (index * colsPerSubClient); // Starting column for each sub-client
            let colOffset = 0;

            worksheet.getColumn(baseCol + colOffset).width = 9;     // Export
            colOffset++;

            if (showAvgColumn) {
                worksheet.getColumn(baseCol + colOffset).width = 10; // Avg Generation DC
                colOffset++;
            }

            if (showAvgAcColumn) {
                worksheet.getColumn(baseCol + colOffset).width = 10; // Avg Generation AC
                colOffset++;
            }

            if (showLoggerColumns) {
                worksheet.getColumn(baseCol + colOffset).width = 9; // Logger Data
                colOffset++;
                worksheet.getColumn(baseCol + colOffset).width = 9; // Internal Loss
                colOffset++;
                worksheet.getColumn(baseCol + colOffset).width = 9; // Loss in%
                colOffset++;
            }

            worksheet.getColumn(baseCol + colOffset).width = 8; // Import
        });

        // AC LINE LOSS DIFF columns
        const acLineStartCol = 2 + colsPerMainClient + (dailyReport.subClient.length * colsPerSubClient);
        worksheet.getColumn(acLineStartCol).width = 9;     // 1st cell
        worksheet.getColumn(acLineStartCol + 1).width = 8; // 2nd cell
        worksheet.getColumn(acLineStartCol + 2).width = 8; // 3rd cell
        worksheet.getColumn(acLineStartCol + 3).width = 8; // 4th cell

        // ====================
        // ROW 2: TITLE ROW
        // ====================
        // Calculate dynamic column counts
        const minSubClients = 2; // Minimum 2 subclients
        const effectiveSubClients = Math.max(subClientCount, minSubClients);
        const subClientColss = effectiveSubClients * colsPerSubClient;
        const totalColss = 1 + colsPerMainClient + subClientColss + 4; // Date(1) + Main + Subs + AC Line(4)

        worksheet.mergeCells(2, 1, 2, totalColss);
        const titleCell = worksheet.getCell(2, 1);
        //height
        worksheet.getRow(2).height = 30;
        titleCell.value = `${dailyReport.mainClient.mainClientDetail.name} - ${(dailyReport.mainClient.mainClientDetail.acCapacityKw / 1000).toFixed(2)} MW AC Generation Details`;
        titleCell.font = {
            name: 'Times New Roman',
            size: 16,
            bold: true,
            color: { argb: 'FF000000' } // Black
        };
        titleCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFbdd7ee' } // bdd7ee background
        };
        titleCell.alignment = {
            horizontal: 'center',
            vertical: 'middle'
        };
        titleCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // ====================
        // ROW 3: DATE ROW
        // ====================
        const monthName = getMonthName(dailyReport.month);
        const yearShort = dailyReport.year.toString().slice(-2);
        const lastDay = new Date(dailyReport.year, dailyReport.month, 0).getDate();

        // Calculate merge ranges
        const monthEndCol = 1 + colsPerMainClient + colsPerSubClient; // NAME (1) + Total GETCO + 1 subclient (colsPerSubClient)
        const dateRangeEndCol = totalColss; // Goes all the way to AC LINE end

        // Month cell (A to monthEndCol)
        worksheet.mergeCells(3, 1, 3, monthEndCol);
        const monthCell = worksheet.getCell(3, 1);
        // Set row height
        worksheet.getRow(3).height = 30;
        monthCell.value = `Month: ${monthName}-${yearShort}`;
        monthCell.font = {
            name: 'Times New Roman',
            size: 14,
            bold: true,
            color: { argb: 'FFFF0000' } // Red
        };
        monthCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFbfbfbf' } // bfbfbf background
        };
        monthCell.alignment = {
            horizontal: 'center',
            vertical: 'middle'
        };
        monthCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // Date range cell (next column after month to end)
        const dateRange = `Generation Period: 01-${String(dailyReport.month).padStart(2, '0')}-${dailyReport.year} to ${lastDay}-${String(dailyReport.month).padStart(2, '0')}-${dailyReport.year}`;
        worksheet.mergeCells(3, monthEndCol + 1, 3, dateRangeEndCol);
        const periodCell = worksheet.getCell(3, monthEndCol + 1);
        periodCell.value = dateRange;
        periodCell.font = {
            name: 'Times New Roman',
            size: 14,
            bold: true,
            color: { argb: 'FFFF0000' } // Red
        };
        periodCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFbfbfbf' } // bfbfbf background
        };
        periodCell.alignment = {
            horizontal: 'center',
            vertical: 'middle'
        };
        periodCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        // ====================
        // ROW 4: CLIENT NAMES
        // ====================

        const clientNamesRow = worksheet.getRow(4);
        clientNamesRow.height = 40; // Set row height (you can increase if needed)

        // First cell (A4)
        const nameLabelCell = worksheet.getCell('A4');
        nameLabelCell.value = 'NAME=>';
        nameLabelCell.font = {
            name: 'Times New Roman',
            size: 10,
            bold: true
        };
        nameLabelCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFbfbfbf' } // Gray background
        };
        nameLabelCell.alignment = {
            horizontal: 'center',
            vertical: 'middle',
            wrapText: true // Enable text wrapping
        };
        nameLabelCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // Main client cells - dynamic end column based on avg columns
        const mainClientEndColNum = 1 + colsPerMainClient; // B is 2, so end is 2 + colsPerMainClient - 1
        const mainClientEndCol = String.fromCharCode(64 + mainClientEndColNum);
        worksheet.mergeCells(`B4:${mainClientEndCol}4`);
        const mainClientCell = worksheet.getCell('B4');
        mainClientCell.value = 'TOTAL-GETCO\nSS'; // Explicit line break
        mainClientCell.font = {
            name: 'Times New Roman',
            size: 10,
            bold: true
        };
        mainClientCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFffc000' } // Orange background
        };
        mainClientCell.alignment = {
            horizontal: 'center',
            vertical: 'middle',
            wrapText: true // Ensure wrapText is enabled
        };
        // Add thick outside border to main client cells
        mainClientCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        // Note: No need to set individual cell borders for merged cells

        // Sub-client cells - alternating colors
        const subClientColors = ['FFFF00', '92D050']; // Yellow and green
        dailyReport.subClient.forEach((subClient, index) => {
            const startCol = 2 + colsPerMainClient + (index * colsPerSubClient);
            const color = subClientColors[index % 2]; // Alternate between colors

            worksheet.mergeCells(4, startCol, 4, startCol + colsPerSubClient - 1);
            const cell = worksheet.getCell(4, startCol);
            cell.value = `${subClient.name}`;
            cell.font = {
                name: 'Times New Roman',
                size: 10,
                bold: true
            };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: `FF${color}` }
            };
            cell.alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true // Add this to all cells that need wrapping
            };
            // Add thick outside border to sub-client cells
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            // Add thin inside borders
            for (let i = 1; i < colsPerSubClient; i++) {
                worksheet.getCell(4, startCol + i).border.left = { style: 'thin' };
            }
        });

        // AC LINE LOSS DIFF header (merge rows 4 and 5)
        worksheet.mergeCells(4, acLineStartCol, 5, acLineStartCol + 3);
        const acLineCell = worksheet.getCell(4, acLineStartCol);
        acLineCell.value = 'AC LINE LOSS DIFF.';
        acLineCell.font = {
            name: 'Times New Roman',
            size: 10,
            bold: true,
            color: { argb: 'FFFF0000' } // Red text
        };
        acLineCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFfac6c6' } // Light red background
        };
        acLineCell.alignment = {
            horizontal: 'center',
            vertical: 'middle',
            wrapText: true // Add wrapText here as well
        };
        // Add thick outside border to AC LINE cells
        acLineCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        // Add thin inside borders
        for (let i = 1; i <= 3; i++) {
            worksheet.getCell(4, acLineStartCol + i).border.left = { style: 'thin' };
        }

        // Ensure column widths are sufficient for text wrapping
        worksheet.columns.forEach(column => {
            column.width = column.width || 15; // Set default width if not set
        });

        // ====================
        // ROW 5: METER NUMBERS
        // ====================
        const meterRow = worksheet.getRow(5);
        meterRow.height = 40; // Set row height

        // Cell A5 (perfect as is)
        const meterLabelCell = worksheet.getCell('A5');
        meterLabelCell.value = 'METER NO.=>';
        meterLabelCell.font = {
            name: 'Times New Roman',
            size: 10,
            bold: true
        };
        meterLabelCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFbfbfbf' } // Gray background
        };
        meterLabelCell.alignment = {
            horizontal: 'center',
            vertical: 'middle'
        };
        meterLabelCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // Merge main client meter number cells
        const mainMeterEndCol = mainClientEndCol;
        worksheet.mergeCells(`B5:${mainMeterEndCol}5`);
        const mainMeterCell = worksheet.getCell('B5');
        mainMeterCell.value = dailyReport.mainClient.meterNumber;
        mainMeterCell.font = {
            name: 'Times New Roman',
            size: 10,
            bold: true
        };
        mainMeterCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFffc000' } // Orange background
        };
        mainMeterCell.alignment = {
            horizontal: 'center',
            vertical: 'middle'
        };
        // Add thick outside border to main client cells
        mainMeterCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        // Note: No need to set individual cell borders for merged cells

        // Merge meter number cells for each sub-client
        dailyReport.subClient.forEach((subClient, index) => {
            const startCol = 2 + colsPerMainClient + (index * colsPerSubClient);
            worksheet.mergeCells(5, startCol, 5, startCol + colsPerSubClient - 1);
            const cell = worksheet.getCell(5, startCol);
            cell.value = subClient.meterNumber;
            cell.font = {
                name: 'Times New Roman',
                size: 10,
                bold: true
            };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: `FF${subClientColors[index % 2]}` } // Alternating colors
            };
            cell.alignment = {
                horizontal: 'center',
                vertical: 'middle'
            };
            // Add thick outside border to sub-client cells
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            // Add thin inside borders
            for (let i = 1; i < colsPerSubClient; i++) {
                worksheet.getCell(5, startCol + i).border.left = { style: 'thin' };
            }
        });

        // ====================
        // ROW 6: HEADERS
        // ====================
        const headerRow = worksheet.getRow(6);
        headerRow.height = 50; // Adjust height as needed

        // Cell A6
        const dateHeaderCell = worksheet.getCell('A6');
        dateHeaderCell.value = 'DATE';
        dateHeaderCell.font = {
            name: 'Times New Roman',
            size: 9,
            bold: true
        };
        dateHeaderCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD9D9D9' }
        };
        dateHeaderCell.alignment = {
            horizontal: 'center',
            vertical: 'middle',
            wrapText: true // Enable text wrapping
        };
        dateHeaderCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // Cell B6
        const exportHeaderCell = worksheet.getCell('B6');
        exportHeaderCell.value = 'Export';
        exportHeaderCell.font = {
            name: 'Times New Roman',
            size: 9,
            bold: true
        };
        exportHeaderCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFffc000' }
        };
        exportHeaderCell.alignment = {
            horizontal: 'center',
            vertical: 'middle',
            wrapText: true
        };
        exportHeaderCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // Cell B6 - Export (already set above)
        // Dynamic column headers for main client
        let mainHeaderColIndex = 3; // Start from column C

        if (showAvgColumn) {
            const avgHeaderCell = worksheet.getCell(6, mainHeaderColIndex);
            avgHeaderCell.value = 'Avg. Gen.';
            avgHeaderCell.font = {
                name: 'Times New Roman',
                size: 9,
                bold: true
            };
            avgHeaderCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFffc000' }
            };
            avgHeaderCell.alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
            };
            avgHeaderCell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            mainHeaderColIndex++;
        }

        if (showAvgAcColumn) {
            const avgAcHeaderCell = worksheet.getCell(6, mainHeaderColIndex);
            avgAcHeaderCell.value = 'Avg. Gen.';
            avgAcHeaderCell.font = {
                name: 'Times New Roman',
                size: 9,
                bold: true
            };
            avgAcHeaderCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFffc000' }
            };
            avgAcHeaderCell.alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
            };
            avgAcHeaderCell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            mainHeaderColIndex++;
        }

        // Import header
        const importCol = String.fromCharCode(64 + mainHeaderColIndex);
        const importHeaderCell = worksheet.getCell(`${importCol}6`);
        importHeaderCell.value = 'Import';
        importHeaderCell.font = {
            name: 'Times New Roman',
            size: 9,
            bold: true
        };
        importHeaderCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFffc000' }
        };
        importHeaderCell.alignment = {
            horizontal: 'center',
            vertical: 'middle',
            wrapText: true
        };
        importHeaderCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // Sub-client headers
        dailyReport.subClient.forEach((subClient, index) => {
            const startCol = 2 + colsPerMainClient + (index * colsPerSubClient);
            const headers = [];
            let colOffset = 0;

            // Export (always first)
            headers.push({ text: 'Export', color: '000000', offset: colOffset });
            colOffset++;

            // Avg Generation DC (if enabled)
            if (showAvgColumn) {
                headers.push({ text: 'Avg. Gen.', color: '000000', offset: colOffset });
                colOffset++;
            }

            // Avg Generation AC (if enabled)
            if (showAvgAcColumn) {
                headers.push({ text: 'Avg. Gen.', color: '000000', offset: colOffset });
                colOffset++;
            }

            // Logger columns (if enabled)
            if (showLoggerColumns) {
                headers.push({ text: 'Logger Data', color: '000000', offset: colOffset });
                colOffset++;
                headers.push({ text: 'Internal Loss', color: 'FF0000', offset: colOffset });
                colOffset++;
                headers.push({ text: 'Loss in%', color: 'FF0000', offset: colOffset });
                colOffset++;
            }

            // Import (always last)
            headers.push({ text: 'Import', color: '000000', offset: colOffset });

            headers.forEach((header) => {
                const cell = worksheet.getCell(6, startCol + header.offset);
                cell.value = header.text;
                cell.font = {
                    name: 'Times New Roman',
                    size: 10,
                    bold: true,
                    color: { argb: header.color }
                };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: `FF${subClientColors[index % 2]}` } // Alternating colors
                };
                cell.alignment = {
                    horizontal: 'center',
                    vertical: 'middle',
                    wrapText: true // Enable text wrapping
                };
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        });

        // AC LINE LOSS DIFF headers
        const acLineHeaders = [
            { text: 'Export', color: 'FF0000' },
            { text: 'Loss in%', color: 'FF0000' },
            { text: 'Import', color: 'FF0000' },
            { text: 'Loss in%', color: 'FF0000' }
        ];

        acLineHeaders.forEach((header, i) => {
            const cell = worksheet.getCell(6, acLineStartCol + i);
            cell.value = header.text;
            cell.font = {
                name: 'Times New Roman',
                size: 10,
                bold: true,
                color: { argb: header.color }
            };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFfac6c6' } // Light red background
            };
            cell.alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true // Enable text wrapping
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // ====================
        // DATA ROWS (7 to last)
        // ====================
        const dataStyle = {
            font: {
                name: 'Times New Roman',
                size: 9
            },
            alignment: {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
            },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            },
            numFmt: '0.0' // Format numbers to show one decimal place
        };

        // Define background colors
        const bgColors = {
            mainClient: 'FFffc000', // Orange
            subClient1: 'FFFFFF00', // Yellow
            subClient2: 'FF92D050', // Green
            acLine: 'FFfac6c6',     // Light red
            negativeLoss: 'FFffc7ce', // Light red for negative loss
            defaultWhite: 'FFFFFFFF' // White for default
        };

        let dataRowNum = 7;
        dailyReport.mainClient.loggerdatas.forEach(dayData => {
            const row = worksheet.getRow(dataRowNum);
            row.height = 24;
            // Date cell
            const dateCell = row.getCell(1);
            dateCell.value = dayData.date;
            Object.assign(dateCell, dataStyle);
            dateCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD9D9D9' } // Light gray
            };

            // Main client data (B:C or B:C:D depending on showAvgColumn)
            // Export (always B/2)
            const exportCell = row.getCell(2);
            exportCell.value = typeof dayData.export === 'number' ? parseFloat(dayData.export.toFixed(1)) : dayData.export;
            Object.assign(exportCell, dataStyle);
            exportCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: bgColors.mainClient }
            };

            // Dynamic column index for main client
            let mainDataColIndex = 3;

            // Avg Generation DC (if enabled)
            if (showAvgColumn) {
                const avgCell = row.getCell(mainDataColIndex);
                avgCell.value = typeof dayData.avgGeneration === 'number' ? parseFloat(dayData.avgGeneration.toFixed(2)) : (dayData.avgGeneration || 0);
                Object.assign(avgCell, dataStyle);
                avgCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: bgColors.mainClient }
                };
                mainDataColIndex++;
            }

            // Avg Generation AC (if enabled)
            if (showAvgAcColumn) {
                const avgAcCell = row.getCell(mainDataColIndex);
                avgAcCell.value = typeof dayData.avgGenerationAc === 'number' ? parseFloat(dayData.avgGenerationAc.toFixed(2)) : (dayData.avgGenerationAc || 0);
                Object.assign(avgAcCell, dataStyle);
                avgAcCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: bgColors.mainClient }
                };
                mainDataColIndex++;
            }

            // Import
            const importCol = mainDataColIndex;
            const importCell = row.getCell(importCol);
            importCell.value = typeof dayData.import === 'number' ? parseFloat(dayData.import.toFixed(1)) : dayData.import;
            Object.assign(importCell, dataStyle);
            importCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: bgColors.mainClient }
            };

            // Sub-client data
            dailyReport.subClient.forEach((subClient, index) => {
                const subDayData = subClient.loggerdatas.find(d => d.date === dayData.date);
                const startCol = 2 + colsPerMainClient + (index * colsPerSubClient);
                const color = index % 2 === 0 ? bgColors.subClient1 : bgColors.subClient2;
                let colOffset = 0;

                if (subDayData) {
                    // Export cell (always shown)
                    const exportCell = row.getCell(startCol + colOffset);
                    exportCell.value = typeof subDayData.export === 'number' ? parseFloat(subDayData.export.toFixed(1)) : subDayData.export;
                    Object.assign(exportCell, dataStyle);
                    exportCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
                    colOffset++;

                    // Avg Generation DC (if enabled)
                    if (showAvgColumn) {
                        const avgCell = row.getCell(startCol + colOffset);
                        avgCell.value = typeof subDayData.avgGeneration === 'number' ? parseFloat(subDayData.avgGeneration.toFixed(2)) : (subDayData.avgGeneration || 0);
                        Object.assign(avgCell, dataStyle);
                        avgCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
                        colOffset++;
                    }

                    // Avg Generation AC (if enabled)
                    if (showAvgAcColumn) {
                        const avgAcCell = row.getCell(startCol + colOffset);
                        avgAcCell.value = typeof subDayData.avgGenerationAc === 'number' ? parseFloat(subDayData.avgGenerationAc.toFixed(2)) : (subDayData.avgGenerationAc || 0);
                        Object.assign(avgAcCell, dataStyle);
                        avgAcCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
                        colOffset++;
                    }

                    if (showLoggerColumns) {
                        // Logger Data
                        const loggerDataCell = row.getCell(startCol + colOffset);
                        loggerDataCell.value = typeof subDayData.loggerdata === 'number' ? parseFloat(subDayData.loggerdata.toFixed(1)) : subDayData.loggerdata;
                        Object.assign(loggerDataCell, dataStyle);
                        loggerDataCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
                        colOffset++;

                        // Internal Loss (red font only when negative)
                        const internalLossCell = row.getCell(startCol + colOffset);
                        internalLossCell.value = typeof subDayData.internallosse === 'number' ? parseFloat(subDayData.internallosse.toFixed(1)) : subDayData.internallosse;
                        Object.assign(internalLossCell, dataStyle);
                        internalLossCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };

                        // Set font color based on value (red if negative, black otherwise)
                        const isNegativeInternalLoss = subDayData.internallosse < 0;
                        internalLossCell.font = {
                            ...dataStyle.font,
                            color: { argb: isNegativeInternalLoss ? 'FFFF0000' : 'FF000000' }
                        };
                        colOffset++;

                        // Loss in%
                        const lossPercentCell = row.getCell(startCol + colOffset);
                        const lossPercentValue = typeof subDayData.lossinparsantege === 'number' ? parseFloat(subDayData.lossinparsantege.toFixed(1)) : subDayData.lossinparsantege;
                        lossPercentCell.value = typeof lossPercentValue === 'number' ? `${lossPercentValue}%` : lossPercentValue;
                        Object.assign(lossPercentCell, dataStyle);

                        const isNegativeLoss = subDayData.lossinparsantege < 0;
                        // For negative values: light red background, red text
                        // For positive values: use the alternating color (green/yellow)
                        lossPercentCell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: isNegativeLoss ? bgColors.negativeLoss : color }
                        };
                        lossPercentCell.font = {
                            ...dataStyle.font,
                            color: { argb: isNegativeLoss ? 'FFFF0000' : 'FF000000' }
                        };
                        colOffset++;
                    }

                    // Import (always last)
                    const importCell = row.getCell(startCol + colOffset);
                    importCell.value = typeof subDayData.import === 'number' ? parseFloat(subDayData.import.toFixed(1)) : subDayData.import;
                    Object.assign(importCell, dataStyle);
                    importCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
                }
            });

            // AC Line Loss Diff data
            const acLineDayData = dailyReport.aclinelossdiffrence.loggerdatas.find(d => d.date === dayData.date);
            if (acLineDayData) {
                const values = [
                    { value: acLineDayData.export, col: 0, isLossPercent: false },
                    { value: acLineDayData.lossinparsantegeexport, col: 1, isLossPercent: true },
                    { value: acLineDayData.import, col: 2, isLossPercent: false },
                    { value: acLineDayData.lossinparsantegeimport, col: 3, isLossPercent: true }
                ];

                values.forEach(item => {
                    const cell = row.getCell(acLineStartCol + item.col);
                    const formattedValue = typeof item.value === 'number' ?
                        (item.isLossPercent ? `${parseFloat(item.value.toFixed(1))}%` : parseFloat(item.value.toFixed(1))) :
                        item.value;
                    cell.value = formattedValue;
                    Object.assign(cell, dataStyle);

                    // New condition for AC LINE LOSS DIFF
                    if (item.value < 0) {
                        // Negative values: light red background, red text
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColors.acLine } };
                        cell.font = { ...dataStyle.font, color: { argb: 'FFFF0000' } };
                    } else {
                        // Positive values: white background, black text
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColors.defaultWhite } };
                        cell.font = { ...dataStyle.font, color: { argb: 'FF000000' } };
                    }
                });
            }

            dataRowNum++;
        });

        // ====================
        // TOTALS ROW
        // ====================
        const totalsRow = worksheet.getRow(dataRowNum);
        totalsRow.height = 40; // Set row height
        const totalsStyle = {
            font: {
                name: 'Times New Roman',
                size: 10,
                bold: true
            },
            alignment: {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
            },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            },
            numFmt: '0.0' // Format numbers to show one decimal place
        };

        // Add thick outside border to totals row
        for (let col = 1; col <= totalColss; col++) {
            const cell = totalsRow.getCell(col);
            cell.border = {
                top: { style: 'thin' },
                left: { style: col === 1 ? 'thin' : 'thin' },
                bottom: { style: 'thin' },
                right: { style: col === totalColss ? 'thin' : 'thin' }
            };
        }

        // Label
        totalsRow.getCell(1).value = 'Total';
        Object.assign(totalsRow.getCell(1), totalsStyle);
        totalsRow.getCell(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD9D9D9' } // Light gray background
        };

        // Main client totals
        // Export (always B/2)
        const mainExportCell = totalsRow.getCell(2);
        mainExportCell.value = typeof dailyReport.mainClient.totalexport === 'number' ? parseFloat(dailyReport.mainClient.totalexport.toFixed(1)) : dailyReport.mainClient.totalexport;
        Object.assign(mainExportCell, totalsStyle);
        mainExportCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: bgColors.mainClient }
        };

        // Dynamic column index for main client totals
        let mainTotalColIndex = 3;

        // Avg Generation DC total (if enabled)
        if (showAvgColumn) {
            const mainAvgCell = totalsRow.getCell(mainTotalColIndex);
            mainAvgCell.value = typeof dailyReport.mainClient.totalAvgGeneration === 'number' ? parseFloat(dailyReport.mainClient.totalAvgGeneration.toFixed(2)) : (dailyReport.mainClient.totalAvgGeneration || 0);
            Object.assign(mainAvgCell, totalsStyle);
            mainAvgCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: bgColors.mainClient }
            };
            mainTotalColIndex++;
        }

        // Avg Generation AC total (if enabled)
        if (showAvgAcColumn) {
            const mainAvgAcCell = totalsRow.getCell(mainTotalColIndex);
            mainAvgAcCell.value = typeof dailyReport.mainClient.totalAvgGenerationAc === 'number' ? parseFloat(dailyReport.mainClient.totalAvgGenerationAc.toFixed(2)) : (dailyReport.mainClient.totalAvgGenerationAc || 0);
            Object.assign(mainAvgAcCell, totalsStyle);
            mainAvgAcCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: bgColors.mainClient }
            };
            mainTotalColIndex++;
        }

        // Import
        const mainImportCol = mainTotalColIndex;
        const mainImportCell = totalsRow.getCell(mainImportCol);
        mainImportCell.value = typeof dailyReport.mainClient.totalimport === 'number' ? parseFloat(dailyReport.mainClient.totalimport.toFixed(1)) : dailyReport.mainClient.totalimport;
        Object.assign(mainImportCell, totalsStyle);
        mainImportCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: bgColors.mainClient }
        };

        // Sub-client totals
        dailyReport.subClient.forEach((subClient, index) => {
            const startCol = 2 + colsPerMainClient + (index * colsPerSubClient);
            const color = index % 2 === 0 ? bgColors.subClient1 : bgColors.subClient2;
            let colOffset = 0;

            // Export (always shown)
            const exportCell = totalsRow.getCell(startCol + colOffset);
            exportCell.value = typeof subClient.totalexport === 'number' ? parseFloat(subClient.totalexport.toFixed(1)) : subClient.totalexport;
            Object.assign(exportCell, totalsStyle);
            exportCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
            colOffset++;

            // Avg Generation DC total (if enabled)
            if (showAvgColumn) {
                const avgCell = totalsRow.getCell(startCol + colOffset);
                avgCell.value = typeof subClient.totalAvgGeneration === 'number' ? parseFloat(subClient.totalAvgGeneration.toFixed(2)) : (subClient.totalAvgGeneration || 0);
                Object.assign(avgCell, totalsStyle);
                avgCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
                colOffset++;
            }

            // Avg Generation AC total (if enabled)
            if (showAvgAcColumn) {
                const avgAcCell = totalsRow.getCell(startCol + colOffset);
                avgAcCell.value = typeof subClient.totalAvgGenerationAc === 'number' ? parseFloat(subClient.totalAvgGenerationAc.toFixed(2)) : (subClient.totalAvgGenerationAc || 0);
                Object.assign(avgAcCell, totalsStyle);
                avgAcCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
                colOffset++;
            }

            if (showLoggerColumns) {
                // Logger Data
                const loggerDataCell = totalsRow.getCell(startCol + colOffset);
                loggerDataCell.value = typeof subClient.totalloggerdata === 'number' ? parseFloat(subClient.totalloggerdata.toFixed(1)) : subClient.totalloggerdata;
                Object.assign(loggerDataCell, totalsStyle);
                loggerDataCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
                colOffset++;

                // Internal Loss (red font only when negative)
                const internalLossCell = totalsRow.getCell(startCol + colOffset);
                internalLossCell.value = typeof subClient.totalinternallosse === 'number' ? parseFloat(subClient.totalinternallosse.toFixed(1)) : subClient.totalinternallosse;
                Object.assign(internalLossCell, totalsStyle);
                internalLossCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };

                // Set font color based on value (red if negative, black otherwise)
                const isNegativeInternalLoss = subClient.totalinternallosse < 0;
                internalLossCell.font = {
                    ...totalsStyle.font,
                    color: { argb: isNegativeInternalLoss ? 'FFFF0000' : 'FF000000' }
                };
                colOffset++;

                // Loss in% total
                const lossPercentCell = totalsRow.getCell(startCol + colOffset);
                const lossPercentValue = typeof subClient.totallossinparsantege === 'number' ? parseFloat(subClient.totallossinparsantege.toFixed(1)) : subClient.totallossinparsantege;
                lossPercentCell.value = typeof lossPercentValue === 'number' ? `${lossPercentValue}%` : lossPercentValue;
                Object.assign(lossPercentCell, totalsStyle);

                const isNegativeLoss = subClient.totallossinparsantege < 0;
                // For negative values: light red background, red text
                // For positive values: use the alternating color (green/yellow)
                lossPercentCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: isNegativeLoss ? bgColors.negativeLoss : color }
                };
                lossPercentCell.font = {
                    ...totalsStyle.font,
                    color: { argb: isNegativeLoss ? 'FFFF0000' : 'FF000000' }
                };
                colOffset++;
            }

            // Import (always last)
            const importCell = totalsRow.getCell(startCol + colOffset);
            importCell.value = typeof subClient.totalimport === 'number' ? parseFloat(subClient.totalimport.toFixed(1)) : subClient.totalimport;
            Object.assign(importCell, totalsStyle);
            importCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
        });

        // AC Line Loss Diff totals
        const acLineTotals = [
            { value: dailyReport.aclinelossdiffrence.totalexport, col: 0, isLossPercent: false },
            { value: dailyReport.aclinelossdiffrence.totallossinparsantegeexport, col: 1, isLossPercent: true },
            { value: dailyReport.aclinelossdiffrence.totalimport, col: 2, isLossPercent: false },
            { value: dailyReport.aclinelossdiffrence.totallossinparsantegeimport, col: 3, isLossPercent: true }
        ];

        acLineTotals.forEach(item => {
            const cell = totalsRow.getCell(acLineStartCol + item.col);
            const formattedValue = typeof item.value === 'number' ?
                (item.isLossPercent ? `${parseFloat(item.value.toFixed(1))}%` : parseFloat(item.value.toFixed(1))) :
                item.value;
            cell.value = formattedValue;
            Object.assign(cell, totalsStyle);

            // Apply same condition to totals row
            if (item.value < 0) {
                // Negative values: light red background, red text
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColors.acLine } };
                cell.font = { ...totalsStyle.font, color: { argb: 'FFFF0000' } };
            } else {
                // Positive values: white background, black text
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColors.defaultWhite } };
                cell.font = { ...totalsStyle.font, color: { argb: 'FF000000' } };
            }
        });

        // ====================
        // NOTE ROW - Capacity Basis
        // ====================
        // Only add note if at least one avg column is shown
        if (showAvgColumn || showAvgAcColumn) {
            const noteRowNum = dataRowNum + 1; // One row after totals
            const noteRow = worksheet.getRow(noteRowNum);
            noteRow.height = 25; // Slightly taller for readability

            // Determine which note to display based on what's shown and what capacity was used
            let noteText = '';
            if (showAvgColumn && dailyReport.capacityBasisDC) {
                noteText = `Note: The Average Generation has been calculated on the basis of the plant ${dailyReport.capacityBasisDC} capacity.`;
            } else if (showAvgAcColumn && dailyReport.capacityBasisAC) {
                noteText = `Note: The Average Generation has been calculated on the basis of the plant ${dailyReport.capacityBasisAC} capacity.`;
            }

            if (noteText) {
                // Merge cells across entire width (same as row 2)
                worksheet.mergeCells(noteRowNum, 1, noteRowNum, totalColss);
                const noteCell = worksheet.getCell(noteRowNum, 1);
                noteCell.value = noteText;
                noteCell.font = {
                    name: 'Times New Roman',
                    size: 10,
                    bold: true,
                    italic: true,
                    color: { argb: 'FF000000' } // Black text
                };
                noteCell.alignment = {
                    horizontal: 'center',
                    vertical: 'middle'
                };
                noteCell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            }
        }

        // Set response headers
        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );

        // Format the filename components
        const formattedMonthName = getMonthName(dailyReport.month);
        const formattedYear = dailyReport.year;
        const clientName = dailyReport.mainClient.mainClientDetail?.name?.trim() || '';

        // Sanitize filename components (remove special characters and extra spaces)
        const sanitize = (str) => str.replace(/[^a-zA-Z0-9\s-]/g, '').replace(/\s+/g, ' ').trim();

        // Build filename in required format
        const sanitizedClientName = sanitize(clientName);
        const fileName = `Daily Generation Report - ${sanitizedClientName} Month of ${formattedMonthName}-${formattedYear}.xlsx`;

        // Set response headers
        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            `attachment; filename="${encodeURIComponent(fileName)}"`
        );
        res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        logger.error(`Error generating Daily Report Excel: ${error.message}`);
        res.status(500).json({
            message: 'Error generating Daily Report Excel',
            error: error.message
        });
    }
};


// Full working version with portrait layout and scaled columns
exports.downloadDailyReportPDF = async (req, res) => {
    try {
        const { dailyReportId } = req.params;
        const dailyReport = await DailyReport.findById(dailyReportId)
            .populate("mainClientId")
            .populate("subClient.subClientId");

        if (!dailyReport) {
            return res.status(404).json({ message: "Daily Report not found" });
        }

        const filemonthName = getMonthName(dailyReport.month).toUpperCase();
        const year = dailyReport.year;
        const mainClientName = dailyReport.mainClient.mainClientDetail.name
            .replace(/\s+/g, "_")
            .replace(/\//g, "_")
            .toUpperCase();

        const fileName = `Daily Generation Report - ${mainClientName} Month of ${filemonthName}-${year}.pdf`;

        res.setHeader("Content-Type", "application/pdf");
        res.setHeader(
            "Content-Disposition",
            `attachment; filename="${encodeURIComponent(fileName)}"`
        );
        res.setHeader("Access-Control-Expose-Headers", "Content-Disposition");

        const doc = new PDFDocument({
            size: "A4",
            layout: "landscape",
            margins: { top: 10, bottom: 10, left: 15, right: 15 },
        });
        doc.pipe(res);

        const subClientCount = Math.max(dailyReport.subClient.length, 2);
        const showLoggerColumns = dailyReport.showLoggerColumns === true; // Default to false
        const showAvgColumn = dailyReport.showAvgColumn === true; // Default to false (DC Capacity)
        const showAvgAcColumn = dailyReport.showAvgAcColumn === true; // Default to false (AC Capacity)
        // Column structure: Export + (Avg DC?) + (Avg AC?) + (Logger Data + Internal Loss + Loss in%?) + Import
        const colsPerSubClient = 1 + (showAvgColumn ? 1 : 0) + (showAvgAcColumn ? 1 : 0) + (showLoggerColumns ? 3 : 0) + 1; // Export + Avg DC + Avg AC + Logger cols + Import
        const colsPerMainClient = 1 + (showAvgColumn ? 1 : 0) + (showAvgAcColumn ? 1 : 0) + 1; // Export + Avg DC + Avg AC + Import

        const titleFont = "Helvetica-Bold";
        const headerFont = "Helvetica-Bold";
        const dataFont = "Helvetica";
        const titleFontSize = 13;
        const headerFontSize = subClientCount > 2 ? Math.max(5, 7.5 - (subClientCount - 2) * 0.2) : 7.5;
        const dataFontSize = subClientCount > 2 ? Math.max(5, 6.5 - (subClientCount - 2) * 0.8) : 6.5;
        const totalFontSize = subClientCount > 2 ? Math.max(5, 7 - (subClientCount - 2) * 0.3) : 7

        console.log(totalFontSize, '<--totalFontSize');

        const colors = {
            titleBg: "#bdd7ee",
            monthBg: "#bfbfbf",
            mainClientBg: "#ffc000",
            subClient1Bg: "#FFFF00",
            subClient2Bg: "#92D050",
            acLineBg: "#fac6c6",
            negativeLossBg: "#ffc7ce",
            redText: "#FF0000",
            blackText: "#000000",
            whiteBg: "#FFFFFF",
            borderColor: "#000000",
            grayBg: "#D9D9D9",
        };

        const rowHeights = {
            title: 20,
            monthDate: 20,
            clientHeader: 28,
            meterHeader: 22,
            dataHeader: subClientCount > 2 ? 24 : 16,
            dataRow: 13.5,
            totalsRow: 15,
        };

        // === Dynamic width logic ===
        const usableWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
        const totalCols = 1 + colsPerMainClient + (colsPerSubClient * subClientCount) + 4;
        const colWidth = usableWidth / totalCols;

        const colWidths = {
            date: colWidth,
            mainExport: colWidth,
            mainAvg: showAvgColumn ? colWidth : 0,
            mainAvgAc: showAvgAcColumn ? colWidth : 0,
            mainImport: colWidth,
            subExport: colWidth,
            subData: colWidth,
            subLoss: colWidth,
            subPercent: colWidth,
            subImport: colWidth,
            acLine: colWidth,
        };

        const totalWidth = colWidth * totalCols;
        const startX = doc.page.margins.left;

        let currentY = 20;

        // Render title
        doc.lineWidth(0.5)
            .rect(startX, currentY, totalWidth, rowHeights.title)
            .fillAndStroke(colors.titleBg, colors.borderColor);

        doc.font(titleFont)
            .fontSize(titleFontSize)
            .fillColor(colors.blackText)
            .text(
                `${dailyReport.mainClient.mainClientDetail.name} - ${(dailyReport.mainClient.mainClientDetail.acCapacityKw / 1000).toFixed(2)} MW AC Generation Details`,
                startX,
                currentY + 7,
                { width: totalWidth, align: "center" }
            );

        currentY += rowHeights.title;

        // ====================
        // MONTH AND DATE RANGE ROW
        // ====================
        const monthName = getMonthName(dailyReport.month);
        const yearShort = dailyReport.year.toString().slice(-2);
        const lastDay = new Date(dailyReport.year, dailyReport.month, 0).getDate();
        const dateRange = `01-${String(dailyReport.month).padStart(2, '0')}-${dailyReport.year} to ${lastDay}-${String(dailyReport.month).padStart(2, '0')}-${dailyReport.year}`;

        // Calculate widths for month and date range sections
        const mainClientNameWidth = colWidths.mainExport + (showAvgColumn ? colWidths.mainAvg : 0) + (showAvgAcColumn ? colWidths.mainAvgAc : 0) + colWidths.mainImport;
        const monthWidth = colWidths.date + mainClientNameWidth +
            (colsPerSubClient * colWidths.subExport); // Width for one sub-client section
        const dateRangeWidth = totalWidth - monthWidth;

        // Month cell
        doc.rect(15, currentY, monthWidth, rowHeights.monthDate)
            .lineWidth(0.5)
            .fillAndStroke(colors.monthBg, colors.borderColor)

        doc.font(headerFont)
            .fontSize(14)
            .fillColor(colors.redText)
            .text(`Month: ${monthName}-${yearShort}`, 15, currentY + 7, {
                width: monthWidth,
                align: 'center'
            });

        // Date range cell
        doc.rect(15 + monthWidth, currentY, dateRangeWidth, rowHeights.monthDate)
            .lineWidth(0.5)
            .fillAndStroke(colors.monthBg, colors.borderColor);

        doc.font(headerFont)
            .fontSize(14)
            .fillColor(colors.redText)
            .text(`Generation Period: ${dateRange}`, 15 + monthWidth, currentY + 7, {
                width: dateRangeWidth,
                align: 'center'
            });

        currentY += rowHeights.monthDate;

        // ====================
        // CLIENT NAMES ROW
        // ====================
        // NAME=> cell
        doc.rect(15, currentY, colWidths.date, rowHeights.clientHeader)
            .lineWidth(0.5)
            .fillAndStroke(colors.grayBg, colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('NAME=>', 15, currentY + 12, {
                width: colWidths.date,
                align: 'center'
            });

        // TOTAL-GETCO SS cell (merged)
        doc.rect(15 + colWidths.date, currentY, mainClientNameWidth, rowHeights.clientHeader)
            .lineWidth(0.5)
            .fillAndStroke(colors.mainClientBg, colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('TOTAL-GETCO\nSS', 15 + colWidths.date, currentY + 7, {
                width: mainClientNameWidth,
                align: 'center',
                lineGap: 3
            });

        // Sub-client cells
        dailyReport.subClient.forEach((subClient, index) => {
            const startX = 15 + colWidths.date + mainClientNameWidth + (index * colsPerSubClient * colWidths.subExport);
            const color = index % 2 === 0 ? colors.subClient1Bg : colors.subClient2Bg;

            doc.rect(startX, currentY, colsPerSubClient * colWidths.subExport, rowHeights.clientHeader)
                .lineWidth(0.5)
                .fillAndStroke(color, colors.borderColor);

            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.blackText)
                .text(subClient.name, startX, currentY + 12, {
                    width: colsPerSubClient * colWidths.subExport,
                    align: 'center'
                });
        });

        // AC LINE LOSS DIFF header (merged with next row)
        const acLineStartX = 15 + colWidths.date + mainClientNameWidth +
            (dailyReport.subClient.length * colsPerSubClient * colWidths.subExport);

        doc.rect(acLineStartX, currentY, 4 * colWidths.acLine, rowHeights.clientHeader * 2)
            .lineWidth(0.5)
            .fillAndStroke(colors.acLineBg, colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.redText)
            .text('AC LINE LOSS DIFF.', acLineStartX, currentY + 25, {
                width: 4 * colWidths.acLine,
                align: 'center'
            });

        currentY += rowHeights.clientHeader;

        // ====================
        // METER NUMBERS ROW
        // ====================
        // METER NO.=> cell
        doc.rect(15, currentY, colWidths.date, rowHeights.meterHeader)
            .lineWidth(0.5)
            .fillAndStroke(colors.grayBg, colors.borderColor);

        const cellHeight = rowHeights.meterHeader;      // height of the current row
        const fontSize = headerFontSize;                // font size used for this text
        const yOffset = currentY + (cellHeight - fontSize) / 2;

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('METER NO.=>', 15, yOffset - 3, {
                width: colWidths.date,
                align: 'center'
            });

        // Main client meter number (merged)
        doc.rect(15 + colWidths.date, currentY, mainClientNameWidth, rowHeights.meterHeader)
            .lineWidth(0.5)
            .fillAndStroke(colors.mainClientBg, colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text(dailyReport.mainClient.meterNumber, 15 + colWidths.date, currentY + 10, {
                width: mainClientNameWidth,
                align: 'center'
            });

        // Sub-client meter numbers
        dailyReport.subClient.forEach((subClient, index) => {
            const startX = 15 + colWidths.date + mainClientNameWidth + (index * colsPerSubClient * colWidths.subExport);
            const color = index % 2 === 0 ? colors.subClient1Bg : colors.subClient2Bg;

            doc.rect(startX, currentY, colsPerSubClient * colWidths.subExport, rowHeights.meterHeader)
                .lineWidth(0.5)
                .fillAndStroke(color, colors.borderColor);

            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.blackText)
                .text(subClient.meterNumber, startX, currentY + 10, {
                    width: colsPerSubClient * colWidths.subExport,
                    align: 'center'
                });
        });

        currentY += rowHeights.meterHeader;

        // ====================
        // HEADERS ROW
        // ====================
        // DATE header
        doc.rect(15, currentY, colWidths.date, rowHeights.dataHeader)
            .lineWidth(0.5)
            .fillAndStroke(colors.grayBg, colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('DATE', 15, currentY + 7, {
                width: colWidths.date,
                align: 'center'
            });

        // Main client headers
        let mainHeaderOffset = 0;

        // Export
        doc.rect(15 + colWidths.date + (mainHeaderOffset * colWidths.mainExport), currentY, colWidths.mainExport, rowHeights.dataHeader)
            .lineWidth(0.5)
            .fillAndStroke(colors.mainClientBg, colors.borderColor);
        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('Export', 15 + colWidths.date + (mainHeaderOffset * colWidths.mainExport), currentY + 7, {
                width: colWidths.mainExport,
                align: 'center'
            });
        mainHeaderOffset++;

        // Avg Generation DC (if enabled)
        if (showAvgColumn) {
            doc.rect(15 + colWidths.date + (mainHeaderOffset * colWidths.mainExport), currentY, colWidths.mainExport, rowHeights.dataHeader)
                .lineWidth(0.5)
                .fillAndStroke(colors.mainClientBg, colors.borderColor);
            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.blackText)
                .text('Avg. Gen.', 15 + colWidths.date + (mainHeaderOffset * colWidths.mainExport), currentY + 7, {
                    width: colWidths.mainExport,
                    align: 'center'
                });
            mainHeaderOffset++;
        }

        // Avg Generation AC (if enabled)
        if (showAvgAcColumn) {
            doc.rect(15 + colWidths.date + (mainHeaderOffset * colWidths.mainExport), currentY, colWidths.mainExport, rowHeights.dataHeader)
                .lineWidth(0.5)
                .fillAndStroke(colors.mainClientBg, colors.borderColor);
            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.blackText)
                .text('Avg. Gen.', 15 + colWidths.date + (mainHeaderOffset * colWidths.mainExport), currentY + 7, {
                    width: colWidths.mainExport,
                    align: 'center'
                });
            mainHeaderOffset++;
        }

        // Import
        doc.rect(15 + colWidths.date + (mainHeaderOffset * colWidths.mainExport), currentY, colWidths.mainImport, rowHeights.dataHeader)
            .lineWidth(0.5)
            .fillAndStroke(colors.mainClientBg, colors.borderColor);
        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('Import', 15 + colWidths.date + (mainHeaderOffset * colWidths.mainExport), currentY + 7, {
                width: colWidths.mainImport,
                align: 'center'
            });

        // Sub-client headers
        dailyReport.subClient.forEach((subClient, index) => {
            const startX = 15 + colWidths.date + mainClientNameWidth + (index * colsPerSubClient * colWidths.subExport);
            const color = index % 2 === 0 ? colors.subClient1Bg : colors.subClient2Bg;
            let colOffset = 0;

            // Export (always first)
            doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataHeader)
                .lineWidth(0.5)
                .fillAndStroke(color, colors.borderColor);
            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.blackText)
                .text('Export', startX + (colOffset * colWidths.subExport), currentY + 7, {
                    width: colWidths.subExport,
                    align: 'center'
                });
            colOffset++;

            // Avg Generation DC (if enabled)
            if (showAvgColumn) {
                doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataHeader)
                    .lineWidth(0.5)
                    .fillAndStroke(color, colors.borderColor);
                doc.font(headerFont)
                    .fontSize(headerFontSize)
                    .fillColor(colors.blackText)
                    .text('Avg. Gen.', startX + (colOffset * colWidths.subExport), currentY + 7, {
                        width: colWidths.subExport,
                        align: 'center'
                    });
                colOffset++;
            }

            // Avg Generation AC (if enabled)
            if (showAvgAcColumn) {
                doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataHeader)
                    .lineWidth(0.5)
                    .fillAndStroke(color, colors.borderColor);
                doc.font(headerFont)
                    .fontSize(headerFontSize)
                    .fillColor(colors.blackText)
                    .text('Avg. Gen.', startX + (colOffset * colWidths.subExport), currentY + 7, {
                        width: colWidths.subExport,
                        align: 'center'
                    });
                colOffset++;
            }

            // Logger columns (if enabled)
            if (showLoggerColumns) {
                doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataHeader)
                    .lineWidth(0.5)
                    .fillAndStroke(color, colors.borderColor);
                doc.font(headerFont)
                    .fontSize(headerFontSize)
                    .fillColor(colors.blackText)
                    .text('Logger Data', startX + (colOffset * colWidths.subExport), currentY + 7, {
                        width: colWidths.subExport,
                        align: 'center'
                    });
                colOffset++;

                doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataHeader)
                    .lineWidth(0.5)
                    .fillAndStroke(color, colors.borderColor);
                doc.font(headerFont)
                    .fontSize(headerFontSize)
                    .fillColor(colors.redText)
                    .text('Internal Loss', startX + (colOffset * colWidths.subExport), currentY + 7, {
                        width: colWidths.subExport,
                        align: 'center'
                    });
                colOffset++;

                doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataHeader)
                    .lineWidth(0.5)
                    .fillAndStroke(color, colors.borderColor);
                doc.font(headerFont)
                    .fontSize(headerFontSize)
                    .fillColor(colors.redText)
                    .text('Loss in%', startX + (colOffset * colWidths.subExport), currentY + 7, {
                        width: colWidths.subExport,
                        align: 'center'
                    });
                colOffset++;
            }

            // Import (always last)
            doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataHeader)
                .lineWidth(0.5)
                .fillAndStroke(color, colors.borderColor);
            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.blackText)
                .text('Import', startX + (colOffset * colWidths.subExport), currentY + 7, {
                    width: colWidths.subExport,
                    align: 'center'
                });
        });

        // AC LINE LOSS DIFF headers
        const acLineHeaders = ['Export', 'Loss in%', 'Import', 'Loss in%'];
        acLineHeaders.forEach((header, i) => {
            doc.rect(acLineStartX + (i * colWidths.acLine), currentY, colWidths.acLine, rowHeights.dataHeader)
                .lineWidth(0.5)
                .fillAndStroke(colors.acLineBg, colors.borderColor);

            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.redText)
                .text(header, acLineStartX + (i * colWidths.acLine), currentY + 7, {
                    width: colWidths.acLine,
                    align: 'center'
                });
        });

        currentY += rowHeights.dataHeader;

        // ====================
        // DATA ROWS
        // ====================


        dailyReport.mainClient.loggerdatas.forEach(dayData => {
            // Check if we need a new page
            if (currentY + rowHeights.dataRow > doc.page.height - 20) {
                doc.addPage({
                    size: 'A4',
                    layout: 'landscape',
                    margins: { top: 20, bottom: 20, left: 15, right: 15 }
                });
                currentY = 20;
            }
            const cellHeight = rowHeights.dataRow;
            const fontSize = dataFontSize;
            const yCentered = (currentY + (cellHeight - fontSize) / 2) + 2;
            // Date cell
            doc.rect(15, currentY, colWidths.date, rowHeights.dataRow)
                .lineWidth(0.5)
                .fillAndStroke(colors.grayBg, colors.borderColor);

            doc.font(dataFont)
                .fontSize(dataFontSize)
                .fillColor(colors.blackText)
                .text(dayData.date, 15, yCentered, {
                    width: colWidths.date,
                    align: 'center'
                });

            // Main client data
            let mainDataOffset = 0;

            // Export
            doc.rect(15 + colWidths.date + (mainDataOffset * colWidths.mainExport), currentY, colWidths.mainExport, rowHeights.dataRow)
                .lineWidth(0.5)
                .fillAndStroke(colors.mainClientBg, colors.borderColor);
            doc.font(dataFont)
                .fontSize(dataFontSize)
                .fillColor(colors.blackText)
                .text(typeof dayData.export === 'number' ? dayData.export.toFixed(1) : dayData.export,
                    15 + colWidths.date + (mainDataOffset * colWidths.mainExport), yCentered, {
                    width: colWidths.mainExport,
                    align: 'center'
                });
            mainDataOffset++;

            // Avg Generation DC (if enabled)
            if (showAvgColumn) {
                doc.rect(15 + colWidths.date + (mainDataOffset * colWidths.mainExport), currentY, colWidths.mainExport, rowHeights.dataRow)
                    .lineWidth(0.5)
                    .fillAndStroke(colors.mainClientBg, colors.borderColor);
                doc.font(dataFont)
                    .fontSize(dataFontSize)
                    .fillColor(colors.blackText)
                    .text(typeof dayData.avgGeneration === 'number' ? dayData.avgGeneration.toFixed(1) : (dayData.avgGeneration || '0.0'),
                        15 + colWidths.date + (mainDataOffset * colWidths.mainExport), yCentered, {
                        width: colWidths.mainExport,
                        align: 'center'
                    });
                mainDataOffset++;
            }

            // Avg Generation AC (if enabled)
            if (showAvgAcColumn) {
                doc.rect(15 + colWidths.date + (mainDataOffset * colWidths.mainExport), currentY, colWidths.mainExport, rowHeights.dataRow)
                    .lineWidth(0.5)
                    .fillAndStroke(colors.mainClientBg, colors.borderColor);
                doc.font(dataFont)
                    .fontSize(dataFontSize)
                    .fillColor(colors.blackText)
                    .text(typeof dayData.avgGenerationAc === 'number' ? dayData.avgGenerationAc.toFixed(1) : (dayData.avgGenerationAc || '0.0'),
                        15 + colWidths.date + (mainDataOffset * colWidths.mainExport), yCentered, {
                        width: colWidths.mainExport,
                        align: 'center'
                    });
                mainDataOffset++;
            }

            // Import
            doc.rect(15 + colWidths.date + (mainDataOffset * colWidths.mainExport), currentY, colWidths.mainImport, rowHeights.dataRow)
                .lineWidth(0.5)
                .fillAndStroke(colors.mainClientBg, colors.borderColor);
            doc.font(dataFont)
                .fontSize(dataFontSize)
                .fillColor(colors.blackText)
                .text(typeof dayData.import === 'number' ? dayData.import.toFixed(1) : dayData.import,
                    15 + colWidths.date + (mainDataOffset * colWidths.mainExport), yCentered, {
                    width: colWidths.mainImport,
                    align: 'center'
                });

            // Sub-client data
            dailyReport.subClient.forEach((subClient, index) => {
                const subDayData = subClient.loggerdatas.find(d => d.date === dayData.date);
                const startX = 15 + colWidths.date + mainClientNameWidth + (index * colsPerSubClient * colWidths.subExport);
                const color = index % 2 === 0 ? colors.subClient1Bg : colors.subClient2Bg;
                const yCentered = (currentY + (rowHeights.dataRow - dataFontSize) / 2) + 2;
                let colOffset = 0;

                if (subDayData) {
                    // Export (always shown)
                    doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataRow)
                        .lineWidth(0.5)
                        .fillAndStroke(color, colors.borderColor);

                    doc.font(dataFont)
                        .fontSize(dataFontSize)
                        .fillColor(colors.blackText)
                        .text(typeof subDayData.export === 'number' ? subDayData.export.toFixed(1) : subDayData.export,
                            startX + (colOffset * colWidths.subExport), yCentered, {
                            width: colWidths.subExport,
                            align: 'center'
                        });
                    colOffset++;

                    // Avg Generation DC (if enabled)
                    if (showAvgColumn) {
                        doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataRow)
                            .lineWidth(0.5)
                            .fillAndStroke(color, colors.borderColor);

                        doc.font(dataFont)
                            .fontSize(dataFontSize)
                            .fillColor(colors.blackText)
                            .text(typeof subDayData.avgGeneration === 'number' ? subDayData.avgGeneration.toFixed(1) : (subDayData.avgGeneration || 0),
                                startX + (colOffset * colWidths.subExport), yCentered, {
                                width: colWidths.subExport,
                                align: 'center'
                            });
                        colOffset++;
                    }

                    // Avg Generation AC (if enabled)
                    if (showAvgAcColumn) {
                        doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataRow)
                            .lineWidth(0.5)
                            .fillAndStroke(color, colors.borderColor);

                        doc.font(dataFont)
                            .fontSize(dataFontSize)
                            .fillColor(colors.blackText)
                            .text(typeof subDayData.avgGenerationAc === 'number' ? subDayData.avgGenerationAc.toFixed(1) : (subDayData.avgGenerationAc || 0),
                                startX + (colOffset * colWidths.subExport), yCentered, {
                                width: colWidths.subExport,
                                align: 'center'
                            });
                        colOffset++;
                    }

                    if (showLoggerColumns) {
                        // Logger Data
                        doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataRow)
                            .lineWidth(0.5)
                            .fillAndStroke(color, colors.borderColor);

                        doc.font(dataFont)
                            .fontSize(dataFontSize)
                            .fillColor(colors.blackText)
                            .text(typeof subDayData.loggerdata === 'number' ? subDayData.loggerdata.toFixed(1) : subDayData.loggerdata,
                                startX + (colOffset * colWidths.subExport), yCentered, {
                                width: colWidths.subExport,
                                align: 'center'
                            });
                        colOffset++;

                        // Internal Loss (red if negative)
                        doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataRow)
                            .lineWidth(0.5)
                            .fillAndStroke(color, colors.borderColor);

                        const internalLoss = typeof subDayData.internallosse === 'number' ? subDayData.internallosse.toFixed(1) : subDayData.internallosse;
                        doc.font(dataFont)
                            .fontSize(dataFontSize)
                            .fillColor(subDayData.internallosse < 0 ? colors.redText : colors.blackText)
                            .text(internalLoss, startX + (colOffset * colWidths.subExport), yCentered, {
                                width: colWidths.subExport,
                                align: 'center'
                            });
                        colOffset++;

                        // Loss in% (red if negative, with special background)
                        const lossPercent = typeof subDayData.lossinparsantege === 'number' ? subDayData.lossinparsantege.toFixed(1) + '%' : subDayData.lossinparsantege;
                        const lossBgColor = subDayData.lossinparsantege < 0 ? colors.negativeLossBg : color;

                        doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataRow)
                            .lineWidth(0.5)
                            .fillAndStroke(lossBgColor, colors.borderColor);

                        doc.font(dataFont)
                            .fontSize(dataFontSize)
                            .fillColor(subDayData.lossinparsantege < 0 ? colors.redText : colors.blackText)
                            .text(lossPercent, startX + (colOffset * colWidths.subExport), yCentered, {
                                width: colWidths.subExport,
                                align: 'center'
                            });
                        colOffset++;
                    }

                    // Import (always last)
                    doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataRow)
                        .lineWidth(0.5)
                        .fillAndStroke(color, colors.borderColor);

                    doc.font(dataFont)
                        .fontSize(dataFontSize)
                        .fillColor(colors.blackText)
                        .text(typeof subDayData.import === 'number' ? subDayData.import.toFixed(1) : subDayData.import,
                            startX + (colOffset * colWidths.subExport), yCentered, {
                            width: colWidths.subExport,
                            align: 'center'
                        });
                }
            });

            // AC Line Loss Diff data
            const acLineDayData = dailyReport.aclinelossdiffrence.loggerdatas.find(d => d.date === dayData.date);
            if (acLineDayData) {
                const values = [
                    { value: acLineDayData.export, isLossPercent: false },
                    { value: acLineDayData.lossinparsantegeexport, isLossPercent: true },
                    { value: acLineDayData.import, isLossPercent: false },
                    { value: acLineDayData.lossinparsantegeimport, isLossPercent: true }
                ];

                values.forEach((item, i) => {
                    const formattedValue = typeof item.value === 'number' ?
                        (item.isLossPercent ? `${item.value.toFixed(1)}%` : item.value.toFixed(1)) :
                        item.value;

                    const bgColor = item.value < 0 ? colors.acLineBg : colors.whiteBg;
                    const textColor = item.value < 0 ? colors.redText : colors.blackText;
                    const yCentered = (currentY + (rowHeights.dataRow - dataFontSize) / 2) + 2;

                    doc.rect(acLineStartX + (i * colWidths.acLine), currentY, colWidths.acLine, rowHeights.dataRow)
                        .lineWidth(0.5)
                        .fillAndStroke(bgColor, colors.borderColor);

                    doc.font(dataFont)
                        .fontSize(dataFontSize)
                        .fillColor(textColor)
                        .text(formattedValue, acLineStartX + (i * colWidths.acLine), yCentered, {
                            width: colWidths.acLine,
                            align: 'center'
                        });
                });
            }

            currentY += rowHeights.dataRow;
        });

        // ====================
        // TOTALS ROW
        // ====================
        // Check if we need a new page for totals row
        if (currentY + rowHeights.totalsRow > doc.page.height - 20) {
            doc.addPage({
                size: 'A4',
                layout: 'landscape',
                margins: { top: 20, bottom: 20, left: 15, right: 15 }
            });
            currentY = 20;
        }
        const yCentered = (currentY + (rowHeights.dataRow - dataFontSize) / 2) + 2;
        // Label
        doc.rect(15, currentY, colWidths.date, rowHeights.totalsRow)
            .lineWidth(0.5)
            .fillAndStroke(colors.grayBg, colors.borderColor);

        doc.font(headerFont)
            .fontSize(totalFontSize)
            .fillColor(colors.blackText)
            .text('Total', 15, yCentered, {
                width: colWidths.date,
                align: 'center'
            });

        // Main client totals
        let mainTotalsOffset = 0;

        // Export total
        doc.rect(15 + colWidths.date + (mainTotalsOffset * colWidths.mainExport), currentY, colWidths.mainExport, rowHeights.totalsRow)
            .lineWidth(0.5)
            .fillAndStroke(colors.mainClientBg, colors.borderColor);
        doc.font(headerFont)
            .fontSize(totalFontSize)
            .fillColor(colors.blackText)
            .text(typeof dailyReport.mainClient.totalexport === 'number' ? dailyReport.mainClient.totalexport.toFixed(1) : dailyReport.mainClient.totalexport,
                15 + colWidths.date + (mainTotalsOffset * colWidths.mainExport), yCentered, {
                width: colWidths.mainExport,
                align: 'center'
            });
        mainTotalsOffset++;

        // Avg Generation DC total (if enabled)
        if (showAvgColumn) {
            doc.rect(15 + colWidths.date + (mainTotalsOffset * colWidths.mainExport), currentY, colWidths.mainExport, rowHeights.totalsRow)
                .lineWidth(0.5)
                .fillAndStroke(colors.mainClientBg, colors.borderColor);
            doc.font(headerFont)
                .fontSize(totalFontSize)
                .fillColor(colors.blackText)
                .text(typeof dailyReport.mainClient.totalAvgGeneration === 'number' ? dailyReport.mainClient.totalAvgGeneration.toFixed(1) : (dailyReport.mainClient.totalAvgGeneration || '0.0'),
                    15 + colWidths.date + (mainTotalsOffset * colWidths.mainExport), yCentered, {
                    width: colWidths.mainExport,
                    align: 'center'
                });
            mainTotalsOffset++;
        }

        // Avg Generation AC total (if enabled)
        if (showAvgAcColumn) {
            doc.rect(15 + colWidths.date + (mainTotalsOffset * colWidths.mainExport), currentY, colWidths.mainExport, rowHeights.totalsRow)
                .lineWidth(0.5)
                .fillAndStroke(colors.mainClientBg, colors.borderColor);
            doc.font(headerFont)
                .fontSize(totalFontSize)
                .fillColor(colors.blackText)
                .text(typeof dailyReport.mainClient.totalAvgGenerationAc === 'number' ? dailyReport.mainClient.totalAvgGenerationAc.toFixed(1) : (dailyReport.mainClient.totalAvgGenerationAc || '0.0'),
                    15 + colWidths.date + (mainTotalsOffset * colWidths.mainExport), yCentered, {
                    width: colWidths.mainExport,
                    align: 'center'
                });
            mainTotalsOffset++;
        }

        // Import total
        doc.rect(15 + colWidths.date + (mainTotalsOffset * colWidths.mainExport), currentY, colWidths.mainImport, rowHeights.totalsRow)
            .lineWidth(0.5)
            .fillAndStroke(colors.mainClientBg, colors.borderColor);
        doc.font(headerFont)
            .fontSize(totalFontSize)
            .fillColor(colors.blackText)
            .text(typeof dailyReport.mainClient.totalimport === 'number' ? dailyReport.mainClient.totalimport.toFixed(1) : dailyReport.mainClient.totalimport,
                15 + colWidths.date + (mainTotalsOffset * colWidths.mainExport), yCentered, {
                width: colWidths.mainImport,
                align: 'center'
            });

        // Sub-client totals
        dailyReport.subClient.forEach((subClient, index) => {
            const startX = 15 + colWidths.date + mainClientNameWidth + (index * colsPerSubClient * colWidths.subExport);
            const color = index % 2 === 0 ? colors.subClient1Bg : colors.subClient2Bg;
            let colOffset = 0;

            // Export total (always shown)
            doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.totalsRow)
                .lineWidth(0.5)
                .fillAndStroke(color, colors.borderColor)

            doc.font(headerFont)
                .fontSize(totalFontSize)
                .fillColor(colors.blackText)
                .text(typeof subClient.totalexport === 'number' ? subClient.totalexport.toFixed(1) : subClient.totalexport,
                    startX + (colOffset * colWidths.subExport), yCentered, {
                    width: colWidths.subExport,
                    align: 'center'
                });
            colOffset++;

            // Avg Generation DC total (if enabled)
            if (showAvgColumn) {
                doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.totalsRow)
                    .lineWidth(0.5)
                    .fillAndStroke(color, colors.borderColor);

                doc.font(headerFont)
                    .fontSize(totalFontSize)
                    .fillColor(colors.blackText)
                    .text(typeof subClient.totalAvgGeneration === 'number' ? subClient.totalAvgGeneration.toFixed(1) : (subClient.totalAvgGeneration || 0),
                        startX + (colOffset * colWidths.subExport), yCentered, {
                        width: colWidths.subExport,
                        align: 'center'
                    });
                colOffset++;
            }

            // Avg Generation AC total (if enabled)
            if (showAvgAcColumn) {
                doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.totalsRow)
                    .lineWidth(0.5)
                    .fillAndStroke(color, colors.borderColor);

                doc.font(headerFont)
                    .fontSize(totalFontSize)
                    .fillColor(colors.blackText)
                    .text(typeof subClient.totalAvgGenerationAc === 'number' ? subClient.totalAvgGenerationAc.toFixed(1) : (subClient.totalAvgGenerationAc || 0),
                        startX + (colOffset * colWidths.subExport), yCentered, {
                        width: colWidths.subExport,
                        align: 'center'
                    });
                colOffset++;
            }

            if (showLoggerColumns) {
                // Logger Data total
                doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.totalsRow)
                    .lineWidth(0.5)
                    .fillAndStroke(color, colors.borderColor);

                doc.font(headerFont)
                    .fontSize(totalFontSize)
                    .fillColor(colors.blackText)
                    .text(typeof subClient.totalloggerdata === 'number' ? subClient.totalloggerdata.toFixed(1) : subClient.totalloggerdata,
                        startX + (colOffset * colWidths.subExport), yCentered, {
                        width: colWidths.subExport,
                        align: 'center'
                    });
                colOffset++;

                // Internal Loss total (red if negative)
                doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.totalsRow)
                    .lineWidth(0.5)
                    .fillAndStroke(color, colors.borderColor);

                const internalLoss = typeof subClient.totalinternallosse === 'number' ? subClient.totalinternallosse.toFixed(1) : subClient.totalinternallosse;
                doc.font(headerFont)
                    .fontSize(totalFontSize)
                    .fillColor(subClient.totalinternallosse < 0 ? colors.redText : colors.blackText)
                    .text(internalLoss, startX + (colOffset * colWidths.subExport), yCentered, {
                        width: colWidths.subExport,
                        align: 'center'
                    });
                colOffset++;

                // Loss in% total (red if negative, with special background)
                const lossPercent = typeof subClient.totallossinparsantege === 'number' ? subClient.totallossinparsantege.toFixed(1) + '%' : subClient.totallossinparsantege;
                const lossBgColor = subClient.totallossinparsantege < 0 ? colors.negativeLossBg : color;

                doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.totalsRow)
                    .lineWidth(0.5)
                    .fillAndStroke(lossBgColor, colors.borderColor);

                doc.font(headerFont)
                    .fontSize(totalFontSize)
                    .fillColor(subClient.totallossinparsantege < 0 ? colors.redText : colors.blackText)
                    .text(lossPercent, startX + (colOffset * colWidths.subExport), yCentered, {
                        width: colWidths.subExport,
                        align: 'center'
                    });
                colOffset++;
            }

            // Import total (always last)
            doc.rect(startX + (colOffset * colWidths.subExport), currentY, colWidths.subExport, rowHeights.totalsRow)
                .lineWidth(0.5)
                .fillAndStroke(color, colors.borderColor);

            doc.font(headerFont)
                .fontSize(totalFontSize)
                .fillColor(colors.blackText)
                .text(typeof subClient.totalimport === 'number' ? subClient.totalimport.toFixed(1) : subClient.totalimport,
                    startX + (colOffset * colWidths.subExport), yCentered, {
                    width: colWidths.subExport,
                    align: 'center'
                });
        });

        // AC Line Loss Diff totals
        const acLineTotals = [
            { value: dailyReport.aclinelossdiffrence.totalexport, isLossPercent: false },
            { value: dailyReport.aclinelossdiffrence.totallossinparsantegeexport, isLossPercent: true },
            { value: dailyReport.aclinelossdiffrence.totalimport, isLossPercent: false },
            { value: dailyReport.aclinelossdiffrence.totallossinparsantegeimport, isLossPercent: true }
        ];

        acLineTotals.forEach((item, i) => {
            const formattedValue = typeof item.value === 'number' ?
                (item.isLossPercent ? `${item.value.toFixed(1)}%` : item.value.toFixed(1)) :
                item.value;

            const bgColor = item.value < 0 ? colors.acLineBg : colors.whiteBg;
            const textColor = item.value < 0 ? colors.redText : colors.blackText;

            doc.rect(acLineStartX + (i * colWidths.acLine), currentY, colWidths.acLine, rowHeights.totalsRow)
                .lineWidth(0.5)
                .fillAndStroke(bgColor, colors.borderColor);

            doc.font(headerFont)
                .fontSize(totalFontSize)
                .fillColor(textColor)
                .text(formattedValue, acLineStartX + (i * colWidths.acLine), yCentered, {
                    width: colWidths.acLine,
                    align: 'center'
                });
        });

        // ====================
        // NOTE TEXT - Capacity Basis
        // ====================
        // Only add note if at least one avg column is shown
        if (showAvgColumn || showAvgAcColumn) {
            // Determine which note to display
            let noteText = '';
            if (showAvgColumn && dailyReport.capacityBasisDC) {
                noteText = `Note: The Average Generation has been calculated on the basis of the plant ${dailyReport.capacityBasisDC} capacity.`;
            } else if (showAvgAcColumn && dailyReport.capacityBasisAC) {
                noteText = `Note: The Average Generation has been calculated on the basis of the plant ${dailyReport.capacityBasisAC} capacity.`;
            }

            if (noteText) {
                currentY += rowHeights.totalsRow + 10; // Add some spacing after totals

                // Check if we need a new page
                if (currentY + 20 > doc.page.height - 20) {
                    doc.addPage({
                        size: 'A4',
                        layout: 'landscape',
                        margins: { top: 20, bottom: 20, left: 15, right: 15 }
                    });
                    currentY = 20;
                }

                doc.font('Times-BoldItalic')
                    .fontSize(10)
                    .fillColor(colors.blackText)
                    .text(noteText, startX, currentY, {
                        width: totalWidth,
                        align: 'center'
                    });
            }
        }

        // Finalize
        doc.end();
    } catch (error) {
        logger.error(`Error generating Daily Report PDF: ${error.message}`);
        res.status(500).json({
            message: "Error generating Daily Report PDF",
            error: error.message,
        });
    }
};


// Helper function to get month name
function getMonthName(month) {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return months[month - 1] || '';
}