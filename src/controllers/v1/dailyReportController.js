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
        const { mainClientId, month, year } = req.body;

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
        let mainClientTotalExport = 0, mainClientTotalImport = 0;
        let currentDate = '', dailyExport = 0, dailyImport = 0;

        for (let data of meterData) {
            for (let entry of data.dataEntries) {
                const entryDate = entry.parameters.Date;
                if (entryDate !== currentDate) {
                    if (currentDate !== '') {
                        loggerData.push({ date: currentDate, export: dailyExport, import: dailyImport });
                        mainClientTotalExport += dailyExport;
                        mainClientTotalImport += dailyImport;
                    }
                    currentDate = entryDate;
                    dailyExport = 0;
                    dailyImport = 0;
                }

                const { export: dayExport, import: dayImport } = calculateExportImport([entry], mainClientData.mf ,mainClientData.pn);
                dailyExport += dayExport;
                dailyImport += dayImport;
            }
        }

        if (currentDate !== '') {
            loggerData.push({ date: currentDate, export: dailyExport, import: dailyImport });
            mainClientTotalExport += dailyExport;
            mainClientTotalImport += dailyImport;
        }

        // Initialize daily report
        const dailyReport = new DailyReport({
            mainClientId,
            month,
            year,
            mainClient: {
                meterNumber: meterData[0].meterNumber,
                meterType: meterData[0].meterType,
                mainClientDetail: mainClientData.toObject(),
                totalexport: mainClientTotalExport,
                totalimport: mainClientTotalImport,
                loggerdatas: loggerData
            },
            subClient: [],
            aclinelossdiffrence: {}
        });

        // Step 5: Process Sub Clients with logger data validation
        for (let subClient of subClients) {
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
            let subClientCurrentDate = '', dailyExport = 0, dailyImport = 0;

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

                            subClientLoggerData.push({
                                date: subClientCurrentDate,
                                export: parseFloat(dailyExport.toFixed(2)),
                                import: parseFloat(dailyImport.toFixed(2)),
                                loggerdata: loggerDataValue,
                                internallosse: parseFloat(internallosse.toFixed(2)),
                                lossinparsantege: parseFloat(lossinparsantege.toFixed(2))
                            });

                            subClientTotalExport += dailyExport;
                            subClientTotalImport += dailyImport;
                            subClientTotalLoggerData += loggerDataValue;
                            subClientTotalInternalLosses += internallosse;
                        }

                        subClientCurrentDate = entryDate;
                        dailyExport = 0;
                        dailyImport = 0;
                    }

                    const { export: dayExport, import: dayImport } = calculateExportImport([entry], subClient.mf,subClient.pn);
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

                subClientLoggerData.push({
                    date: subClientCurrentDate,
                    export: parseFloat(dailyExport.toFixed(2)),
                    import: parseFloat(dailyImport.toFixed(2)),
                    loggerdata: loggerDataValue,
                    internallosse: parseFloat(internallosse.toFixed(2)),
                    lossinparsantege: parseFloat(lossinparsantege.toFixed(2))
                });

                subClientTotalExport += dailyExport;
                subClientTotalImport += dailyImport;
                subClientTotalLoggerData += loggerDataValue;
                subClientTotalInternalLosses += internallosse;
            }

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
        const subClientCols = subClientCount * 5;
        const totalCols = 1 + 2 + subClientCols + 8; // Date(1) + Main(2) + Subs + AC Line(4)
        // Determine how many columns to cover in row 3 (minimum K, maximum same as row 2)
        const row3CoverCols = Math.max(11, totalCols); // At least up to column K (11), but more if needed

        // ====================
        // SET COLUMN WIDTHS
        // ====================

        // Main client columns
        worksheet.getColumn('A').width = 12;  // A cell width is 12
        worksheet.getColumn('B').width = 10;   // B cell width is 8
        worksheet.getColumn('C').width = 10;   // C cell width is 6

        // Sub-client columns (repeated for each sub-client)
        dailyReport.subClient.forEach((subClient, index) => {
            const baseCol = 4 + (index * 5); // Starting column for each sub-client

            worksheet.getColumn(baseCol).width = 9;     // D (Export)
            worksheet.getColumn(baseCol + 1).width = 9; // E (Data)
            worksheet.getColumn(baseCol + 2).width = 9; // F (Loss)
            worksheet.getColumn(baseCol + 3).width = 9; // G (%)
            worksheet.getColumn(baseCol + 4).width = 8; // H (Import)
        });

        // AC LINE LOSS DIFF columns
        const acLineStartCol = 4 + (dailyReport.subClient.length * 5);
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
        const subClientColss = effectiveSubClients * 5;
        const totalColss = 1 + 2 + subClientColss + 4; // Date(1) + Main(2) + Subs + AC Line(4)

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
        const monthEndCol = 1 + 2 + 5; // NAME (1) + Total GETCO (2) + 1 subclient (5)
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

        // Main client cells (B4:C4)
        worksheet.mergeCells('B4:C4');
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
        // Add thin inside border
        worksheet.getCell('B4').border.right = { style: 'thin' };

        // Sub-client cells - alternating colors
        const subClientColors = ['FFFF00', '92D050']; // Yellow and green
        dailyReport.subClient.forEach((subClient, index) => {
            const startCol = 4 + (index * 5);
            const color = subClientColors[index % 2]; // Alternate between colors

            worksheet.mergeCells(4, startCol, 4, startCol + 4);
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
            for (let i = 1; i <= 4; i++) {
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

        // Merge B5:C5 and add main client meter number
        worksheet.mergeCells('B5:C5');
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
        // Add thin inside border
        worksheet.getCell('B5').border.right = { style: 'thin' };

        // Merge D5:H5 for each sub-client
        dailyReport.subClient.forEach((subClient, index) => {
            const startCol = 4 + (index * 5);
            worksheet.mergeCells(5, startCol, 5, startCol + 4);
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
            for (let i = 1; i <= 4; i++) {
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

        // Cell C6
        const importHeaderCell = worksheet.getCell('C6');
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
            const startCol = 4 + (index * 5);
            const headers = [
                { text: 'Export', color: '000000' },
                { text: 'Logger Data', color: '000000' },
                { text: 'Internal Loss', color: 'FF0000' },
                { text: 'Loss in%', color: 'FF0000' },
                { text: 'Import', color: '000000' }
            ];

            headers.forEach((header, i) => {
                const cell = worksheet.getCell(6, startCol + i);
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

            // Main client data (B:C)
            [2, 3].forEach(col => {
                const cell = row.getCell(col);
                const value = col === 2 ? dayData.export : dayData.import;
                cell.value = typeof value === 'number' ? parseFloat(value.toFixed(1)) : value;
                Object.assign(cell, dataStyle);
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: bgColors.mainClient }
                };
            });

            // Sub-client data
            dailyReport.subClient.forEach((subClient, index) => {
                const subDayData = subClient.loggerdatas.find(d => d.date === dayData.date);
                const startCol = 4 + (index * 5);
                const color = index % 2 === 0 ? bgColors.subClient1 : bgColors.subClient2;

                if (subDayData) {
                    // Format all cells for this sub-client
                    const exportCell = row.getCell(startCol);
                    exportCell.value = typeof subDayData.export === 'number' ? parseFloat(subDayData.export.toFixed(1)) : subDayData.export;
                    Object.assign(exportCell, dataStyle);
                    exportCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };

                    const loggerDataCell = row.getCell(startCol + 1);
                    loggerDataCell.value = typeof subDayData.loggerdata === 'number' ? parseFloat(subDayData.loggerdata.toFixed(1)) : subDayData.loggerdata;
                    Object.assign(loggerDataCell, dataStyle);
                    loggerDataCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };

                    // Internal Loss (red font only when negative)
                    const internalLossCell = row.getCell(startCol + 2);
                    internalLossCell.value = typeof subDayData.internallosse === 'number' ? parseFloat(subDayData.internallosse.toFixed(1)) : subDayData.internallosse;
                    Object.assign(internalLossCell, dataStyle);
                    internalLossCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };

                    // Set font color based on value (red if negative, black otherwise)
                    const isNegativeInternalLoss = subDayData.internallosse < 0;
                    internalLossCell.font = {
                        ...dataStyle.font,
                        color: { argb: isNegativeInternalLoss ? 'FFFF0000' : 'FF000000' }
                    };

                    const lossPercentCell = row.getCell(startCol + 3);
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

                    const importCell = row.getCell(startCol + 4);
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
        [2, 3].forEach(col => {
            const cell = totalsRow.getCell(col);
            const value = col === 2 ? dailyReport.mainClient.totalexport : dailyReport.mainClient.totalimport;
            cell.value = typeof value === 'number' ? parseFloat(value.toFixed(1)) : value;
            Object.assign(cell, totalsStyle);
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: bgColors.mainClient } // Orange background
            };
        });

        // Sub-client totals
        dailyReport.subClient.forEach((subClient, index) => {
            const startCol = 4 + (index * 5);
            const color = index % 2 === 0 ? bgColors.subClient1 : bgColors.subClient2;

            // Export
            const exportCell = totalsRow.getCell(startCol);
            exportCell.value = typeof subClient.totalexport === 'number' ? parseFloat(subClient.totalexport.toFixed(1)) : subClient.totalexport;
            Object.assign(exportCell, totalsStyle);
            exportCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };

            // Logger Data
            const loggerDataCell = totalsRow.getCell(startCol + 1);
            loggerDataCell.value = typeof subClient.totalloggerdata === 'number' ? parseFloat(subClient.totalloggerdata.toFixed(1)) : subClient.totalloggerdata;
            Object.assign(loggerDataCell, totalsStyle);
            loggerDataCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };

            // Internal Loss (red font only when negative)
            const internalLossCell = totalsRow.getCell(startCol + 2);
            internalLossCell.value = typeof subClient.totalinternallosse === 'number' ? parseFloat(subClient.totalinternallosse.toFixed(1)) : subClient.totalinternallosse;
            Object.assign(internalLossCell, totalsStyle);
            internalLossCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };

            // Set font color based on value (red if negative, black otherwise)
            const isNegativeInternalLoss = subClient.totalinternallosse < 0;
            internalLossCell.font = {
                ...totalsStyle.font,
                color: { argb: isNegativeInternalLoss ? 'FFFF0000' : 'FF000000' }
            };

            // Loss in% total (4th cell)
            const lossPercentCell = totalsRow.getCell(startCol + 3);
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

            // Import
            const importCell = totalsRow.getCell(startCol + 4);
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


exports.downloadDailyReportPDF = async (req, res) => {
    try {
        const { dailyReportId } = req.params;
        const dailyReport = await DailyReport.findById(dailyReportId)
            .populate('mainClientId')
            .populate('subClient.subClientId');

        if (!dailyReport) {
            return res.status(404).json({ message: 'Daily Report not found' });
        }

        // Get month name and format for filename
        const filemonthName = getMonthName(dailyReport.month).toUpperCase();
        const year = dailyReport.year;
        const mainClientName = dailyReport.mainClient.mainClientDetail.name
            .replace(/\s+/g, '_')       // Replace spaces with underscores
            .replace(/\//g, '_')        // Replace forward slashes with underscores
            .toUpperCase();

        // Format filename
        const fileName = `Daily Generation Report - ${mainClientName} Month of ${filemonthName}-${year}.pdf`;

        // Set response headers
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader(
            'Content-Disposition',
            `attachment; filename="${encodeURIComponent(fileName)}"`
        );
        res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

        // Create PDF document in landscape mode
        const doc = new PDFDocument({
            size: 'A4',
            layout: 'landscape',
            margins: { top: 20, bottom: 20, left: 15, right: 15 }
        });

        // Pipe the PDF to the response
        doc.pipe(res);

        // ====================
        // PDF STYLING
        // ====================
        const titleFont = 'Helvetica-Bold';
        const headerFont = 'Helvetica-Bold';
        const dataFont = 'Helvetica';
        const titleFontSize = 16;
        const headerFontSize = 10;
        const dataFontSize = 9;

        // Color scheme matching Excel
        const colors = {
            titleBg: '#bdd7ee',      // Light blue (title background)
            monthBg: '#bfbfbf',       // Gray (month cell background)
            mainClientBg: '#ffc000',  // Orange (main client)
            subClient1Bg: '#FFFF00',  // Yellow (sub-client 1)
            subClient2Bg: '#92D050',  // Green (sub-client 2)
            acLineBg: '#fac6c6',      // Light red (AC line)
            negativeLossBg: '#ffc7ce',// Light red (negative loss)
            redText: '#FF0000',       // Red text
            blackText: '#000000',     // Black text
            whiteBg: '#FFFFFF',        // White background
            borderColor: '#000000',    // Black borders
            grayBg: '#D9D9D9'          // Light gray
        };

        // Column widths (approximate proportions from Excel)
        const colWidths = {
            date: 50,                  // Date column
            mainExport: 40,            // Main client export
            mainImport: 40,            // Main client import
            subExport: 35,             // Sub-client export
            subData: 35,               // Sub-client logger data
            subLoss: 35,               // Sub-client internal loss
            subPercent: 35,            // Sub-client loss %
            subImport: 35,             // Sub-client import
            acLine: 35                // AC line columns
        };

        // Row heights
        const rowHeights = {
            title: 30,
            monthDate: 30,
            clientHeader: 40,
            meterHeader: 40,
            dataHeader: 25,
            dataRow: 24,
            totalsRow: 25
        };

        // Starting position
        let currentY = 20;

        // ====================
        // TITLE ROW
        // ====================
        const titleText = `${dailyReport.mainClient.mainClientDetail.name} - ${(dailyReport.mainClient.mainClientDetail.acCapacityKw / 1000).toFixed(2)} MW AC Generation Details`;
        
        // Calculate total width based on number of sub-clients
        const subClientCount = Math.max(dailyReport.subClient.length, 2); // Minimum 2 sub-clients
        const subClientCols = subClientCount * 5;
        const totalWidth = colWidths.date + colWidths.mainExport + colWidths.mainImport + 
                          (subClientCols * colWidths.subExport) + (4 * colWidths.acLine);

        // Title background
        doc.rect(15, currentY, totalWidth, rowHeights.title)
            .fill(colors.titleBg)
            .stroke(colors.borderColor);

        // Title text
        doc.font(titleFont)
            .fontSize(titleFontSize)
            .fillColor(colors.blackText)
            .text(titleText, 15, currentY + 7, {
                width: totalWidth,
                align: 'center'
            });

        currentY += rowHeights.title;

        // ====================
        // MONTH AND DATE RANGE ROW
        // ====================
        const monthName = getMonthName(dailyReport.month);
        const yearShort = dailyReport.year.toString().slice(-2);
        const lastDay = new Date(dailyReport.year, dailyReport.month, 0).getDate();
        const dateRange = `01-${String(dailyReport.month).padStart(2, '0')}-${dailyReport.year} to ${lastDay}-${String(dailyReport.month).padStart(2, '0')}-${dailyReport.year}`;

        // Calculate widths for month and date range sections
        const monthWidth = colWidths.date + colWidths.mainExport + colWidths.mainImport + 
                          (5 * colWidths.subExport); // Width for one sub-client section
        const dateRangeWidth = totalWidth - monthWidth;

        // Month cell
        doc.rect(15, currentY, monthWidth, rowHeights.monthDate)
            .fill(colors.monthBg)
            .stroke(colors.borderColor);

        doc.font(headerFont)
            .fontSize(14)
            .fillColor(colors.redText)
            .text(`Month: ${monthName}-${yearShort}`, 15, currentY + 7, {
                width: monthWidth,
                align: 'center'
            });

        // Date range cell
        doc.rect(15 + monthWidth, currentY, dateRangeWidth, rowHeights.monthDate)
            .fill(colors.monthBg)
            .stroke(colors.borderColor);

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
            .fill(colors.grayBg)
            .stroke(colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('NAME=>', 15, currentY + 15, {
                width: colWidths.date,
                align: 'center'
            });

        // TOTAL-GETCO SS cell (merged)
        doc.rect(15 + colWidths.date, currentY, colWidths.mainExport + colWidths.mainImport, rowHeights.clientHeader)
            .fill(colors.mainClientBg)
            .stroke(colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('TOTAL-GETCO\nSS', 15 + colWidths.date, currentY + 10, {
                width: colWidths.mainExport + colWidths.mainImport,
                align: 'center',
                lineGap: 3
            });

        // Sub-client cells
        dailyReport.subClient.forEach((subClient, index) => {
            const startX = 15 + colWidths.date + colWidths.mainExport + colWidths.mainImport + (index * 5 * colWidths.subExport);
            const color = index % 2 === 0 ? colors.subClient1Bg : colors.subClient2Bg;

            doc.rect(startX, currentY, 5 * colWidths.subExport, rowHeights.clientHeader)
                .fill(color)
                .stroke(colors.borderColor);

            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.blackText)
                .text(subClient.name, startX, currentY + 15, {
                    width: 5 * colWidths.subExport,
                    align: 'center'
                });
        });

        // AC LINE LOSS DIFF header (merged with next row)
        const acLineStartX = 15 + colWidths.date + colWidths.mainExport + colWidths.mainImport + 
                           (dailyReport.subClient.length * 5 * colWidths.subExport);
        
        doc.rect(acLineStartX, currentY, 4 * colWidths.acLine, rowHeights.clientHeader * 2)
            .fill(colors.acLineBg)
            .stroke(colors.borderColor);

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
            .fill(colors.grayBg)
            .stroke(colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('METER NO.=>', 15, currentY + 15, {
                width: colWidths.date,
                align: 'center'
            });

        // Main client meter number (merged)
        doc.rect(15 + colWidths.date, currentY, colWidths.mainExport + colWidths.mainImport, rowHeights.meterHeader)
            .fill(colors.mainClientBg)
            .stroke(colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text(dailyReport.mainClient.meterNumber, 15 + colWidths.date, currentY + 15, {
                width: colWidths.mainExport + colWidths.mainImport,
                align: 'center'
            });

        // Sub-client meter numbers
        dailyReport.subClient.forEach((subClient, index) => {
            const startX = 15 + colWidths.date + colWidths.mainExport + colWidths.mainImport + (index * 5 * colWidths.subExport);
            const color = index % 2 === 0 ? colors.subClient1Bg : colors.subClient2Bg;

            doc.rect(startX, currentY, 5 * colWidths.subExport, rowHeights.meterHeader)
                .fill(color)
                .stroke(colors.borderColor);

            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.blackText)
                .text(subClient.meterNumber, startX, currentY + 15, {
                    width: 5 * colWidths.subExport,
                    align: 'center'
                });
        });

        currentY += rowHeights.meterHeader;

        // ====================
        // HEADERS ROW
        // ====================
        // DATE header
        doc.rect(15, currentY, colWidths.date, rowHeights.dataHeader)
            .fill(colors.grayBg)
            .stroke(colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('DATE', 15, currentY + 7, {
                width: colWidths.date,
                align: 'center'
            });

        // Main client headers
        doc.rect(15 + colWidths.date, currentY, colWidths.mainExport, rowHeights.dataHeader)
            .fill(colors.mainClientBg)
            .stroke(colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('Export', 15 + colWidths.date, currentY + 7, {
                width: colWidths.mainExport,
                align: 'center'
            });

        doc.rect(15 + colWidths.date + colWidths.mainExport, currentY, colWidths.mainImport, rowHeights.dataHeader)
            .fill(colors.mainClientBg)
            .stroke(colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('Import', 15 + colWidths.date + colWidths.mainExport, currentY + 7, {
                width: colWidths.mainImport,
                align: 'center'
            });

        // Sub-client headers
        dailyReport.subClient.forEach((subClient, index) => {
            const startX = 15 + colWidths.date + colWidths.mainExport + colWidths.mainImport + (index * 5 * colWidths.subExport);
            const color = index % 2 === 0 ? colors.subClient1Bg : colors.subClient2Bg;
            const headers = ['Export', 'Logger Data', 'Internal Loss', 'Loss in%', 'Import'];
            const textColors = ['000000', '000000', 'FF0000', 'FF0000', '000000'];

            headers.forEach((header, i) => {
                doc.rect(startX + (i * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataHeader)
                    .fill(color)
                    .stroke(colors.borderColor);

                doc.font(headerFont)
                    .fontSize(headerFontSize)
                    .fillColor(textColors[i] === 'FF0000' ? colors.redText : colors.blackText)
                    .text(header, startX + (i * colWidths.subExport), currentY + 7, {
                        width: colWidths.subExport,
                        align: 'center'
                    });
            });
        });

        // AC LINE LOSS DIFF headers
        const acLineHeaders = ['Export', 'Loss in%', 'Import', 'Loss in%'];
        acLineHeaders.forEach((header, i) => {
            doc.rect(acLineStartX + (i * colWidths.acLine), currentY, colWidths.acLine, rowHeights.dataHeader)
                .fill(colors.acLineBg)
                .stroke(colors.borderColor);

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

            // Date cell
            doc.rect(15, currentY, colWidths.date, rowHeights.dataRow)
                .fill(colors.grayBg)
                .stroke(colors.borderColor);

            doc.font(dataFont)
                .fontSize(dataFontSize)
                .fillColor(colors.blackText)
                .text(dayData.date, 15, currentY + 7, {
                    width: colWidths.date,
                    align: 'center'
                });

            // Main client data
            doc.rect(15 + colWidths.date, currentY, colWidths.mainExport, rowHeights.dataRow)
                .fill(colors.mainClientBg)
                .stroke(colors.borderColor);

            doc.font(dataFont)
                .fontSize(dataFontSize)
                .fillColor(colors.blackText)
                .text(typeof dayData.export === 'number' ? dayData.export.toFixed(1) : dayData.export, 
                      15 + colWidths.date, currentY + 7, {
                    width: colWidths.mainExport,
                    align: 'center'
                });

            doc.rect(15 + colWidths.date + colWidths.mainExport, currentY, colWidths.mainImport, rowHeights.dataRow)
                .fill(colors.mainClientBg)
                .stroke(colors.borderColor);

            doc.font(dataFont)
                .fontSize(dataFontSize)
                .fillColor(colors.blackText)
                .text(typeof dayData.import === 'number' ? dayData.import.toFixed(1) : dayData.import, 
                      15 + colWidths.date + colWidths.mainExport, currentY + 7, {
                    width: colWidths.mainImport,
                    align: 'center'
                });

            // Sub-client data
            dailyReport.subClient.forEach((subClient, index) => {
                const subDayData = subClient.loggerdatas.find(d => d.date === dayData.date);
                const startX = 15 + colWidths.date + colWidths.mainExport + colWidths.mainImport + (index * 5 * colWidths.subExport);
                const color = index % 2 === 0 ? colors.subClient1Bg : colors.subClient2Bg;

                if (subDayData) {
                    // Export
                    doc.rect(startX, currentY, colWidths.subExport, rowHeights.dataRow)
                        .fill(color)
                        .stroke(colors.borderColor);

                    doc.font(dataFont)
                        .fontSize(dataFontSize)
                        .fillColor(colors.blackText)
                        .text(typeof subDayData.export === 'number' ? subDayData.export.toFixed(1) : subDayData.export, 
                              startX, currentY + 7, {
                            width: colWidths.subExport,
                            align: 'center'
                        });

                    // Logger Data
                    doc.rect(startX + colWidths.subExport, currentY, colWidths.subExport, rowHeights.dataRow)
                        .fill(color)
                        .stroke(colors.borderColor);

                    doc.font(dataFont)
                        .fontSize(dataFontSize)
                        .fillColor(colors.blackText)
                        .text(typeof subDayData.loggerdata === 'number' ? subDayData.loggerdata.toFixed(1) : subDayData.loggerdata, 
                              startX + colWidths.subExport, currentY + 7, {
                            width: colWidths.subExport,
                            align: 'center'
                        });

                    // Internal Loss (red if negative)
                    doc.rect(startX + (2 * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataRow)
                        .fill(color)
                        .stroke(colors.borderColor);

                    const internalLoss = typeof subDayData.internallosse === 'number' ? subDayData.internallosse.toFixed(1) : subDayData.internallosse;
                    doc.font(dataFont)
                        .fontSize(dataFontSize)
                        .fillColor(subDayData.internallosse < 0 ? colors.redText : colors.blackText)
                        .text(internalLoss, startX + (2 * colWidths.subExport), currentY + 7, {
                            width: colWidths.subExport,
                            align: 'center'
                        });

                    // Loss in% (red if negative, with special background)
                    const lossPercent = typeof subDayData.lossinparsantege === 'number' ? subDayData.lossinparsantege.toFixed(1) + '%' : subDayData.lossinparsantege;
                    const lossBgColor = subDayData.lossinparsantege < 0 ? colors.negativeLossBg : color;
                    
                    doc.rect(startX + (3 * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataRow)
                        .fill(lossBgColor)
                        .stroke(colors.borderColor);

                    doc.font(dataFont)
                        .fontSize(dataFontSize)
                        .fillColor(subDayData.lossinparsantege < 0 ? colors.redText : colors.blackText)
                        .text(lossPercent, startX + (3 * colWidths.subExport), currentY + 7, {
                            width: colWidths.subExport,
                            align: 'center'
                        });

                    // Import
                    doc.rect(startX + (4 * colWidths.subExport), currentY, colWidths.subExport, rowHeights.dataRow)
                        .fill(color)
                        .stroke(colors.borderColor);

                    doc.font(dataFont)
                        .fontSize(dataFontSize)
                        .fillColor(colors.blackText)
                        .text(typeof subDayData.import === 'number' ? subDayData.import.toFixed(1) : subDayData.import, 
                              startX + (4 * colWidths.subExport), currentY + 7, {
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

                    doc.rect(acLineStartX + (i * colWidths.acLine), currentY, colWidths.acLine, rowHeights.dataRow)
                        .fill(bgColor)
                        .stroke(colors.borderColor);

                    doc.font(dataFont)
                        .fontSize(dataFontSize)
                        .fillColor(textColor)
                        .text(formattedValue, acLineStartX + (i * colWidths.acLine), currentY + 7, {
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

        // Label
        doc.rect(15, currentY, colWidths.date, rowHeights.totalsRow)
            .fill(colors.grayBg)
            .stroke(colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text('Total', 15, currentY + 7, {
                width: colWidths.date,
                align: 'center'
            });

        // Main client totals
        doc.rect(15 + colWidths.date, currentY, colWidths.mainExport, rowHeights.totalsRow)
            .fill(colors.mainClientBg)
            .stroke(colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text(typeof dailyReport.mainClient.totalexport === 'number' ? dailyReport.mainClient.totalexport.toFixed(1) : dailyReport.mainClient.totalexport, 
                  15 + colWidths.date, currentY + 7, {
                width: colWidths.mainExport,
                align: 'center'
            });

        doc.rect(15 + colWidths.date + colWidths.mainExport, currentY, colWidths.mainImport, rowHeights.totalsRow)
            .fill(colors.mainClientBg)
            .stroke(colors.borderColor);

        doc.font(headerFont)
            .fontSize(headerFontSize)
            .fillColor(colors.blackText)
            .text(typeof dailyReport.mainClient.totalimport === 'number' ? dailyReport.mainClient.totalimport.toFixed(1) : dailyReport.mainClient.totalimport, 
                  15 + colWidths.date + colWidths.mainExport, currentY + 7, {
                width: colWidths.mainImport,
                align: 'center'
            });

        // Sub-client totals
        dailyReport.subClient.forEach((subClient, index) => {
            const startX = 15 + colWidths.date + colWidths.mainExport + colWidths.mainImport + (index * 5 * colWidths.subExport);
            const color = index % 2 === 0 ? colors.subClient1Bg : colors.subClient2Bg;

            // Export total
            doc.rect(startX, currentY, colWidths.subExport, rowHeights.totalsRow)
                .fill(color)
                .stroke(colors.borderColor);

            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.blackText)
                .text(typeof subClient.totalexport === 'number' ? subClient.totalexport.toFixed(1) : subClient.totalexport, 
                      startX, currentY + 7, {
                    width: colWidths.subExport,
                    align: 'center'
                });

            // Logger Data total
            doc.rect(startX + colWidths.subExport, currentY, colWidths.subExport, rowHeights.totalsRow)
                .fill(color)
                .stroke(colors.borderColor);

            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.blackText)
                .text(typeof subClient.totalloggerdata === 'number' ? subClient.totalloggerdata.toFixed(1) : subClient.totalloggerdata, 
                      startX + colWidths.subExport, currentY + 7, {
                    width: colWidths.subExport,
                    align: 'center'
                });

            // Internal Loss total (red if negative)
            doc.rect(startX + (2 * colWidths.subExport), currentY, colWidths.subExport, rowHeights.totalsRow)
                .fill(color)
                .stroke(colors.borderColor);

            const internalLoss = typeof subClient.totalinternallosse === 'number' ? subClient.totalinternallosse.toFixed(1) : subClient.totalinternallosse;
            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(subClient.totalinternallosse < 0 ? colors.redText : colors.blackText)
                .text(internalLoss, startX + (2 * colWidths.subExport), currentY + 7, {
                    width: colWidths.subExport,
                    align: 'center'
                });

            // Loss in% total (red if negative, with special background)
            const lossPercent = typeof subClient.totallossinparsantege === 'number' ? subClient.totallossinparsantege.toFixed(1) + '%' : subClient.totallossinparsantege;
            const lossBgColor = subClient.totallossinparsantege < 0 ? colors.negativeLossBg : color;
            
            doc.rect(startX + (3 * colWidths.subExport), currentY, colWidths.subExport, rowHeights.totalsRow)
                .fill(lossBgColor)
                .stroke(colors.borderColor);

            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(subClient.totallossinparsantege < 0 ? colors.redText : colors.blackText)
                .text(lossPercent, startX + (3 * colWidths.subExport), currentY + 7, {
                    width: colWidths.subExport,
                    align: 'center'
                });

            // Import total
            doc.rect(startX + (4 * colWidths.subExport), currentY, colWidths.subExport, rowHeights.totalsRow)
                .fill(color)
                .stroke(colors.borderColor);

            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.blackText)
                .text(typeof subClient.totalimport === 'number' ? subClient.totalimport.toFixed(1) : subClient.totalimport, 
                      startX + (4 * colWidths.subExport), currentY + 7, {
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
                .fill(bgColor)
                .stroke(colors.borderColor);

            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(textColor)
                .text(formattedValue, acLineStartX + (i * colWidths.acLine), currentY + 7, {
                    width: colWidths.acLine,
                    align: 'center'
                });
        });

        // Finalize the PDF
        doc.end();

    } catch (error) {
        logger.error(`Error generating Daily Report PDF: ${error.message}`);
        res.status(500).json({
            message: 'Error generating Daily Report PDF',
            error: error.message
        });
    }
};

// Helper function to get month name
function getMonthName(month) {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return months[month - 1] || '';
}