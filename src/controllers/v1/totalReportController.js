const MainClient = require('../../models/v1/mainClient.model');
const MeterData = require('../../models/v1/meterData.model');
const TotalReport = require('../../models/v1/totalReport.model');
const logger = require('../../utils/logger');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const fs = require('fs');

// Helper function to calculate units from meter data
const calculateUnits = (meterData, mf, pn) => {
  let grossInjected = 0;
  let grossDrawl = 0;

  meterData.forEach(data => {
    data.dataEntries.forEach(entry => {
      const reactiveE = parseFloat(entry.parameters['Reactive(E) Total']) || 0;
      const reactiveI = parseFloat(entry.parameters['Reactive(I) Total']) || 0;
      const activeI = parseFloat(entry.parameters['Active(I) Total']) || 0;
      const activeE = parseFloat(entry.parameters['Active(E) Total']) || 0;

      if (pn === -1) {
        // Original calculation if pn is -1
        grossInjected += (activeE - reactiveE) * mf;
        grossDrawl += (activeI - reactiveI) * mf;
      } else {
        // Swapped calculation if pn is not -1
        grossDrawl += (activeE - reactiveE) * mf;
        grossInjected += (activeI - reactiveI) * mf;
      }
    });
  });

  return {
    grossInjectedUnits: parseFloat(grossInjected.toFixed(2)),
    grossDrawlUnits: parseFloat(grossDrawl.toFixed(2)),
    totalimport: parseFloat((grossInjected - grossDrawl).toFixed(2))
  };
};

// Function to generate the total report
exports.generateTotalReport = async (req, res) => {
    try {
        const { mainClientIds, month, year } = req.body;

        // Validate input parameters
        if (!mainClientIds || !Array.isArray(mainClientIds) || mainClientIds.length === 0) {
            return res.status(400).json({ message: 'Please provide at least one main client ID' });
        }

        // Step 0: Check for existing reports with all the requested mainClientIds
        logger.info("Checking for existing total report...");

        const existingReports = await TotalReport.find({
            month,
            year,
            'clients.mainClientId': { $in: mainClientIds }
        });

        let exactMatchReport = null;
        if (existingReports.length > 0) {
            exactMatchReport = existingReports.find(report => {
                const reportClientIds = report.clients.map(c => c.mainClientId.toString());
                const requestedClientIds = mainClientIds.map(id => id.toString());

                return reportClientIds.length === requestedClientIds.length &&
                    reportClientIds.every(id => requestedClientIds.includes(id));
            });
        }

        // If exact match found, return cached data
        if (exactMatchReport) {
            logger.info("Existing total report found with exact matching parameters. Returning cached data.");
            exactMatchReport.updatedAt = new Date(); // Update the timestamp
            await exactMatchReport.save();
            return res.status(200).json({
                message: 'Existing total report retrieved successfully (timestamp updated).',
                data: exactMatchReport
            });
        }

        // Step 1: Generate new report if no exact match found
        logger.info("Generating new total report...");
        const reportData = {
            month,
            year,
            clients: []
        };

        // Step 2: Process each main client
        for (const mainClientId of mainClientIds) {
            const existingClientReports = existingReports.filter(report =>
                report.clients.some(c => c.mainClientId.toString() === mainClientId.toString())
            );

            let existingClientData = null;
            if (existingClientReports.length > 0) {
                existingClientData = existingClientReports[0].clients.find(
                    c => c.mainClientId.toString() === mainClientId.toString()
                );
            }

            if (existingClientData) {
                logger.info(`Using cached data for main client ${mainClientId}`);
                reportData.clients.push(existingClientData);
                continue; // Skip to next client if data already exists
            }

            // Step 3: Fetch main client data from database
            const mainClient = await MainClient.findById(mainClientId);
            if (!mainClient) {
                logger.error(`Main Client not found: ${mainClientId}`);
                return res.status(404).json({ message: `Main Client not found: ${mainClientId}` });
            }

            // Step 4: Fetch meter data for both meters
            const abtMainMeterData = await MeterData.find({
                meterNumber: mainClient.abtMainMeter.meterNumber,
                meterType: 'abtMainMeter',
                month,
                year
            });

            const abtCheckMeterData = await MeterData.find({
                meterNumber: mainClient.abtCheckMeter.meterNumber,
                meterType: 'abtCheckMeter',
                month,
                year
            });

            // Step 5: Validate meter data availability, but do NOT stop if missing
            const hasMainMeterData = abtMainMeterData.length > 0;
            const hasCheckMeterData = abtCheckMeterData.length > 0;

            if (!hasMainMeterData && !hasCheckMeterData) {
                logger.warn(`No meter data found for both ABT meters for client ${mainClientId} in ${month}-${year}. Including client with zero data.`);
            }

            // Step 6: Calculate values for both meters, default to zero if missing
            const abtMainMeter = hasMainMeterData
                ? {
                    meterNumber: mainClient.abtMainMeter.meterNumber,
                    ...calculateUnits(abtMainMeterData, mainClient.mf, mainClient.pn)
                }
                : {
                    meterNumber: mainClient.abtMainMeter.meterNumber || 'N/A',
                    grossInjectedUnits: 0,
                    grossDrawlUnits: 0,
                    totalimport: 0
                };

            const abtCheckMeter = hasCheckMeterData
                ? {
                    meterNumber: mainClient.abtCheckMeter.meterNumber,
                    ...calculateUnits(abtCheckMeterData, mainClient.mf, mainClient.pn)
                }
                : {
                    meterNumber: mainClient.abtCheckMeter.meterNumber || 'N/A',
                    grossInjectedUnits: 0,
                    grossDrawlUnits: 0,
                    totalimport: 0
                };

            // Step 7: Calculate differences - set to 0 if either meter is missing
            let difference;
            if (!hasMainMeterData || !hasCheckMeterData) {
                difference = {
                    grossInjectedUnits: 0,
                    grossDrawlUnits: 0
                };
            } else {
                difference = {
                    grossInjectedUnits: parseFloat((abtMainMeter.grossInjectedUnits - abtCheckMeter.grossInjectedUnits).toFixed(2)),
                    grossDrawlUnits: parseFloat((abtMainMeter.grossDrawlUnits - abtCheckMeter.grossDrawlUnits).toFixed(2))
                };
            }

            // Step 8: Add to report
            reportData.clients.push({
                mainClientId,
                mainClientDetail: {
                    name: mainClient.name,
                    subTitle: mainClient.subTitle,
                    abtMainMeter: mainClient.abtMainMeter,
                    abtCheckMeter: mainClient.abtCheckMeter,
                    voltageLevel: mainClient.voltageLevel,
                    acCapacityKw: mainClient.acCapacityKw,
                    dcCapacityKwp: mainClient.dcCapacityKwp,
                    noOfModules: mainClient.noOfModules,
                    sharingPercentage: mainClient.sharingPercentage,
                    contactNo: mainClient.contactNo,
                    email: mainClient.email,
                    mf: mainClient.mf
                },
                abtMainMeter,
                abtCheckMeter,
                difference
            });
        }

        // Step 9: Save the generated report
        const totalReport = new TotalReport(reportData);
        await totalReport.save();

        // Step 10: Return success response
        res.status(201).json({
            message: 'Total Report generated successfully',
            data: totalReport
        });

    } catch (error) {
        logger.error(`Error generating Total Report: ${error.message}`);
        res.status(500).json({
            message: 'Error generating Total Report',
            error: error.message
        });
    }
};


// Controller to get latest 10 total reports
// Controller to get latest 10 total reports
exports.getLatestTotalReports = async (req, res) => {
    try {
        // Step 1: Get latest 10 reports sorted by update date (newest first)
        logger.info("Fetching the latest 10 total reports...");

        const latestReports = await TotalReport.find({})
            .sort({ updatedAt: -1 }) // Sorting by last updated date
            .limit(10) // Limiting to 10 reports
            .lean(); // Convert to plain JavaScript objects for better performance

        // Step 2: Check if reports exist
        if (!latestReports || latestReports.length === 0) {
            logger.warn("No reports found.");
            return res.status(404).json({
                message: 'No reports found'
            });
        }

        // Step 3: Transform the data to include only required fields
        const simplifiedReports = latestReports.map(report => {
            // Extract client names from the report
            const clientNames = report.clients.map(client =>
                client.mainClientDetail?.name || 'N/A' // Safeguard in case name is missing
            ).join(', ');

            return {
                id: report._id,
                month: report.month,
                year: report.year,
                clientNames: clientNames,
                lastUpdated: report.updatedAt, // Last update timestamp
                generatedAt: report.createdAt, // Original creation timestamp
                clientCount: report.clients.length // Count of clients in the report
            };
        });

        // Step 4: Return the simplified report data
        res.status(200).json({
            message: 'Latest 10 reports retrieved successfully (sorted by last update)',
            data: simplifiedReports
        });

    } catch (error) {
        // Log the error details for debugging
        logger.error(`Error fetching latest reports: ${error.message}`);

        // Return a generic error message to the client
        res.status(500).json({
            message: 'Error fetching latest reports',
            error: error.message
        });
    }
};
exports.downloadTotalReportExcel = async (req, res) => {
    try {
        const { totalReportId } = req.params;

        // Fetch the total report
        const totalReport = await TotalReport.findById(totalReportId)
            .populate('clients.mainClientId', 'mainClientDetail abtMainMeter abtCheckMeter mf');

        if (!totalReport) {
            return res.status(404).json({ message: 'Total Report not found' });
        }

        // Create a new workbook
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Total Generation Report');
        worksheet.pageSetup = {
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
            fitToWidth: 1,       // Fit all columns to one page width
            fitToHeight: 100,     // Allow multiple pages if content is long
            paperSize: 9,         // A4
            orientation: 'portrait',
            printTitlesRow: '1:3', // Repeat header rows on each page
            footer: {
                firstFooter: "&P of &N", // "1 of 3" format
                oddFooter: "&P of &N",   // For odd pages
                evenFooter: "&P of &N"   // For even pages
            }
        };

        // Set column widths
        worksheet.columns = [
            { key: 'A', width: 50 },  // Column A width 50
            { key: 'B', width: 20 },  // Column B width 20
            { key: 'C', width: 20 },  // Column C width 20
            { key: 'D', width: 15 }   // Column D width 15
        ];

        // Define border styles
        const thickBorder = {
            top: { style: 'medium' }, // 1.5px equivalent
            left: { style: 'medium' },
            bottom: { style: 'medium' },
            right: { style: 'medium' }
        };

        const thinBorder = {
            top: { style: 'thin' }, // 1px
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // 1. First row - Note about units
        worksheet.mergeCells('A1:D1');
        const noteCell = worksheet.getCell('A1');
        noteCell.value = 'Note: - All Units are in KWH';
        noteCell.font = {
            italic: true,
            size: 14,
            name: 'Times New Roman'
        };
        noteCell.alignment = { horizontal: 'left' };

        // 2. Second row - Title and date
        // Title (A2:B2)
        worksheet.mergeCells('A2:B2');
        const titleCell = worksheet.getCell('A2');
        worksheet.getRow(2).height = 45;
        titleCell.value = 'TOTAL GENERATION REPORT';
        titleCell.font = {
            bold: true,
            size: 22,
            color: { argb: 'FFFF0000' },
            name: 'Times New Roman'
        };
        titleCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD9D9D9' }
        };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        titleCell.border = thickBorder;

        // Date (C2:D2)
        worksheet.mergeCells('C2:D2');
        const dateCell = worksheet.getCell('C2');
        dateCell.value = `${getMonthName(totalReport.month)}-${totalReport.year.toString().slice(-2)}`;
        dateCell.font = {
            bold: true,
            size: 22,
            color: { argb: 'FFFF0000' },
            name: 'Times New Roman'
        };
        dateCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD9D9D9' }
        };
        dateCell.alignment = { horizontal: 'center', vertical: 'middle' };
        dateCell.border = thickBorder;

        // 3. Third row - Date range (no gap after this row)
        worksheet.mergeCells('A3:D3');
        const dateRangeCell = worksheet.getCell('A3');
        worksheet.getRow(3).height = 45; // 
        const lastDay = new Date(totalReport.year, totalReport.month, 0).getDate();
        dateRangeCell.value = `01-${String(totalReport.month).padStart(2, '0')}-${totalReport.year}  to  ${lastDay}-${String(totalReport.month).padStart(2, '0')}-${totalReport.year}`;
        dateRangeCell.font = {
            bold: true,
            size: 22,
            color: { argb: 'FFFF0000' },
            name: 'Times New Roman'
        };
        dateRangeCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD9D9D9' }
        };
        dateRangeCell.alignment = { horizontal: 'center', vertical: 'middle' };
        dateRangeCell.border = thickBorder;

        let currentRow = 4; // Start immediately after date range row (no gap)

        // Process each client
        totalReport.clients.forEach(client => {
            // 4. Client header row (merged A-D) - Set row height to 36
            worksheet.mergeCells(`A${currentRow}:D${currentRow}`);
            const row4 = worksheet.getRow(currentRow);
            row4.height = 50; // Set row height to 36

            const clientHeaderCell = worksheet.getCell(`A${currentRow}`);
            clientHeaderCell.value = `${client.mainClientDetail.name} (Lead Generator) ${(client.mainClientDetail.acCapacityKw / 1000).toFixed(2)} MW AC\nGeneration Details - ${client.mainClientDetail.subTitle}`;
            clientHeaderCell.font = {
                bold: true,
                size: 14,
                color: { argb: 'FF000000' },
                name: 'Times New Roman'
            };
            clientHeaderCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF92D050' }
            };
            clientHeaderCell.alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
            };
            clientHeaderCell.border = thickBorder;
            currentRow++;

            // 5. Meter headers row
            // Column A (merged with next row)
            worksheet.getRow(currentRow).height = 25;
            worksheet.mergeCells(`A${currentRow}:A${currentRow + 1}`);
            const meterLabelCell = worksheet.getCell(`A${currentRow}`);
            meterLabelCell.value = 'Meter Number';
            meterLabelCell.font = {
                bold: true,
                size: 10,
                name: 'Times New Roman'
            };
            meterLabelCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFF00' }
            };
            meterLabelCell.alignment = {
                horizontal: 'center',
                vertical: 'middle'
            };
            meterLabelCell.border = thickBorder; // Changed from thinBorder to thickBorder

            // Column B (MAIN METER)
            const mainMeterHeader = worksheet.getCell(`B${currentRow}`);
            mainMeterHeader.value = 'MAIN METER';
            mainMeterHeader.font = {
                bold: true,
                size: 10,
                name: 'Times New Roman'
            };
            mainMeterHeader.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFF00' }
            };
            mainMeterHeader.alignment = { horizontal: 'center', vertical: 'middle' };
            mainMeterHeader.border = thickBorder; // Changed from thinBorder to thickBorder

            // Column C (CHECK METER)
            const checkMeterHeader = worksheet.getCell(`C${currentRow}`);
            checkMeterHeader.value = 'CHECK METER';
            checkMeterHeader.font = {
                bold: true,
                size: 10,
                name: 'Times New Roman'
            };
            checkMeterHeader.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFF00' }
            };
            checkMeterHeader.alignment = { horizontal: 'center', vertical: 'middle' };
            checkMeterHeader.border = thickBorder; // Changed from thinBorder to thickBorder

            // Column D (Difference - merged with next row)
            worksheet.mergeCells(`D${currentRow}:D${currentRow + 1}`);
            const diffHeader = worksheet.getCell(`D${currentRow}`);
            diffHeader.value = 'Difference';
            diffHeader.font = {
                bold: true,
                size: 10,
                color: { argb: 'FFFF0000' },
                name: 'Times New Roman'
            };
            diffHeader.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFF00' }
            };
            diffHeader.alignment = {
                horizontal: 'center',
                vertical: 'middle'
            };
            diffHeader.border = thickBorder; // Changed from thinBorder to thickBorder
            currentRow++;

            // 6. Meter numbers row
            // MAIN METER number
            worksheet.getRow(currentRow).height = 25;
            const mainMeterNum = worksheet.getCell(`B${currentRow}`);
            mainMeterNum.value = client.abtMainMeter.meterNumber;
            mainMeterNum.font = {
                bold: true,
                size: 10,
                name: 'Times New Roman'
            };
            mainMeterNum.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFF00' }
            };
            mainMeterNum.alignment = { horizontal: 'center', vertical: 'middle' };
            mainMeterNum.border = thickBorder; // Changed from thinBorder to thickBorder

            // CHECK METER number
            const checkMeterNum = worksheet.getCell(`C${currentRow}`);
            checkMeterNum.value = client.abtCheckMeter.meterNumber;
            checkMeterNum.font = {
                bold: true,
                size: 10,
                name: 'Times New Roman'
            };
            checkMeterNum.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFF00' }
            };
            checkMeterNum.alignment = { horizontal: 'center', vertical: 'middle' };
            checkMeterNum.border = thickBorder; // Changed from thinBorder to thickBorder

            // Column A (already merged from previous row)
            worksheet.getCell(`A${currentRow}`).border = thickBorder; // Changed from thinBorder to thickBorder

            // Column D (already merged from previous row)
            worksheet.getCell(`D${currentRow}`).border = thickBorder; // Changed from thinBorder to thickBorder

            currentRow++;

            // 7-9. Data rows (rows 5-10 in your numbering)
            const dataRows = [
                // 7. Gross Injected Units
                {
                    label: `Gross Injected Units to ${client.mainClientDetail.subTitle}`,
                    mainValue: client.abtMainMeter.grossInjectedUnits,
                    checkValue: client.abtCheckMeter.grossInjectedUnits,
                    diffValue: client.difference.grossInjectedUnits
                },
                // 8. Gross Drawl Units
                {
                    label: `Gross Drawl Units from ${client.mainClientDetail.subTitle}`,
                    mainValue: client.abtMainMeter.grossDrawlUnits,
                    checkValue: client.abtCheckMeter.grossDrawlUnits,
                    diffValue: client.difference.grossDrawlUnits
                },
                // 9. Net Injected Units
                {
                    label: `Net Injected Units to ${client.mainClientDetail.subTitle}`,
                    mainValue: client.abtMainMeter.totalimport,
                    checkValue: client.abtCheckMeter.totalimport,
                    diffValue: '',
                    bold: true
                }
            ];

            // want to set vertically centered text in the cell


            dataRows.forEach((rowData, index) => {
                const row = worksheet.getRow(currentRow);
                row.height = 25; // Set row height to 25

                // Set values
                row.getCell('A').value = rowData.label;
                row.getCell('B').value = rowData.mainValue;
                row.getCell('C').value = rowData.checkValue;
                row.getCell('D').value = rowData.diffValue;
                row.getCell('A').alignment = { vertical: 'middle', horizontal: 'left' };
                row.getCell('B').alignment = { vertical: 'middle', horizontal: 'right' };
                row.getCell('C').alignment = { vertical: 'middle', horizontal: 'right' };
                row.getCell('D').alignment = { vertical: 'middle', horizontal: 'right' };

                // Set fonts (all 10px)
                row.getCell('A').font = {
                    size: 10,
                    name: 'Times New Roman',
                    bold: rowData.bold || false
                };
                row.getCell('B').font = {
                    size: 10,
                    name: 'Times New Roman',
                    bold: rowData.bold || false
                };
                row.getCell('C').font = {
                    size: 10,
                    name: 'Times New Roman',
                    bold: rowData.bold || false
                };
                row.getCell('D').font = {
                    size: 10,
                    name: 'Times New Roman',
                    bold: rowData.bold || false
                };

                // Set borders
                if (index === 2) { // 9th row (Net Injected)
                    ['A', 'B', 'C', 'D'].forEach(col => { // Added 'D' to have thick border all around
                        row.getCell(col).border = {
                            top: { style: 'medium' },
                            left: { style: 'medium' },
                            bottom: { style: 'medium' },
                            right: { style: 'medium' }
                        };
                    });
                } else {
                    ['A', 'B', 'C', 'D'].forEach(col => {
                        row.getCell(col).border = {
                            top: { style: 'thin' },
                            left: { style: 'medium' },
                            bottom: { style: 'thin' },
                            right: { style: 'medium' }
                        };
                    });
                }

                currentRow++;
            });

            // 10. Empty row (merged A-D)
            worksheet.mergeCells(`A${currentRow}:D${currentRow}`);
            currentRow++;
        });

        // Apply outer border to all cells from row 2 to last row
        for (let i = 2; i < currentRow; i++) {
            const row = worksheet.getRow(i);
            ['A', 'B', 'C', 'D'].forEach(col => {
                const cell = row.getCell(col);
                if (!cell.border) {
                    cell.border = thinBorder;
                }
                // Ensure thick borders stay thick
                if (cell.border.top && cell.border.top.style === 'medium') {
                    cell.border = { ...cell.border, ...thickBorder };
                }
            });
        }

        // Format the filename with proper sanitization
        const sanitize = (str) => str.replace(/[^a-zA-Z0-9\s-]/g, '').replace(/\s+/g, ' ').trim();

        // Get month name (ensure this matches your frontend expectation)
        const getFullMonthName = (monthNumber) => {
            const months = [
                'January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'
            ];
            return months[monthNumber - 1] || '';
        };

        const formattedMonth = getFullMonthName(totalReport.month);
        const formattedYear = totalReport.year;

        // Build the sanitized filename
        const fileName = `Total Generation Unit Sheet - ${sanitize(formattedMonth)}-${formattedYear}.xlsx`;

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

        // Send the workbook
        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        logger.error(`Error generating Excel report: ${error.message}`);
        res.status(500).json({
            message: 'Error generating Excel report',
            error: error.message
        });
    }
};



exports.downloadTotalReportPDF = async (req, res) => {
    try {
        const { totalReportId } = req.params;
        // Fetch the total report
        const totalReport = await TotalReport.findById(totalReportId)
            .populate('clients.mainClientId', 'mainClientDetail abtMainMeter abtCheckMeter mf');

        if (!totalReport) {
            return res.status(404).json({ message: 'Total Report not found' });
        }

        // Sanitization function
        const sanitize = (str) => {
            if (typeof str !== 'string') return '';
            return str.replace(/[^a-zA-Z0-9\s-]/g, '')
                .replace(/\s+/g, ' ')
                .trim();
        };

        // Get month name
        const getFullMonthName = (monthNumber) => {
            const months = [
                'January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'
            ];
            return months[monthNumber - 1] || '';
        };

        // Build filename
        const formattedMonth = getFullMonthName(totalReport.month);
        const formattedYear = totalReport.year;
        const fileName = `Total Generation Unit Sheet - ${sanitize(formattedMonth)}-${formattedYear}.pdf`;

        // Set response headers
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader(
            'Content-Disposition',
            `attachment; filename="${fileName}"` // Removed encodeURIComponent to avoid double encoding
        );
        res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

        // Create a PDF document with strict A4 landscape dimensions
        const doc = new PDFDocument({
            size: [842, 595], // Exact A4 landscape dimensions in points (595 x 842)
            layout: 'landscape',
            margins: { top: 30, bottom: 30, left: 30, right: 30 }
        });

        // Pipe the PDF to the response
        doc.pipe(res);

        // ====================
        // PDF STYLING - FINAL FIXES
        // ====================
        const titleFont = 'Times-Roman';
        const headerFont = 'Helvetica-Bold';
        const dataFont = 'Helvetica';
        const titleFontSize = 16;
        const headerFontSize = 10;
        const dataFontSize = 9;

        // Final color scheme with guaranteed contrast
        const colors = {
            titleBg: '#D9D9D9',      // Light gray
            clientHeaderBg: '#92D050',// Green
            meterHeaderBg: '#FFFF00', // Yellow
            redText: '#FF0000',       // Red
            blackText: '#000000',     // Black (for all text)
            whiteBg: '#FFFFFF',       // White
            borderColor: '#000000',   // Black borders
            valueText: '#000000'      // Black for all values
        };

        // Final column widths
        const colWidths = {
            label: 270,   // For long labels
            mainMeter: 90,
            checkMeter: 90,
            difference: 80
        };

        // Final row heights
        const rowHeights = {
            title: 22,
            clientHeader: 40,
            meterHeader: 20,
            dataRow: 18
        };

        // Starting position
        let currentY = 30;

        // ====================
        // HEADER SECTION - FINAL
        // ====================

        // 1. Note about units
        doc.font('Times-Italic')
            .fontSize(10)
            .fillColor(colors.blackText)
            .text('Note: - All Units are in KWH', 30, currentY);

        currentY += 15;

        // 2. Title and date section
        const titleWidth = colWidths.label + colWidths.mainMeter;
        const dateWidth = colWidths.checkMeter + colWidths.difference;

        // Title background
        doc.rect(30, currentY, titleWidth, rowHeights.title)
            .fillAndStroke(colors.titleBg, colors.borderColor);

        // Date background
        doc.rect(30 + titleWidth, currentY, dateWidth, rowHeights.title)
            .fillAndStroke(colors.titleBg, colors.borderColor);

        // Title text (black for visibility)
        doc.font(titleFont)
            .fontSize(titleFontSize)
            .fillColor(colors.blackText) // Changed from red to black
            .text('TOTAL GENERATION REPORT', 30, currentY + 5, {
                width: titleWidth,
                align: 'center'
            });

        // Date text (black for visibility)
        doc.fillColor(colors.blackText)
            .text(`${getMonthName(totalReport.month)}-${totalReport.year.toString().slice(-2)}`,
                30 + titleWidth, currentY + 5, {
                width: dateWidth,
                align: 'center'
            });

        currentY += rowHeights.title;

        // 3. Date range
        const lastDay = new Date(totalReport.year, totalReport.month, 0).getDate();
        const dateRange = `01-${String(totalReport.month).padStart(2, '0')}-${totalReport.year}  to  ${lastDay}-${String(totalReport.month).padStart(2, '0')}-${totalReport.year}`;

        doc.rect(30, currentY,
            colWidths.label + colWidths.mainMeter + colWidths.checkMeter + colWidths.difference,
            rowHeights.title)
            .fillAndStroke(colors.titleBg, colors.borderColor);

        doc.font(titleFont)
            .fontSize(titleFontSize)
            .fillColor(colors.blackText) // Changed from red to black
            .text(dateRange, 30, currentY + 5, {
                width: colWidths.label + colWidths.mainMeter + colWidths.checkMeter + colWidths.difference,
                align: 'center'
            });

        currentY += rowHeights.title + 15;

        // ====================
        // CLIENT SECTIONS - FINAL FIXES
        // ====================
        totalReport.clients.forEach((client, clientIndex) => {
            // Check if we need a new page (entire client section must fit)
            const estimatedHeight = rowHeights.clientHeader + (rowHeights.meterHeader * 2) + (rowHeights.dataRow * 3) + 20;

            if (currentY + estimatedHeight > doc.page.height - 30) {
                doc.addPage({ size: [842, 595], layout: 'landscape' });
                currentY = 30;
            }

            // Client header
            const clientTitle = `${client.mainClientDetail.name} (Lead Generator) ${(client.mainClientDetail.acCapacityKw / 1000).toFixed(2)} MW AC\nGeneration Details - ${client.mainClientDetail.subTitle}`;

            doc.rect(30, currentY,
                colWidths.label + colWidths.mainMeter + colWidths.checkMeter + colWidths.difference,
                rowHeights.clientHeader)
                .fillAndStroke(colors.clientHeaderBg, colors.borderColor);

            doc.font(headerFont)
                .fontSize(12)
                .fillColor(colors.blackText)
                .text(clientTitle, 30, currentY + 7, {
                    width: colWidths.label + colWidths.mainMeter + colWidths.checkMeter + colWidths.difference - 10,
                    align: 'center',
                    lineGap: 3
                });

            currentY += rowHeights.clientHeader;

            // ====================
            // METER TABLE - FINAL FIXES
            // ====================

            // Meter headers row
            // Label column (merged)
            doc.rect(30, currentY, colWidths.label, rowHeights.meterHeader * 2)
                .fillAndStroke(colors.meterHeaderBg, colors.borderColor);

            doc.font(headerFont)
                .fontSize(headerFontSize)
                .fillColor(colors.blackText) // Black text on yellow
                .text('Meter Number', 30, currentY + 15, {
                    width: colWidths.label,
                    align: 'center'
                });

            // Main Meter header (black text)
            doc.rect(30 + colWidths.label, currentY, colWidths.mainMeter, rowHeights.meterHeader)
                .fillAndStroke(colors.meterHeaderBg, colors.borderColor);

            doc.fillColor(colors.blackText) // Black text
                .text('MAIN METER', 30 + colWidths.label, currentY + 5, {
                    width: colWidths.mainMeter,
                    align: 'center',
                });

            // Check Meter header (black text)
            doc.rect(30 + colWidths.label + colWidths.mainMeter, currentY, colWidths.checkMeter, rowHeights.meterHeader)
                .fillAndStroke(colors.meterHeaderBg, colors.borderColor);

            doc.fillColor(colors.blackText) // Black text
                .text('CHECK METER', 30 + colWidths.label + colWidths.mainMeter, currentY + 5, {
                    width: colWidths.checkMeter,
                    align: 'center'
                });

            // Difference header (merged)
            doc.rect(30 + colWidths.label + colWidths.mainMeter + colWidths.checkMeter, currentY,
                colWidths.difference, rowHeights.meterHeader * 2)
                .fillAndStroke(colors.meterHeaderBg, colors.borderColor);

            doc.fillColor(colors.blackText) // Changed from red to black
                .text('Difference', 30 + colWidths.label + colWidths.mainMeter + colWidths.checkMeter, currentY + 15, {
                    width: colWidths.difference,
                    align: 'center'
                });

            currentY += rowHeights.meterHeader;

            // Meter numbers row - all black text on yellow
            // Main Meter number
            doc.rect(30 + colWidths.label, currentY, colWidths.mainMeter, rowHeights.meterHeader)
                .fillAndStroke(colors.meterHeaderBg, colors.borderColor);

            doc.fillColor(colors.blackText) // Black text
                .text(client.abtMainMeter.meterNumber, 30 + colWidths.label, currentY + 5, {
                    width: colWidths.mainMeter,
                    align: 'center'
                });

            // Check Meter number (black text)
            doc.rect(30 + colWidths.label + colWidths.mainMeter, currentY, colWidths.checkMeter, rowHeights.meterHeader)
                .fillAndStroke(colors.meterHeaderBg, colors.borderColor);

            doc.fillColor(colors.blackText) // Black text
                .text(client.abtCheckMeter.meterNumber, 30 + colWidths.label + colWidths.mainMeter, currentY + 5, {
                    width: colWidths.checkMeter,
                    align: 'center'
                });

            currentY += rowHeights.meterHeader;

            // ====================
            // DATA ROWS - FINAL FIXES
            // ====================
            const dataRows = [
                {
                    label: `Gross Injected Units to ${client.mainClientDetail.subTitle} S/S`,
                    mainValue: client.abtMainMeter.grossInjectedUnits,
                    checkValue: client.abtCheckMeter.grossInjectedUnits,
                    diffValue: client.difference.grossInjectedUnits
                },
                {
                    label: `Gross Drawl Units from ${client.mainClientDetail.subTitle} S/S`,
                    mainValue: client.abtMainMeter.grossDrawlUnits,
                    checkValue: client.abtCheckMeter.grossDrawlUnits,
                    diffValue: client.difference.grossDrawlUnits
                },
                {
                    label: `Net Injected Units to ${client.mainClientDetail.subTitle} S/S`,
                    mainValue: client.abtMainMeter.totalimport,
                    checkValue: client.abtCheckMeter.totalimport,
                    diffValue: '',
                    bold: true,
                    border: 'thick'
                }
            ];

            dataRows.forEach((rowData) => {
                // Calculate required height for wrapped text
                const textHeight = doc.font(dataFont)
                    .fontSize(dataFontSize)
                    .heightOfString(rowData.label, {
                        width: colWidths.label - 10
                    });

                const rowHeight = Math.max(rowHeights.dataRow, textHeight + 4);

                // Label cell
                doc.rect(30, currentY, colWidths.label, rowHeight)
                    .fillAndStroke(colors.whiteBg, colors.borderColor);

                doc.font(dataFont)
                    .fontSize(dataFontSize)
                    .fillColor(colors.blackText) // Black text
                    .font(rowData.bold ? 'Helvetica-Bold' : 'Helvetica')
                    .text(rowData.label, 30 + 5, currentY + 4, {
                        width: colWidths.label - 10,
                        lineGap: 2
                    });

                // Main Meter value (black text)
                doc.rect(30 + colWidths.label, currentY, colWidths.mainMeter, rowHeight)
                    .fillAndStroke(colors.whiteBg, colors.borderColor);

                doc.font(rowData.bold ? 'Helvetica-Bold' : 'Helvetica')
                    .fillColor(colors.blackText) // Black text
                    .text(rowData.mainValue, 30 + colWidths.label, currentY + (rowHeight / 2 - 5), {
                        width: colWidths.mainMeter,
                        align: 'center'
                    });

                // Check Meter value (black text)
                doc.rect(30 + colWidths.label + colWidths.mainMeter, currentY, colWidths.checkMeter, rowHeight)
                    .fillAndStroke(colors.whiteBg, colors.borderColor);

                doc.fillColor(colors.blackText) // Black text
                    .text(rowData.checkValue, 30 + colWidths.label + colWidths.mainMeter, currentY + (rowHeight / 2 - 5), {
                        width: colWidths.checkMeter,
                        align: 'center'
                    });

                // Difference value (black text)
                doc.rect(30 + colWidths.label + colWidths.mainMeter + colWidths.checkMeter, currentY,
                    colWidths.difference, rowHeight)
                    .fillAndStroke(colors.whiteBg, colors.borderColor);

                doc.fillColor(colors.blackText) // Black text
                    .text(rowData.diffValue, 30 + colWidths.label + colWidths.mainMeter + colWidths.checkMeter, currentY + (rowHeight / 2 - 5), {
                        width: colWidths.difference,
                        align: 'center'
                    });

                // Thick bottom border for last row
                if (rowData.border === 'thick') {
                    doc.moveTo(30, currentY + rowHeight)
                        .lineTo(30 + colWidths.label + colWidths.mainMeter + colWidths.checkMeter + colWidths.difference,
                            currentY + rowHeight)
                        .lineWidth(1.5)
                        .stroke(colors.borderColor);
                }

                currentY += rowHeight;
            });

            // Space between client sections
            currentY += 15;
        });

        // Finalize the PDF
        doc.end();

    } catch (error) {
        logger.error(`Error generating PDF report: ${error.message}`);
        res.status(500).json({
            message: 'Error generating PDF report',
            error: error.message
        });
    }
};

// Helper function to get month name
function getMonthName(monthNumber) {
    const months = [
        'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
        'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'
    ];
    return months[monthNumber - 1];
}
// Helper function to get full month name
function getFullMonthName(monthNumber) {
    const months = [
        'JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE',
        'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'
    ];
    return months[monthNumber - 1];
}