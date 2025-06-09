const mongoose = require('mongoose');
const LossesCalculationData = require('../../models/v1/lossesCalculation.model');
const ExcelJS = require('exceljs');

exports.getLossesCalculationData = async (req, res) => {
    try {
        const { mainClientId, startMonth, startYear, endMonth, endYear } = req.body;

        if (!mainClientId || !startMonth || !startYear || !endMonth || !endYear) {
            return res.status(400).json({ message: 'Please provide mainClientId, startMonth, startYear, endMonth, and endYear.' });
        }

        // Validate month and year values
        if (
            startMonth < 1 || startMonth > 12 ||
            endMonth < 1 || endMonth > 12 ||
            startYear < 1900 || endYear < 1900 ||
            (endYear < startYear) ||
            (endYear === startYear && endMonth < startMonth)
        ) {
            return res.status(400).json({ message: 'Invalid month or year range.' });
        }

        // Query LossesCalculationData between the range
        const data = await LossesCalculationData.find({
            mainClientId: new mongoose.Types.ObjectId(mainClientId),
            $or: [
                { year: { $gt: startYear, $lt: endYear } },
                { year: startYear, month: { $gte: startMonth } },
                { year: endYear, month: { $lte: endMonth } }
            ]
        }).sort({ year: 1, month: 1 });

        // Filter strictly inside range
        const filteredData = data.filter(d =>
            (d.year > startYear || (d.year === startYear && d.month >= startMonth)) &&
            (d.year < endYear || (d.year === endYear && d.month <= endMonth))
        );

        const result = filteredData.map(entry => ({
            month: `${entry.month}-${entry.year}`,
            monthNum: entry.month,
            year: entry.year,
            mainClient: {
                mainClientDetail: {
                    name: entry.mainClient?.mainClientDetail?.name || '',
                    acCapacityKw: entry.mainClient?.mainClientDetail?.acCapacityKw || 0
                },
                grossInjectionMWH: entry.mainClient?.grossInjectionMWH || 0,
                drawlMWH: entry.mainClient?.drawlMWH || 0
            },
            mainClientName: entry.mainClient?.mainClientDetail?.name || '',
            mainClientInjectedUnitsKWh: Math.round((entry.mainClient?.grossInjectionMWH || 0) * 1000),
            mainClientDrawlUnitsKWh: Math.round((entry.mainClient?.drawlMWH || 0) * 1000),
            subClients: entry.subClient?.map(sc => ({
                name: sc.name,
                InjectedUnitsKWh: Math.round((sc.subClientsData?.grossInjectionMWHAfterLosses || 0) * 1000),
                DrawlUnitsKWh: Math.round((sc.subClientsData?.drawlMWHAfterLosses || 0) * 1000)
            })) || []
        }));

        const mainClientName = result.length > 0 ? (result[0].mainClient?.mainClientDetail?.name || 'Client') : 'Client';
        const acCapacityKw = result.length > 0 ? (result[0].mainClient?.mainClientDetail?.acCapacityKw || 0) : 0;
        const acCapacityMW = (acCapacityKw / 1000).toFixed(2);

        // Get all unique sub-clients across all entries
        const allSubClients = new Map();
        result.forEach(entry => {
            entry.subClients.forEach(sc => {
                if (!allSubClients.has(sc.name)) {
                    allSubClients.set(sc.name, { name: sc.name });
                }
            });
        });

        // Convert to array (up to 4 sub-clients)
        const subClients = Array.from(allSubClients.values()).slice(0, 4);

        // Create Excel workbook and worksheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Generation Report');
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
            paperSize: 9,
            orientation: 'landscape'
        };
        worksheet.pageSetup = worksheetSetup;

        // Set column widths
        const baseColumns = [
            { width: 8 },  // Sr. No.
            { width: 13 }, // Month Name
            { width: 15 }, // Injected Units (Main Client)
            { width: 15 }, // Drawl Units (Main Client)
        ];

        const subClientColumns = [];
        for (let i = 0; i < subClients.length; i++) {
            subClientColumns.push({ width: 15 });
            subClientColumns.push({ width: 15 });
        }

        worksheet.columns = [...baseColumns, ...subClientColumns];

        const totalColumns = 4 + (subClients.length * 2);

        // Format date strings
        const formatDateString = (month, year) => {
            const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
            return `${months[month - 1]}-${String(year).slice(-2)}`;
        };

        const startDateStr = formatDateString(startMonth, startYear);
        const endDateStr = formatDateString(endMonth, endYear);

        // Add title row
        const titleRow = worksheet.addRow([`${mainClientName} - ${acCapacityMW} MW AC Generation Details`]);
        worksheet.mergeCells(`A1:${String.fromCharCode(64 + totalColumns)}1`);
        titleRow.font = { name: 'Times New Roman', bold: true, size: 14 };
        titleRow.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        titleRow.height = 45;
        titleRow.eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
            cell.border = {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
            };
        });

        // Add date range row
        const dateRangeText = `Report From ${startDateStr} to ${endDateStr}`;
        const dateRangeRow = worksheet.addRow([dateRangeText]);
        worksheet.mergeCells(`A2:${String.fromCharCode(64 + totalColumns)}2`);
        dateRangeRow.height = 35;
        dateRangeRow.font = { name: 'Times New Roman', bold: true, size: 14 };
        dateRangeRow.alignment = { horizontal: 'center', vertical: 'middle' };
        dateRangeRow.eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } };
            cell.border = {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
            };
        });

        // Create header rows with vertical merging
        const headerRow1 = worksheet.addRow([]);

        // Set values and merge cells vertically for row 3 and 4
        worksheet.getCell('A3').value = 'Sr. No.';
        worksheet.mergeCells('A3:A4');  // Merge Sr. No. vertically

        worksheet.getCell('B3').value = 'Month Name';
        worksheet.mergeCells('B3:B4');  // Merge Month Name vertically
        worksheet.getCell('B3').alignment = { wrapText: true };

        worksheet.getCell('C3').value = 'TOTAL GENERATION';
        worksheet.mergeCells('C3:D3');  // Merge horizontally for Total Generation

        // Add sub-client headers (merged horizontally)
        subClients.forEach((sc, i) => {
            const startCol = 5 + (i * 2);
            const startColChar = String.fromCharCode(64 + startCol);
            const endColChar = String.fromCharCode(65 + startCol);

            worksheet.getCell(`${startColChar}3`).value = sc.name;
            worksheet.mergeCells(`${startColChar}3:${endColChar}3`);
        });

        // Style for header row 3
        worksheet.getRow(3).height = 80;
        worksheet.getRow(3).eachCell(cell => {
            cell.font = { name: 'Times New Roman', bold: true, size: 12 };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
            };
        });

        // Sub-header row (row 4)
        worksheet.getCell('C4').value = 'Injected Units (KWh)';
        worksheet.getCell('D4').value = 'Drawl Units S/S (KWh)';

        // Add sub-client column headers
        subClients.forEach((sc, i) => {
            const injectedCol = 5 + (i * 2);
            const drawlCol = 6 + (i * 2);

            worksheet.getCell(`${String.fromCharCode(64 + injectedCol)}4`).value = 'Injected Units (KWh)';
            worksheet.getCell(`${String.fromCharCode(64 + drawlCol)}4`).value = 'Drawl Units S/S (KWh)';
        });

        // Style for header row 4
        worksheet.getRow(4).height = 50;
        worksheet.getRow(4).eachCell(cell => {
            cell.font = { name: 'Times New Roman', bold: true, size: 11 };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
            };
        });

        // Generate all months in the range
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        let currentMonth = startMonth;
        let currentYear = startYear;
        let rowIndex = 1;

        while (currentYear < endYear || (currentYear === endYear && currentMonth <= endMonth)) {
            const monthName = `${months[currentMonth - 1]}-${String(currentYear).slice(-2)}`;

            const monthData = result.find(r => r.monthNum === currentMonth && r.year === currentYear);

            let mainClientInjected = 0;
            let mainClientDrawl = 0;
            const subClientValues = {};

            subClients.forEach(sc => {
                subClientValues[sc.name] = { injected: 0, drawl: 0 };
            });

            if (monthData) {
                mainClientInjected = monthData.mainClientInjectedUnitsKWh || 0;
                mainClientDrawl = monthData.mainClientDrawlUnitsKWh || 0;

                monthData.subClients.forEach(sc => {
                    if (subClientValues[sc.name]) {
                        subClientValues[sc.name].injected = sc.InjectedUnitsKWh || 0;
                        subClientValues[sc.name].drawl = sc.DrawlUnitsKWh || 0;
                    }
                });
            }

            const rowValues = [
                rowIndex,
                monthName,
                mainClientInjected,
                mainClientDrawl
            ];

            subClients.forEach(sc => {
                rowValues.push(subClientValues[sc.name].injected);
                rowValues.push(subClientValues[sc.name].drawl);
            });

            const dataRow = worksheet.addRow(rowValues);
            dataRow.eachCell(cell => {
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                // Format numbers to show exact values without decimals
                if (cell.value !== null && typeof cell.value === 'number' && cell.col > 2) {
                    cell.numFmt = '0';
                }
            });

            rowIndex++;

            if (currentMonth === 12) {
                currentMonth = 1;
                currentYear++;
            } else {
                currentMonth++;
            }
        }

        // Add empty row
        worksheet.addRow([]);

        // Add total row
        const lastDataRowIndex = rowIndex + 4;
        const totalRowValues = ['', 'TOTAL'];

        totalRowValues.push(
            { formula: `SUM(C5:C${lastDataRowIndex})` },
            { formula: `SUM(D5:D${lastDataRowIndex})` }
        );

        subClients.forEach((sc, index) => {
            const injectedCol = String.fromCharCode(69 + (index * 2));
            const drawlCol = String.fromCharCode(70 + (index * 2));

            totalRowValues.push(
                { formula: `SUM(${injectedCol}5:${injectedCol}${lastDataRowIndex})` },
                { formula: `SUM(${drawlCol}5:${drawlCol}${lastDataRowIndex})` }
            );
        });

        const totalRow = worksheet.addRow(totalRowValues);
        totalRow.eachCell(cell => {
            cell.font = { name: 'Times New Roman', bold: true, size: 12 };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
            };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD9D9D9' }
            };
            // Format numbers to show exact values without decimals
            if (cell.value !== null && typeof cell.value === 'number' && cell.col > 2) {
                cell.numFmt = '0';
            }
        });

        // Apply border and alignment style for all data rows
        worksheet.eachRow((row, rowNum) => {
            if (rowNum >= 5) {
                row.height = 42;
                row.eachCell(cell => {
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    cell.font = { name: 'Times New Roman' };
                    cell.border = {
                        top: { style: 'medium' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            }
        });

        // Add outer medium border
        const lastRow = worksheet.lastRow.number;
        const firstDataRow = 1;
        const lastColLetter = String.fromCharCode(64 + totalColumns);

        for (let i = firstDataRow; i <= lastRow; i++) {
            const row = worksheet.getRow(i);

            const leftCell = row.getCell(1);
            leftCell.border = {
                ...leftCell.border,
                left: { style: 'medium' }
            };

            const rightCell = row.getCell(totalColumns);
            rightCell.border = {
                ...rightCell.border,
                right: { style: 'medium' }
            };
        }

        worksheet.getRow(firstDataRow).eachCell(cell => {
            cell.border = {
                ...cell.border,
                top: { style: 'medium' }
            };
        });

        worksheet.getRow(lastRow).eachCell(cell => {
            cell.border = {
                ...cell.border,
                bottom: { style: 'medium' }
            };
        });

        const sanitize = (str) => {
            if (typeof str !== 'string') return '';
            return str.replace(/[^a-zA-Z0-9\s-]/g, '')
                .replace(/\s+/g, ' ')
                .trim();
        };

        // Get client name from the first record (since we already have the data)
        const clientName = result.length > 0
            ? result[0].mainClient?.mainClientDetail?.name || 'Report'
            : 'Report';

        // Format month names
        const getMonthName = (monthNum) => {
            const months = ['January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'];
            return months[parseInt(monthNum) - 1] || monthNum;
        };

        // Build filename
        const sanitizedClient = sanitize(clientName);
        const startMonthName = getMonthName(startMonth);
        const endMonthName = getMonthName(endMonth);

        const fileName = `Yearly Report ${sanitizedClient} - ${startMonthName}-${startYear} to ${endMonthName}-${endYear}.xlsx`;

        // Set response headers
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fileName)}"`);
        res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error('Error generating Excel report:', error);
        res.status(500).json({ message: 'Server Error', error: error.message });
    }
};