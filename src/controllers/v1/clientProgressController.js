// /src/controllers/v1/clientProgressController.js
const ClientProgress = require('../../models/v1/clientProgress.model');
const MainClient = require('../../models/v1/mainClient.model');
const ClientProgressField = require('../../models/v1/ClientProgressFIled.model');
const ExcelJS = require('exceljs');

// Create a new client progress entry
exports.createClientProgress = async (req, res) => {
  try {
    const newClientProgress = new ClientProgress(req.body);
    await newClientProgress.save();
    res.status(201).json(newClientProgress);
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
};

// Update a client progress entry by ID
exports.updateClientProgress = async (req, res) => {
  try {
    const updatedClientProgress = await ClientProgress.findByIdAndUpdate(
      req.params.id,
      req.body,
      { new: true }
    );
    if (!updatedClientProgress) {
      return res.status(404).json({ message: 'Client progress not found' });
    }
    res.json(updatedClientProgress);
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
};

// Get all client progress entries by month and year
exports.getClientProgressByMonthYear = async (req, res) => {
  try {
    const { month, year } = req.params;
    const clientProgress = await ClientProgress.find({ month, year }).populate('clients.clientId', 'name email'); // Populate clientId with client name and email
    if (!clientProgress.length) {
      return res.status(404).json({ message: 'No data found for this month and year' });
    }
    res.json(clientProgress);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};


exports.downloadClientProgressExcel = async (req, res) => {
  try {
    const { month, year } = req.params;

    // Load progress data for month-year
    const progressData = await ClientProgress.findOne({ month, year }).populate('clients.clientId', 'name');
    if (!progressData) {
      console.log('No client progress data found');
      return res.status(404).json({ message: 'No client progress data found for this month and year' });
    }
    console.log(`Loaded client progress for ${progressData.clients.length} clients`);

    // Load field template (static fields)
    const progressFieldData = await ClientProgressField.findOne();
    if (!progressFieldData || !progressFieldData.clients.length) {
      console.log('No client progress field data found');
      return res.status(404).json({ message: 'No client progress field data found' });
    }
    console.log(`Loaded client progress field template for ${progressFieldData.clients.length} clients`);

    // Use first client in fields as header template
    const firstClientFields = progressFieldData.clients[0];
    const stageOneFields = firstClientFields.stageOne.map(t => t.name);
    const stageTwoFields = firstClientFields.stageTwo.map(t => t.name);
    const stageThreeFields = firstClientFields.stageThree.map(t => t.name);
    const stageBillingFields = firstClientFields.stageBilling.map(t => t.name);

    // Compose headers with empty columns after each stage
    // We'll add empty '' between stages as visual separator columns
    let headers = [
      'Sr. No',          // changed here
      'Client Name',
      ...stageOneFields,
      '', // empty col after stage 1
      ...stageTwoFields,
      '', // empty col after stage 2
      ...stageThreeFields,
      '', // empty col after stage 3
      ...stageBillingFields,
      'Remark'
    ];

    // Calculate stage ranges for merging stage labels on row 5
    // Indices in headers are 1-based in ExcelJS (first column is 1)
    // So Sr No = col 1, Client Name = col 2
    const stage1Start = 3;
    const stage1End = stage1Start + stageOneFields.length - 1;
    const stage2Start = stage1End + 2; // skip 1 empty col
    const stage2End = stage2Start + stageTwoFields.length - 1;
    const stage3Start = stage2End + 2; // skip 1 empty col
    const stage3End = stage3Start + stageThreeFields.length - 1;
    const stageBillingStart = stage3End + 2; // skip 1 empty col
    const stageBillingEnd = stageBillingStart + stageBillingFields.length - 1;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Client Progress');

    // Set print page setup to landscape and fit to page width
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

    // ----- Add empty first row -----
    worksheet.addRow([]);

    // ----- Second row: merged header -----
    const secondRow = worksheet.addRow([]);
    worksheet.mergeCells('A2:E2');
    const titleCell = worksheet.getCell('A2');
    titleCell.value = 'Monthly Work Progress Report';
    titleCell.font = { name: 'Times New Roman', size: 16, bold: true, underline: true };
    titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
    worksheet.getRow(2).height = 22;


    const lastCol = String.fromCharCode(65 + headers.length - 1);
    worksheet.mergeCells(`F2:${lastCol}2`);
    const monthYearCell = worksheet.getCell('F2');
    const monthName = new Date(year, month - 1).toLocaleString('default', { month: 'short' });
    monthYearCell.value = `Month : - ${monthName}-${year}`;
    monthYearCell.font = { name: 'Times New Roman', size: 16, bold: true };
    monthYearCell.alignment = { horizontal: 'right', vertical: 'middle' };

    // ----- Third row empty -----
    worksheet.addRow([]);

    // ----- Fourth row: header row with your modified headers -----
    const headerRow = worksheet.addRow(headers);

    // Set specific styles for cells
    headerRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    headerRow.getCell(1).font = { name: 'Times New Roman', size: 11, bold: true, };
    headerRow.getCell(2).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    headerRow.getCell(2).font = { name: 'Times New Roman', size: 11, bold: true };

    // Adjust columns width for first two columns
    worksheet.getColumn(1).width = 7;  // Sr No col width 7
    worksheet.getColumn(2).width = 40; // Client Name col width 40

    // Adjust width for all stage fields (stage1Start to stageBillingEnd)
    for (let col = stage1Start; col <= stageBillingEnd; col++) {
      worksheet.getColumn(col).width = 7; // Set your desired width here
    }

    // For all other header cells (from 3 to last), set vertical text, font size 11, font "Times New Roman", bold, wrap text
    for (let i = 3; i <= headers.length; i++) {
      const cell = headerRow.getCell(i);
      cell.alignment = { vertical: 'bottom', horizontal: 'center', textRotation: 90, wrapText: true }; // vertical align bottom now
      cell.font = { name: 'Times New Roman', size: 11, bold: true };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
    }
    worksheet.getRow(4).height = 150;

    // ----- Fifth row: stage labels -----
    // Prepare a blank row with same number of columns as headerRow
    const stageRow = worksheet.addRow(new Array(headers.length).fill(''));

    // Merge and set stage labels horizontally with bold font, centered text
    const setStageLabel = (start, end, label) => {
      worksheet.mergeCells(5, start, 5, end);
      const cell = worksheet.getCell(5, start);
      cell.value = label;
      cell.font = { name: 'Times New Roman', size: 11, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    };

    setStageLabel(stage1Start, stage1End + 1, '1st Stage');      // +1 to include empty col after stage
    setStageLabel(stage2Start, stage2End + 1, '2nd Stage');
    setStageLabel(stage3Start, stage3End + 1, '3rd Stage');
    setStageLabel(stageBillingStart, stageBillingEnd, 'Billing Stage');

    // --- Set column widths for empty cols (3 empty cols)
    worksheet.getColumn(stage1End + 1).width = 2;
    worksheet.getColumn(stage2End + 1).width = 2;
    worksheet.getColumn(stage3End + 1).width = 2;

    // --- Add borders to stage label row cells
    worksheet.getRow(5).eachCell(cell => {
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
    });

    // Helper to safely get clientId string from various formats
    const getClientIdString = (clientId) => {
      if (!clientId) return null;
      if (typeof clientId === 'string') return clientId;
      if (clientId._id) return clientId._id.toString();
      if (clientId.toString) return clientId.toString();
      return null;
    };

    // Helper: find client data in array by clientId string
    const findClientData = (clientsArray, clientIdStr) => {
      return clientsArray.find(c => {
        const cId = getClientIdString(c.clientId);
        return cId === clientIdStr;
      });
    };

    // Function to process each stage's tasks and determine marks
    const processStage = (fieldTasks, progressTasks) => {
      return fieldTasks.map(fieldTask => {
        const fieldStatus = Boolean(fieldTask.status);
        let progressStatus = false;

        if (progressTasks && progressTasks.length) {
          const progressTask = progressTasks.find(t =>
            t.name.trim().toLowerCase() === fieldTask.name.trim().toLowerCase()
          );
          progressStatus = progressTask &&
            (progressTask.status === true ||
              (typeof progressTask.status === 'string' && progressTask.status.toLowerCase() === 'true'));
        }

        if (fieldStatus === false) {
          return { value: 'X', bgColor: 'FFFFFF00' }; // Yellow background
        }
        if (fieldStatus === true && progressStatus === true) {
          return { value: '✓', bgColor: '92D050' }; // Green background
        }
        return { value: '', bgColor: null };
      });
    };

    // Process each client
    for (let i = 0; i < progressFieldData.clients.length; i++) {
      const fieldClient = progressFieldData.clients[i];
      const clientIdStr = getClientIdString(fieldClient.clientId);
      const progressClient = findClientData(progressData.clients, clientIdStr);

      // Process each stage
      const stageOneMarks = processStage(fieldClient.stageOne, progressClient?.stageOne || []);
      const stageTwoMarks = processStage(fieldClient.stageTwo, progressClient?.stageTwo || []);
      const stageThreeMarks = processStage(fieldClient.stageThree, progressClient?.stageThree || []);
      const stageBillingMarks = processStage(fieldClient.stageBilling, progressClient?.stageBilling || []);

      // Combine all marks with empty columns between stages
      const allMarks = [
        ...stageOneMarks,
        { value: '', bgColor: null }, // Empty column after stage 1
        ...stageTwoMarks,
        { value: '', bgColor: null }, // Empty column after stage 2
        ...stageThreeMarks,
        { value: '', bgColor: null }, // Empty column after stage 3
        ...stageBillingMarks
      ];

      // Build row values - start from index 1 (ExcelJS is 1-based)
      const rowValues = [];
      rowValues[1] = i + 1; // Sr. No (column A)
      rowValues[2] = fieldClient.clientName; // Client Name (column B)

      // Add marks to row values starting from column C (index 3)
      allMarks.forEach((mark, idx) => {
        rowValues[3 + idx] = mark.value;
      });

      // Add remark (last column)
      const remarkColIndex = 3 + allMarks.length;
      rowValues[remarkColIndex] = fieldClient.remark || '';

      // Add row to worksheet
      const newRow = worksheet.addRow(rowValues);

      // Set row height to 30 for all data rows (row 6 and below)
      newRow.height = 30;

      // Apply styles to each cell
      for (let c = 1; c <= remarkColIndex; c++) {
        const cell = newRow.getCell(c);

        // Set base styles
        cell.font = { name: 'Times New Roman', size: 11, bold: true };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' },
          bottom: { style: 'thin' }, right: { style: 'thin' }
        };

        // Set alternating row colors
        if (i % 2 === 0) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
        }

        // Apply background colors based on mark
        if (c >= 3 && c < remarkColIndex) {
          if (rowValues[c] === 'X') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Yellow
          } else if (rowValues[c] === '✓') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '92D050' } }; // Green
          }
        }

        // Set alignment
        if (c === 1) { // Sr. No
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = {
            top: { style: 'medium' }, left: { style: 'medium' },
            bottom: { style: 'medium' }, right: { style: 'medium' }
          };
        } else if (c === 2 || c === remarkColIndex) { // Client Name and Remark
          // --- UPDATED: wrapText for Client Name ---
          cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: c === 2 };
        } else { // All other cells
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
      }
    }

    // Freeze header row (now at row 5)
    worksheet.views = [{ state: 'frozen', ySplit: 5 }];

    // -- Add thick borders around stages and table (excluding empty columns) --

    // Total rows in sheet (including header etc)
    const lastDataRow = worksheet.lastRow.number;

    // Columns count
    const lastColNum = headers.length;

    // Border style for thick border
    const thickBorder = { style: 'medium', color: { argb: 'FF000000' } };

    // Function to set thick border around a range
    const setThickBorder = (startRow, endRow, startCol, endCol) => {
      for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
        const row = worksheet.getRow(rowNum);
        for (let colNum = startCol; colNum <= endCol; colNum++) {
          const cell = row.getCell(colNum);

          const border = {};

          if (rowNum === startRow) border.top = thickBorder;
          if (rowNum === endRow) border.bottom = thickBorder;
          if (colNum === startCol) border.left = thickBorder;
          if (colNum === endCol) border.right = thickBorder;

          cell.border = { ...cell.border, ...border };
        }
      }
    };

    // Table range starts from header row (4) to last data row
    const tableStartRow = 4;
    const tableEndRow = lastDataRow;

    // Full table border
    setThickBorder(tableStartRow, tableEndRow, 1, lastColNum);

    // Stage borders excluding empty separator columns
    setThickBorder(tableStartRow, tableEndRow, stage1Start, stage1End);
    setThickBorder(tableStartRow, tableEndRow, stage2Start, stage2End);
    setThickBorder(tableStartRow, tableEndRow, stage3Start, stage3End);
    setThickBorder(tableStartRow, tableEndRow, stageBillingStart, stageBillingEnd);
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

    const formattedMonth = getFullMonthName(month);
    const formattedYear = year;

    // Build the sanitized filename
    const fileName = `Client Progress Report - ${sanitize(formattedMonth)}-${formattedYear}.xlsx`;

    // Set response headers and send file
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fileName)}"`);
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    console.error('Error generating client progress Excel:', error);
    res.status(500).json({ message: 'Error generating client progress Excel', error: error.message });
  }
};

// Helper to safely get clientId string from various formats
const getClientIdString = (clientId) => {
  if (!clientId) return null;
  if (typeof clientId === 'string') return clientId;
  if (clientId._id) return clientId._id.toString();
  if (clientId.toString) return clientId.toString();
  return null;
};

// Helper: find client data in array by clientId string
const findClientData = (clientsArray, clientIdStr) => {
  return clientsArray.find(c => {
    const cId = getClientIdString(c.clientId);
    const match = cId === clientIdStr;
    if (!match) {
      console.log(`No match: searching for ${clientIdStr}, found clientId ${cId}`);
    } else {
      console.log(`Matched clientId: ${clientIdStr}`);
    }
    return match;
  });
};