const multer = require('multer');
const path = require('path');
const xlsx = require('xlsx');
const FormatDataF1 = require('../../models/v2/formatDataF1.model');
const FormatDataF2 = require('../../models/v2/formatDataF2.model');
const FormatDataF3 = require('../../models/v2/formatDataF3.model');
const FormatDataF4 = require('../../models/v2/formatDataF4.model');
const FormatDataF5 = require('../../models/v2/formatDataF5.model');
const fs = require('fs');
const { parse } = require('json2csv');

// Helper function to convert Excel date serial to DD/MM/YYYY HH:MM format (24-hour)
const excelDateToFormattedString = (excelDate) => {
  if (!excelDate) return '';
  
  // If already a string, check if it has AM/PM and convert to 24-hour
  if (typeof excelDate === 'string') {
    // Check if it has AM/PM
    if (excelDate.includes('AM') || excelDate.includes('PM')) {
      const parts = excelDate.match(/(\d{2})[\/\-](\d{2})[\/\-](\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s+(AM|PM)/i);
      if (parts) {
        let [, day, month, year, hours, minutes, seconds, period] = parts;
        hours = parseInt(hours);
        
        // Convert to 24-hour format
        if (period.toUpperCase() === 'PM' && hours !== 12) {
          hours += 12;
        } else if (period.toUpperCase() === 'AM' && hours === 12) {
          hours = 0;
        }
        
        return `${day}/${month}/${year} ${String(hours).padStart(2, '0')}:${minutes}`;
      }
    }
    return excelDate;
  }
  
  // If it's an Excel serial number
  if (typeof excelDate === 'number') {
    const date = xlsx.SSF.parse_date_code(excelDate);
    if (date) {
      const day = String(date.d).padStart(2, '0');
      const month = String(date.m).padStart(2, '0');
      const year = date.y;
      const hours = String(date.H || 0).padStart(2, '0');
      const minutes = String(date.M || 0).padStart(2, '0');
      return `${day}/${month}/${year} ${hours}:${minutes}`;
    }
  }
  
  return excelDate.toString();
};

// Helper function to get days in month
const getDaysInMonth = (month, year) => {
  const monthIndex = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ].indexOf(month);
  
  if (monthIndex === -1) return 31; // Default
  
  return new Date(parseInt(year), monthIndex + 1, 0).getDate();
};

// Helper function to generate 15-minute intervals for a date
const generate15MinIntervals = (date, monthIndex, year) => {
  const intervals = [];
  // Format: DD/MM/YYYY (shorter format for Excel compatibility)
  const dateStr = `${String(date).padStart(2, '0')}/${String(monthIndex).padStart(2, '0')}/${year}`;
  
  for (let hour = 0; hour < 24; hour++) {
    for (let minute = 0; minute < 60; minute += 15) {
      const startTime = `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
      const endMinute = minute + 15;
      const endHour = endMinute >= 60 ? hour + 1 : hour;
      const endTime = `${String(endHour % 24).padStart(2, '0')}:${String(endMinute % 60).padStart(2, '0')}`;
      
      intervals.push({
        date: dateStr,
        intervalStart: startTime,
        intervalEnd: endTime
      });
    }
  }
  
  return intervals;
};

// Helper function to detect format type from file structure
const detectFormatType = (jsonData) => {
  if (!jsonData || jsonData.length === 0) return null;
  
  // Check F5 format (has specific headers in lines 1-7)
  if (jsonData[0] && jsonData[0][0] && 
      jsonData[0][0].toString().includes('Consumption') && 
      jsonData[0][0].toString().includes('Energy')) {
    return 'F5';
  }
  
  // Check F3 format (has metadata structure with "Meter Serial Number" in row 2)
  if (jsonData[1] && jsonData[1][0] && 
      jsonData[1][0].toString().toLowerCase().includes('meter serial')) {
    return 'F3';
  }
  
  // Check F4 format (has title with "Load Survey data of")
  if (jsonData[0] && jsonData[0][0] && 
      jsonData[0][0].toString().includes('Load Survey data of')) {
    return 'F4';
  }
  
  // Check F1/F2 format by column count and headers
  const headerRow = jsonData[0];
  if (headerRow && headerRow.length > 10) {
    // F2 has more columns (around 32)
    if (headerRow.length > 25) {
      return 'F2';
    }
    // F1 has fewer columns (around 17)
    return 'F1';
  }
  
  return 'F1'; // Default to F1
};

// Helper function to detect month and year from file data
const detectMonthYear = (jsonData, formatType) => {
  let detectedDate = null;
  
  try {
    if (formatType === 'F1' || formatType === 'F2') {
      // F1/F2: Check meterDataCaptureTimestamp in row 2, column 5
      if (jsonData[1] && jsonData[1][4]) {
        detectedDate = jsonData[1][4];
      }
    } else if (formatType === 'F3') {
      // F3: Check data entries starting from row 19
      if (jsonData[18] && jsonData[18][0]) {
        detectedDate = jsonData[18][0];
      }
    } else if (formatType === 'F4') {
      // F4: Check time field in row 4, column 2
      if (jsonData[3] && jsonData[3][1]) {
        detectedDate = jsonData[3][1];
      }
    } else if (formatType === 'F5') {
      // F5: Check date in row 9, column 1
      if (jsonData[8] && jsonData[8][0]) {
        detectedDate = jsonData[8][0];
      }
    }
    
    if (!detectedDate) return null;
    
    // Parse date
    let month, year;
    
    if (typeof detectedDate === 'number') {
      // Excel serial number
      const date = xlsx.SSF.parse_date_code(detectedDate);
      if (date) {
        month = date.m;
        year = date.y;
      }
    } else if (typeof detectedDate === 'string') {
      // String date - try different formats
      detectedDate = detectedDate.toString().trim();
      
      // Try DD/MM/YYYY or DD-MM-YYYY
      let match = detectedDate.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
      if (match) {
        month = parseInt(match[2]);
        year = parseInt(match[3]);
      } else {
        // Try YYYY-MM-DD or YYYY/MM/DD
        match = detectedDate.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
        if (match) {
          year = parseInt(match[1]);
          month = parseInt(match[2]);
        }
      }
    }
    
    if (month && year && month >= 1 && month <= 12) {
      const monthNames = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
      ];
      return {
        month: monthNames[month - 1],
        year: year.toString()
      };
    }
  } catch (error) {
    console.error('Error detecting month/year:', error);
  }
  
  return null;
};

// Configure multer
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    const dir = './uploads';
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir);
    }
    cb(null, dir);
  },
  filename: function (req, file, cb) {
    cb(null, Date.now() + path.extname(file.originalname));
  }
});
const upload = multer({ storage: storage }).single('file');

const processExcel = async (filePath, month, year, formatType) => {
  const workbook = xlsx.readFile(filePath);

  // Validation function to check if a sheet contains expected headers/data
  const validateSheet = (sheetName, type) => {
    if (!sheetName || !workbook.Sheets[sheetName]) return false;
    // We can add deeper header validation here if needed, 
    // but checking for sheet existence/name is a good start if they are named correctly.
    // Since users might upload a single sheet named 'Sheet1', we'll be slightly lenient 
    // but throw an error if we definitively can't process it.
    return true; 
  };

  const errors = [];

  // Process F1
  if (formatType === 'All' || formatType === 'F1') {
    const f1SheetName = workbook.SheetNames.find(s => s.toUpperCase().includes('F1')) || (formatType === 'F1' ? workbook.SheetNames[0] : null);
    const f1Sheet = f1SheetName ? workbook.Sheets[f1SheetName] : null;
    
    if (!f1Sheet) {
      if (formatType === 'F1') throw new Error("Could not find F1 data in the uploaded file.");
      errors.push("F1 sheet missing");
    } else {
      const jsonData = xlsx.utils.sheet_to_json(f1Sheet, { header: 1 });
      const meterGroups = {};
      let hasData = false;

      // F1 Upload Format: Row 1 is header, data starts from Row 2
      // Columns: 0=SN, 1=Meter Serial Number, 2=Meter Timestamp, 3=Entry Timestamp, 
      //          4=Meter Data Capture Timestamp, 5=Block Frequency, 6=KWH-Import, etc.
      for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (row && row.length > 0) {
          const meterSerialNumber = row[1]?.toString();
          if (!meterSerialNumber) continue;
          hasData = true;

          if (!meterGroups[meterSerialNumber]) {
            meterGroups[meterSerialNumber] = [];
          }

          meterGroups[meterSerialNumber].push({
            meterTimestamp: excelDateToFormattedString(row[2]),
            entryTimestamp: excelDateToFormattedString(row[3]),
            meterDataCaptureTimestamp: excelDateToFormattedString(row[4]),
            blockFrequency: row[5],
            kwhImport: row[6],
            blockEnergyKWhExp: row[7],
            blockEnergyKVArhQ1: row[8],
            blockEnergyKVArhQ2: row[9],
            blockEnergyKVArhQ3: row[10],
            blockEnergyKVArhQ4: row[11],
            kvahImport: row[12],
            blockEnergyKVahExp: row[13],
            netActiveEnergy: row[14],
            averagePhaseVoltages: row[15],
            averageLineCurrents: row[16]
          });
        }
      }
      
      if (!hasData && formatType === 'F1') throw new Error("F1 sheet is empty or invalid format.");

      for (const [meterSerialNumber, dataEntries] of Object.entries(meterGroups)) {
        await FormatDataF1.deleteMany({ month, year, meterSerialNumber });
        await FormatDataF1.create({ month, year, meterSerialNumber, dataEntries });
      }
    }
  }

  // Process F2
  if (formatType === 'All' || formatType === 'F2') {
    const f2SheetName = workbook.SheetNames.find(s => s.toUpperCase().includes('F2')) || (formatType === 'F2' ? workbook.SheetNames[0] : null);
    const f2Sheet = f2SheetName ? workbook.Sheets[f2SheetName] : null;
    
    if (!f2Sheet) {
      if (formatType === 'F2') throw new Error("Could not find F2 data in the uploaded file.");
      errors.push("F2 sheet missing");
    } else {
      const jsonData = xlsx.utils.sheet_to_json(f2Sheet, { header: 1 });
      const meterGroups = {};
      let hasData = false;

      for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (row && row.length > 0) {
          const meterSerialNumber = row[1]?.toString();
          if (!meterSerialNumber) continue;
          hasData = true;

          if (!meterGroups[meterSerialNumber]) {
            meterGroups[meterSerialNumber] = [];
          }

          meterGroups[meterSerialNumber].push({
            meterTimestamp: excelDateToFormattedString(row[2]),
            entryTimestamp: excelDateToFormattedString(row[3]),
            meterDataCaptureTimestamp: excelDateToFormattedString(row[4]),
            kwhImport: row[5],
            kvahImport: row[6],
            fundamentalEnergyImportActiveEnergy: row[7],
            blockFrequency: row[8],
            avgVoltageVRN: row[9],
            avgVoltageVYN: row[10],
            avgVoltageVBN: row[11],
            netActiveEnergy: row[12],
            blockEnergyKWhExp: row[13],
            blockEnergyKVArhQ1: row[14],
            blockEnergyKVArhQ2: row[15],
            blockEnergyKVArhQ3: row[16],
            blockEnergyKVArhQ4: row[17],
            blockEnergyKVahExp: row[18],
            avgCurrentIR: row[19],
            avgCurrentIY: row[20],
            avgCurrentIB: row[21],
            reactiveEnergyHigh: row[22],
            reactiveEnergyLow: row[23],
            exportActiveEnergy: row[24],
            totalPowerFactor: row[25],
            exportPowerFactor: row[26],
            codedFrequency: row[27],
            averageVoltage: row[28],
            netKVARH: row[29],
            loadSurveyTamperSnap: row[30]?.toString(), // Ensure string
            billingAverageVoltage: row[31]
          });
        }
      }

      if (!hasData && formatType === 'F2') throw new Error("F2 sheet is empty or invalid format.");

      for (const [meterSerialNumber, dataEntries] of Object.entries(meterGroups)) {
        await FormatDataF2.deleteMany({ month, year, meterSerialNumber });
        await FormatDataF2.create({ month, year, meterSerialNumber, dataEntries });
      }
    }
  }

  // Process F3
  if (formatType === 'All' || formatType === 'F3') {
    const f3SheetName = workbook.SheetNames.find(s => s.toUpperCase().includes('F3')) || (formatType === 'F3' ? workbook.SheetNames[0] : null);
    const f3Sheet = f3SheetName ? workbook.Sheets[f3SheetName] : null;
    
    if (!f3Sheet) {
      if (formatType === 'F3') throw new Error("Could not find F3 data in the uploaded file.");
      errors.push("F3 sheet missing");
    } else {
      const jsonData = xlsx.utils.sheet_to_json(f3Sheet, { header: 1 });
      
      // F3 Upload Format: Metadata in rows 1-17, data header in row 18, data starts from row 19
      // Row 2, Column 1: Meter Serial Number
      const meterSerialNumber = jsonData[1]?.[1]?.toString();
      
      if (!meterSerialNumber && formatType === 'F3') {
         throw new Error("Invalid F3 format. Missing meter serial number at expected position (row 2, column B).");
      }
      
      if (meterSerialNumber) {
        // Extract metadata from rows 2-17
        const metadata = {
          dataDumpTime: jsonData[2]?.[1]?.toString(),
          dataReadTimeMRI: jsonData[3]?.[1]?.toString(),
          dataReadTimeMeter: jsonData[4]?.[1]?.toString(),
          meterType: jsonData[5]?.[1]?.toString(),
          location: jsonData[6]?.[1]?.toString(),
          logInterval: jsonData[7]?.[1],
          emf: jsonData[8]?.[1],
          emfApplied: jsonData[9]?.[1]?.toString(),
          installedCTRatio: jsonData[10]?.[1],
          installedPTRatio: jsonData[11]?.[1],
          commissionCTRatio: jsonData[12]?.[1],
          commissionPTRatio: jsonData[13]?.[1],
          lsDays: jsonData[14]?.[1],
          startDate: jsonData[15]?.[1]?.toString(),
          endDate: jsonData[16]?.[1]?.toString()
        };

        const dataEntries = [];
        // Data starts at row 19 (index 18)
        for (let i = 18; i < jsonData.length; i++) {
          const row = jsonData[i];
          if (row && row.length > 0 && row[0]) {
            // Convert date from Excel serial to DD/MM/YYYY format
            let dateStr = row[0];
            if (typeof dateStr === 'number') {
              const date = xlsx.SSF.parse_date_code(dateStr);
              if (date) {
                dateStr = `${String(date.d).padStart(2, '0')}/${String(date.m).padStart(2, '0')}/${date.y}`;
              }
            } else if (dateStr) {
              dateStr = dateStr.toString();
              // Ensure DD/MM/YYYY format
              if (dateStr.includes('-')) {
                dateStr = dateStr.replace(/-/g, '/');
              }
            }
            
            // Convert time values from Excel serial to HH:MM format
            let intervalStart = row[1];
            let intervalEnd = row[2];
            
            if (typeof intervalStart === 'number') {
              const hours = Math.floor(intervalStart * 24);
              const minutes = Math.floor((intervalStart * 24 * 60) % 60);
              intervalStart = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
            } else if (intervalStart) {
              intervalStart = intervalStart.toString();
            }
            
            if (typeof intervalEnd === 'number') {
              const hours = Math.floor(intervalEnd * 24);
              const minutes = Math.floor((intervalEnd * 24 * 60) % 60);
              intervalEnd = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
            } else if (intervalEnd) {
              intervalEnd = intervalEnd.toString();
            }
            
            dataEntries.push({
              date: dateStr,
              intervalStart: intervalStart,
              intervalEnd: intervalEnd,
              kwI: row[3],
              kwE: row[4],
              kvaI: row[5],
              kvaE: row[6],
              kwhI: row[7],
              kwhE: row[8],
              kwhNet: row[9],
              kvahI: row[10],
              kvahE: row[11]
            });
          }
        }
        
        if (dataEntries.length === 0 && formatType === 'F3') throw new Error("F3 sheet has no data entries.");

        await FormatDataF3.deleteMany({ month, year, meterSerialNumber });
        await FormatDataF3.create({ month, year, meterSerialNumber, ...metadata, dataEntries });
      }
    }
  }

  // Process F4
  if (formatType === 'All' || formatType === 'F4') {
    const f4SheetName = workbook.SheetNames.find(s => s.toUpperCase().includes('F4')) || (formatType === 'F4' ? workbook.SheetNames[0] : null);
    const f4Sheet = f4SheetName ? workbook.Sheets[f4SheetName] : null;
    
    if (!f4Sheet) {
      if (formatType === 'F4') throw new Error("Could not find F4 data in the uploaded file.");
      errors.push("F4 sheet missing");
    } else {
      const jsonData = xlsx.utils.sheet_to_json(f4Sheet, { header: 1 });
      
      // F4 Upload Format: Row 1 has title, Row 3 has headers, data starts from row 4
      // Title format: "Load Survey data of DG0997B from 2025/11/01 to 2025/12/01"
      const titleStr = jsonData[0]?.[0]?.toString() || '';
      const match = titleStr.match(/data of ([a-zA-Z0-9_-]+) from/i);
      const meterSerialNumber = match ? match[1] : 'UNKNOWN_METER';

      if (!match && formatType === 'F4') {
         throw new Error("Invalid F4 format. Missing title with meter serial number at row 1.");
      }

      const dataEntries = [];
      // Data starts at row 4 (index 3)
      for (let i = 3; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (row && row.length > 1 && row[1]) {
          // Column 1 has timestamp in format: "01/11/2025 00:15:00"
          let timeStr = row[1];
          if (typeof timeStr === 'number') {
            // Excel serial number to datetime
            timeStr = excelDateToFormattedString(timeStr);
          } else if (timeStr) {
            timeStr = timeStr.toString().trim();
            // Ensure DD/MM/YYYY HH:MM format (remove seconds if present)
            if (timeStr.includes(':')) {
              // Try to match with seconds: DD/MM/YYYY HH:MM:SS
              let parts = timeStr.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?/);
              if (parts) {
                timeStr = `${parts[1]}/${parts[2]}/${parts[3]} ${parts[4]}:${parts[5]}`;
              }
            }
          }
          
          dataEntries.push({
            time: timeStr,
            kwhImport: row[2],
            kwhExport: row[3],
            kvahImport: row[4],
            kvahExport: row[5],
            q1: row[6],
            q2: row[7],
            q3: row[8],
            q4: row[9],
            volR: row[10]?.toString() || '',
            volY: row[11]?.toString() || '',
            volB: row[12]?.toString() || '',
            curR: row[13]?.toString() || '',
            curY: row[14]?.toString() || '',
            curB: row[15]?.toString() || '',
            frequency: row[16],
            pfImport: row[17]?.toString() || '',
            pfExport: row[18]?.toString() || ''
          });
        }
      }
      
      if (dataEntries.length === 0 && formatType === 'F4') throw new Error("F4 sheet has no data entries.");

      await FormatDataF4.deleteMany({ month, year, meterSerialNumber });
      await FormatDataF4.create({ month, year, meterSerialNumber, title: titleStr, dataEntries });
    }
  }

  // Process F5 (CSV format)
  if (formatType === 'All' || formatType === 'F5') {
    const f5SheetName = workbook.SheetNames.find(s => s.toUpperCase().includes('F5')) || (formatType === 'F5' ? workbook.SheetNames[0] : null);
    const f5Sheet = f5SheetName ? workbook.Sheets[f5SheetName] : null;
    
    if (!f5Sheet) {
      if (formatType === 'F5') throw new Error("Could not find F5 data in the uploaded file.");
      errors.push("F5 sheet missing");
    } else {
      const jsonData = xlsx.utils.sheet_to_json(f5Sheet, { header: 1 });
      
      // F5 Format: Lines 1-7 are headers, Line 8 has column headers, data starts from line 9
      // Extract meter serial number from filename or use a default
      const meterSerialNumber = 'F5_METER'; // You can extract this from filename if needed
      
      // Get month index for filtering (1-12)
      const targetMonthIndex = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
      ].indexOf(month) + 1;
      
      const dataEntries = [];
      // Data starts at row 9 (index 8)
      for (let i = 8; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (row && row.length > 0 && row[0]) {
          // Handle date conversion - can be Excel serial number or DD/MM/YYYY string
          let dateStr = row[0];
          let parsedMonth = null;
          let parsedYear = null;
          
          if (typeof dateStr === 'number') {
            // Excel serial number to DD/MM/YYYY
            const date = xlsx.SSF.parse_date_code(dateStr);
            if (date) {
              dateStr = `${String(date.d).padStart(2, '0')}/${String(date.m).padStart(2, '0')}/${date.y}`;
              parsedMonth = date.m;
              parsedYear = date.y;
            }
          } else if (dateStr) {
            // Already a string, ensure DD/MM/YYYY format
            dateStr = dateStr.toString().trim();
            
            // Check if it's in YYYY-MM-DD or YYYY/MM/DD format
            const ymdMatch = dateStr.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
            if (ymdMatch) {
              // Convert YYYY-MM-DD to DD/MM/YYYY
              const [, yearVal, monthVal, dayVal] = ymdMatch;
              dateStr = `${String(dayVal).padStart(2, '0')}/${String(monthVal).padStart(2, '0')}/${yearVal}`;
              parsedMonth = parseInt(monthVal);
              parsedYear = parseInt(yearVal);
            } else {
              // If it's in different format with dashes, normalize to slashes
              if (dateStr.includes('-')) {
                // Convert DD-MM-YYYY to DD/MM/YYYY
                dateStr = dateStr.replace(/-/g, '/');
              }
              // Ensure DD/MM/YYYY format with proper padding
              const dmyMatch = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
              if (dmyMatch) {
                const [, dayVal, monthVal, yearVal] = dmyMatch;
                dateStr = `${String(dayVal).padStart(2, '0')}/${String(monthVal).padStart(2, '0')}/${yearVal}`;
                parsedMonth = parseInt(monthVal);
                parsedYear = parseInt(yearVal);
              }
            }
          }
          
          // Filter: Only include data for the selected month and year
          if (parsedMonth !== targetMonthIndex || parsedYear !== parseInt(year)) {
            continue; // Skip this row if it doesn't match the selected month/year
          }
          
          dataEntries.push({
            date: dateStr || '',
            intervalStart: row[1]?.toString() || '',
            intervalEnd: row[2]?.toString() || '',
            kwImp: row[3] || 0,
            kwExp: row[4] || 0,
            kvaImp: row[5] || 0,
            kvaExp: row[6] || 0,
            kwhImp: row[7] || 0,
            kwhExp: row[8] || 0,
            netKwh: row[9] || 0,
            kvahImp: row[10] || 0,
            kvahExp: row[11] || 0,
            kvarhLgDurWhImp: row[12] || 0,
            kvarhLdDurWhImp: row[13] || 0,
            kvarhLgDurWhExp: row[14] || 0,
            kvarhLdDurWhExp: row[15] || 0,
            avgRPhVolt: row[16] || 0,
            avgYPhVolt: row[17] || 0,
            avgBPhVolt: row[18] || 0,
            averageVolt: row[19] || 0,
            avgRPhAmp: row[20] || 0,
            avgYPhAmp: row[21] || 0,
            avgBPhAmp: row[22] || 0,
            averageAmp: row[23] || 0,
            codedFreq: row[24]?.toString() || '',
            avgFreq: row[25] || 0,
            powerOffMinutes: row[26] || 0,
            kvarhHighImport: row[27] || 0,
            kvarhHighExport: row[28] || 0,
            kvarhLowImport: row[29] || 0,
            kvarhLowExport: row[30] || 0,
            kvarhNetReacHigh: row[31] || 0,
            kvarhNetReacLow: row[32] || 0,
            powerFactorImp: row[33]?.toString() || '',
            powerFactorExp: row[34] || 0
          });
        }
      }
      
      if (dataEntries.length === 0 && formatType === 'F5') throw new Error("F5 sheet has no data entries.");

      await FormatDataF5.deleteMany({ month, year, meterSerialNumber });
      await FormatDataF5.create({ month, year, meterSerialNumber, dataEntries });
    }
  }

  if (formatType === 'All' && errors.length === 5) {
     throw new Error("Uploaded file does not contain any valid F1, F2, F3, F4, or F5 sheets.");
  }

  // Clean up
  fs.unlinkSync(filePath);
};

exports.uploadFormatData = (req, res) => {
  upload(req, res, async (err) => {
    if (err) {
      return res.status(500).json({ success: false, message: 'File upload error', error: err.message });
    }

    if (!req.file) {
      return res.status(400).json({ success: false, message: 'No file uploaded' });
    }

    const { month, year, formatType = 'All' } = req.body;
    if (!month || !year) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ success: false, message: 'Month and Year are required' });
    }

    try {
      await processExcel(req.file.path, month, year, formatType);
      res.status(200).json({ success: true, message: `Format Data (${formatType}) successfully processed and saved.` });
    } catch (error) {
      console.error("Excel processing error:", error);
      if (fs.existsSync(req.file.path)) {
        fs.unlinkSync(req.file.path);
      }
      // Send the specific error message to the frontend
      res.status(400).json({ success: false, message: error.message || 'Error processing excel file' });
    }
  });
};

// Auto-detect upload - detects format, month, year from file
exports.uploadFormatDataAuto = (req, res) => {
  upload(req, res, async (err) => {
    if (err) {
      return res.status(500).json({ success: false, message: 'File upload error', error: err.message });
    }

    if (!req.file) {
      return res.status(400).json({ success: false, message: 'No file uploaded' });
    }

    try {
      // Read file to detect format and month/year
      const workbook = xlsx.readFile(req.file.path);
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = xlsx.utils.sheet_to_json(firstSheet, { header: 1 });
      
      // Detect format type
      const detectedFormat = detectFormatType(jsonData);
      if (!detectedFormat) {
        fs.unlinkSync(req.file.path);
        return res.status(400).json({ success: false, message: 'Could not detect file format' });
      }
      
      // Detect month and year
      const detectedMonthYear = detectMonthYear(jsonData, detectedFormat);
      if (!detectedMonthYear) {
        fs.unlinkSync(req.file.path);
        return res.status(400).json({ success: false, message: 'Could not detect month/year from file data' });
      }
      
      // Process the file with detected parameters
      await processExcel(req.file.path, detectedMonthYear.month, detectedMonthYear.year, detectedFormat);
      
      res.status(200).json({ 
        success: true, 
        message: `Format Data successfully processed and saved.`,
        data: {
          formatType: detectedFormat,
          month: detectedMonthYear.month,
          year: detectedMonthYear.year
        }
      });
    } catch (error) {
      console.error("Excel processing error:", error);
      if (fs.existsSync(req.file.path)) {
        fs.unlinkSync(req.file.path);
      }
      res.status(400).json({ success: false, message: error.message || 'Error processing excel file' });
    }
  });
};


// Download Format Data as CSV
exports.downloadFormatDataCSV = async (req, res) => {
  try {
    const { month, year, formatType, meterSerialNumber } = req.query;

    if (!month || !year || !formatType) {
      return res.status(400).json({ 
        success: false, 
        message: 'Month, Year, and Format Type are required' 
      });
    }

    let data;
    let Model;

    // Select appropriate model
    switch (formatType) {
      case 'F1':
        Model = FormatDataF1;
        break;
      case 'F2':
        Model = FormatDataF2;
        break;
      case 'F3':
        Model = FormatDataF3;
        break;
      case 'F4':
        Model = FormatDataF4;
        break;
      case 'F5':
        Model = FormatDataF5;
        break;
      default:
        return res.status(400).json({ 
          success: false, 
          message: 'Invalid format type. Must be F1, F2, F3, F4, or F5' 
        });
    }

    // Query data
    const query = { month, year };
    if (meterSerialNumber) {
      query.meterSerialNumber = meterSerialNumber;
    }

    data = await Model.findOne(query).lean();

    if (!data) {
      return res.status(404).json({ 
        success: false, 
        message: `No ${formatType} data found for ${month} ${year}` 
      });
    }

    // Get days in month
    const daysInMonth = getDaysInMonth(month, year);
    const monthIndex = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ].indexOf(month) + 1;

    // Prepare CSV data based on format type
    let csvData = [];
    let headers = [];
    let fixedLines = [];

    // F5 has its own unique format - download exactly as uploaded
    if (formatType === 'F5') {
      // F5 Fixed header lines (1-7)
      fixedLines = [
        ['Consumption (Energy) are in k (e.g. kWH, kVAh, kVArh)'],
        ['Demand (Power) are in k (e.g. kW, kVA, kVAr)'],
        ['Voltages are in V'],
        ['Currents are in A'],
        ['Dates are in dd/mm/yyyy'],
        ['Value ***.** indicates meter was power off during complete interval'],
        ['']
      ];

      // F5 uses actual data entries directly (no interval generation needed)
      // Filter data entries to only include the selected month and year
      if (data.dataEntries && data.dataEntries.length > 0) {
        // Filter entries by month and year
        const filteredEntries = data.dataEntries.filter(entry => {
          if (!entry.date) return false;
          
          // Parse date to check month and year
          const dateMatch = entry.date.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
          if (dateMatch) {
            const [, day, entryMonth, entryYear] = dateMatch;
            return parseInt(entryMonth) === monthIndex && parseInt(entryYear) === parseInt(year);
          }
          return false;
        });
        
        csvData = filteredEntries.map(entry => ({
          'Date': entry.date || '',
          'Interval Start': entry.intervalStart || '',
          'Interval End': entry.intervalEnd || '',
          'KW Imp.': entry.kwImp || 0,
          'KW Exp.': entry.kwExp || 0,
          'KVA Imp.': entry.kvaImp || 0,
          'KVA Exp.': entry.kvaExp || 0,
          'KWh Imp. [1.0.1.29.0.255]': entry.kwhImp || 0,
          'KWh Exp. [1.0.2.29.0.255]': entry.kwhExp || 0,
          'Net KWh [1.0.16.29.0.255]': entry.netKwh || 0,
          'KVAh Imp. [1.0.9.29.0.255]': entry.kvahImp || 0,
          'KVAh Exp. [1.0.10.29.0.255]': entry.kvahExp || 0,
          'Kvarh Lg Dur Wh Imp. [1.0.5.29.0.255]': entry.kvarhLgDurWhImp || 0,
          'Kvarh Ld Dur Wh Imp. [1.0.8.29.0.255]': entry.kvarhLdDurWhImp || 0,
          'Kvarh Lg Dur Wh Exp. [1.0.7.29.0.255]': entry.kvarhLgDurWhExp || 0,
          'Kvarh Ld Dur Wh Exp. [1.0.6.29.0.255]': entry.kvarhLdDurWhExp || 0,
          'Avg. R-Ph  Volt [1.0.32.27.0.255]': entry.avgRPhVolt || 0,
          'Avg. Y-Ph  Volt [1.0.52.27.0.255]': entry.avgYPhVolt || 0,
          'Avg. B-Ph  Volt [1.0.72.27.0.255]': entry.avgBPhVolt || 0,
          'Average Volt': entry.averageVolt || 0,
          'Avg. R-Ph  Amp. [1.0.31.27.0.255]': entry.avgRPhAmp || 0,
          'Avg. Y-Ph  Amp. [1.0.51.27.0.255]': entry.avgYPhAmp || 0,
          'Avg. B-Ph  Amp. [1.0.71.27.0.255]': entry.avgBPhAmp || 0,
          'Average Amp.': entry.averageAmp || 0,
          'Coded Freq. [1.0.207.29.0.255]': entry.codedFreq || '',
          'Avg. Freq. [1.0.14.27.0.255](Hz.)': entry.avgFreq || 0,
          'power off minutes of last IP [1.0.139.29.0.255] (Min)': entry.powerOffMinutes || 0,
          'Kvarh High Import [1.0.146.29.0.255]': entry.kvarhHighImport || 0,
          'Kvarh High Export [1.0.147.29.0.255]': entry.kvarhHighExport || 0,
          'Kvarh Low Import [1.0.148.29.0.255]': entry.kvarhLowImport || 0,
          'Kvarh Low Export [1.0.149.29.0.255]': entry.kvarhLowExport || 0,
          'Kvarh Net Reac.High(Q1 + Q2 - Q3 - Q4)[1.0.196.8.0.255]': entry.kvarhNetReacHigh || 0,
          'Kvarh Net Reac.Low(Q1 + Q2 - Q3 - Q4)[1.0.197.8.0.255]': entry.kvarhNetReacLow || 0,
          'Power Factor Imp.': entry.powerFactorImp || '',
          'Power Factor Exp.': entry.powerFactorExp || 0
        }));
        
        // console.log(`\nF5: Downloaded ${csvData.length} data entries for ${month} ${year}\n`);
      }
      
    } else {
    // F1, F2, F3, F4 - UNIFIED DOWNLOAD FORMAT FOR ALL FORMATS
    // All formats download in the same structure as "Format" sheet
    // Only HIGHLIGHTED columns will have data, rest will be empty ("")
    
    // Fixed header lines (Rows 1-7) - Same for all formats
    fixedLines = [
      ['Energy parameters are in k (e.g. kWH, kVAh, kVArh)'],
      ['Demand (Power) are in k (e.g. kW, kVA, kVAr)'],
      ['Voltages are in V'],
      ['Currents are in A'],
      ["Dates are in 'dd/mm/yyyy' format"],
      ['Value ***.** indicates either meter was power off or data not available during complete interval'],
      [' ']
    ];

    // For F4, use uploaded intervals directly (no generation)
    // For F1, F2, F3, generate all 15-minute intervals for the month
    let allIntervals = [];
    
    if (formatType === 'F4') {
      // F4: Use only the intervals from uploaded data
      if (data.dataEntries && data.dataEntries.length > 0) {
        allIntervals = data.dataEntries.map(entry => {
          // Parse time field: DD/MM/YYYY HH:MM
          const match = entry.time.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);
          if (match) {
            const [, day, monthNum, yearNum, hours, minutes] = match;
            const startMinutes = parseInt(minutes);
            const endMinutes = startMinutes + 15;
            const endHours = endMinutes >= 60 ? parseInt(hours) + 1 : parseInt(hours);
            
            return {
              date: `${day}/${monthNum}/${yearNum}`,
              intervalStart: `${hours}:${minutes}`,
              intervalEnd: `${String(endHours % 24).padStart(2, '0')}:${String(endMinutes % 60).padStart(2, '0')}`,
              entry: entry // Store the entry directly
            };
          }
          return null;
        }).filter(Boolean);
        
        // console.log(`\nF4: Using ${allIntervals.length} intervals from uploaded data\n`);
      }
    } else {
      // F1, F2, F3: Generate all 15-minute intervals for the month
      for (let day = 1; day <= daysInMonth; day++) {
        const dayIntervals = generate15MinIntervals(day, monthIndex, year);
        allIntervals.push(...dayIntervals);
      }
    }

    // Map data entries to intervals based on format type
    const dataMap = {};
    let matchCount = 0;
    
    if (formatType === 'F1') {
      // F1: Use meterDataCaptureTimestamp
      if (data.dataEntries && data.dataEntries.length > 0) {
        data.dataEntries.forEach((entry) => {
          let timestamp = entry.meterDataCaptureTimestamp;
          if (timestamp) {
            const match = timestamp.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
            if (match) {
              let [, day, monthNum, yearNum, hours, minutes] = match;
              day = parseInt(day);
              monthNum = parseInt(monthNum);
              hours = parseInt(hours);
              minutes = Math.floor(parseInt(minutes) / 15) * 15;
              
              if (monthNum === monthIndex) {
                const key = `${String(day).padStart(2, '0')}/${String(monthNum).padStart(2, '0')}/${yearNum} ${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
                dataMap[key] = entry;
              }
            }
          }
        });
      }
    } else if (formatType === 'F2') {
      // F2: Use meterDataCaptureTimestamp
      if (data.dataEntries && data.dataEntries.length > 0) {
        data.dataEntries.forEach((entry) => {
          let timestamp = entry.meterDataCaptureTimestamp;
          if (timestamp) {
            const match = timestamp.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
            if (match) {
              let [, day, monthNum, yearNum, hours, minutes] = match;
              day = parseInt(day);
              monthNum = parseInt(monthNum);
              hours = parseInt(hours);
              minutes = Math.floor(parseInt(minutes) / 15) * 15;
              
              if (monthNum === monthIndex) {
                const key = `${String(day).padStart(2, '0')}/${String(monthNum).padStart(2, '0')}/${yearNum} ${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
                dataMap[key] = entry;
              }
            }
          }
        });
      }
    } else if (formatType === 'F3') {
      // F3: Use date + intervalStart
      if (data.dataEntries && data.dataEntries.length > 0) {
        data.dataEntries.forEach((entry) => {
          if (entry.date && entry.intervalStart) {
            // Normalize date to DD/MM/YYYY
            let dateStr = entry.date.toString().replace(/-/g, '/');
            const key = `${dateStr} ${entry.intervalStart}`;
            dataMap[key] = entry;
          }
        });
      }
    } else if (formatType === 'F4') {
      // F4: Data is already in allIntervals with entry attached, no need for dataMap
      // console.log(`\nF4 Debug: Using ${allIntervals.length} intervals from uploaded data`);
      // if (allIntervals.length > 0) {
      //   console.log('Sample interval:', allIntervals[0]);
      // }
    }

    // Generate CSV rows with UNIFIED column structure
    // Only HIGHLIGHTED columns get data, rest remain empty ("")
    csvData = allIntervals.map(interval => {
      let entry;
      
      if (formatType === 'F4') {
        // F4: Entry is already attached to interval
        entry = interval.entry;
        if (entry) matchCount++;
      } else {
        // F1, F2, F3: Look up entry in dataMap
        const key = `${interval.date} ${interval.intervalStart}`;
        entry = dataMap[key];
        if (entry) matchCount++;
      }

      // Map data to columns based on format type
      // HIGHLIGHTED COLUMNS for F1: Active(I), Active(E), Net Active
      let activeI = 0, activeE = 0, netActive = 0;

      if (entry) {
        if (formatType === 'F1') {
          activeI = entry.kwhImport || 0;
          activeE = entry.blockEnergyKWhExp || 0;
          netActive = entry.netActiveEnergy || 0;
        } else if (formatType === 'F2') {
          activeI = entry.kwhImport || 0;
          activeE = entry.blockEnergyKWhExp || 0;
          netActive = entry.netActiveEnergy || 0;
        } else if (formatType === 'F3') {
          activeI = entry.kwhI || 0;
          activeE = entry.kwhE || 0;
          netActive = entry.kwhNet || 0;
        } else if (formatType === 'F4') {
          activeI = entry.kwhImport || 0;
          activeE = entry.kwhExport || 0;
          netActive = activeI - activeE;
        }
      }

      // Return row with ONLY highlighted columns having data
      // Columns 5-10, 12-16 are empty ("")
      return {
        'Date': interval.date,
        'Interval Start': interval.intervalStart,
        'Interval End': interval.intervalEnd,
        'Active(I) Total': activeI,
        'Active(E) Total': activeE,
        'Reactive(I)-Active(I)': '',
        'Reactive(E)-Active(I)': '',
        'Reactive(I)-Active(E)': '',
        'Reactive(E)-Active(E)': '',
        'Apparent-Active(I) - type 2': '',
        'Apparent-Active(E) - type 6': '',
        'Net Active': netActive,
        'Average Voltage': '',
        'Average Current': '',
        'Frequency': '',
        'Calculated Avg Imp Power Factor (signed)': '',
        'Calculated Avg Exp Power Factor (signed)': ''
      };
    });
    
    // console.log(`\n${formatType}: Matched ${matchCount} out of ${allIntervals.length} intervals`);
    // console.log('Match percentage:', ((matchCount / allIntervals.length) * 100).toFixed(2) + '%\n');
    } // End of F1-F4 unified format

    // Generate CSV content manually (more reliable than json2csv)
    let csvLines = [];
    
    // Check if csvData is empty
    if (!csvData || csvData.length === 0) {
      return res.status(404).json({ 
        success: false, 
        message: `No data found for ${formatType} in ${month} ${year}` 
      });
    }
    
    // Add fixed header lines
    fixedLines.forEach(line => {
      csvLines.push(line.join(','));
    });
    
    // Add column headers
    const columnHeaders = Object.keys(csvData[0]);
    csvLines.push(columnHeaders.join(','));
    
    // Add data rows
    csvData.forEach(row => {
      const values = columnHeaders.map(header => {
        const value = row[header];
        // Escape commas and quotes in values
        if (typeof value === 'string' && (value.includes(',') || value.includes('"'))) {
          return `"${value.replace(/"/g, '""')}"`;
        }
        return value;
      });
      csvLines.push(values.join(','));
    });
    
    const finalCSV = csvLines.join('\n');

    // Generate filename in format: Load Survey - {MeterNumber} - {StartDate} to {EndDate} - Logger1.csv
    // Date format: DD-MM-YY
    const startDate = `01-${String(monthIndex).padStart(2, '0')}-${year.toString().slice(-2)}`;
    const endDate = `${String(daysInMonth).padStart(2, '0')}-${String(monthIndex).padStart(2, '0')}-${year.toString().slice(-2)}`;
    
    const filename = `Load Survey - ${data.meterSerialNumber} - ${startDate} to ${endDate} - Logger1.csv`;
    
    // Set response headers
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

    res.status(200).send(finalCSV);

  } catch (error) {
    console.error("Download error:", error);
    res.status(500).json({ 
      success: false, 
      message: error.message || 'Error generating CSV file' 
    });
  }
};

// Get all uploaded format data (list of all month/year/format combinations)
exports.getAllUploadedFormats = async (req, res) => {
  try {
    const allFormats = [];

    // Get all F1 data
    const f1Data = await FormatDataF1.find({}).select('month year meterSerialNumber createdAt').lean();
    f1Data.forEach(item => {
      allFormats.push({
        formatType: 'F1',
        month: item.month,
        year: item.year,
        meterSerialNumber: item.meterSerialNumber,
        uploadedAt: item.createdAt
      });
    });

    // Get all F2 data
    const f2Data = await FormatDataF2.find({}).select('month year meterSerialNumber createdAt').lean();
    f2Data.forEach(item => {
      allFormats.push({
        formatType: 'F2',
        month: item.month,
        year: item.year,
        meterSerialNumber: item.meterSerialNumber,
        uploadedAt: item.createdAt
      });
    });

    // Get all F3 data
    const f3Data = await FormatDataF3.find({}).select('month year meterSerialNumber createdAt').lean();
    f3Data.forEach(item => {
      allFormats.push({
        formatType: 'F3',
        month: item.month,
        year: item.year,
        meterSerialNumber: item.meterSerialNumber,
        uploadedAt: item.createdAt
      });
    });

    // Get all F4 data
    const f4Data = await FormatDataF4.find({}).select('month year meterSerialNumber createdAt').lean();
    f4Data.forEach(item => {
      allFormats.push({
        formatType: 'F4',
        month: item.month,
        year: item.year,
        meterSerialNumber: item.meterSerialNumber,
        uploadedAt: item.createdAt
      });
    });

    // Get all F5 data
    const f5Data = await FormatDataF5.find({}).select('month year meterSerialNumber createdAt').lean();
    f5Data.forEach(item => {
      allFormats.push({
        formatType: 'F5',
        month: item.month,
        year: item.year,
        meterSerialNumber: item.meterSerialNumber,
        uploadedAt: item.createdAt
      });
    });

    // Sort by upload date (newest first)
    allFormats.sort((a, b) => new Date(b.uploadedAt) - new Date(a.uploadedAt));

    res.status(200).json({ 
      success: true, 
      data: allFormats 
    });

  } catch (error) {
    console.error("Error fetching uploaded formats:", error);
    res.status(500).json({ 
      success: false, 
      message: error.message || 'Error fetching uploaded formats' 
    });
  }
};


// Get available meters for a format type
exports.getAvailableMeters = async (req, res) => {
  try {
    const { month, year, formatType } = req.query;

    if (!month || !year || !formatType) {
      return res.status(400).json({ 
        success: false, 
        message: 'Month, Year, and Format Type are required' 
      });
    }

    let Model;
    switch (formatType) {
      case 'F1':
        Model = FormatDataF1;
        break;
      case 'F2':
        Model = FormatDataF2;
        break;
      case 'F3':
        Model = FormatDataF3;
        break;
      case 'F4':
        Model = FormatDataF4;
        break;
      case 'F5':
        Model = FormatDataF5;
        break;
      default:
        return res.status(400).json({ 
          success: false, 
          message: 'Invalid format type' 
        });
    }

    const meters = await Model.find({ month, year }).select('meterSerialNumber').lean();
    
    res.status(200).json({ 
      success: true, 
      meters: meters.map(m => m.meterSerialNumber) 
    });

  } catch (error) {
    console.error("Error fetching meters:", error);
    res.status(500).json({ 
      success: false, 
      message: error.message || 'Error fetching available meters' 
    });
  }
};


// Debug endpoint to check stored data format
exports.debugFormatData = async (req, res) => {
  try {
    const { month, year, formatType, meterSerialNumber } = req.query;

    if (!month || !year || !formatType) {
      return res.status(400).json({ 
        success: false, 
        message: 'Month, Year, and Format Type are required' 
      });
    }

    let Model;
    switch (formatType) {
      case 'F1':
        Model = FormatDataF1;
        break;
      case 'F2':
        Model = FormatDataF2;
        break;
      case 'F3':
        Model = FormatDataF3;
        break;
      case 'F4':
        Model = FormatDataF4;
        break;
      case 'F5':
        Model = FormatDataF5;
        break;
      default:
        return res.status(400).json({ 
          success: false, 
          message: 'Invalid format type' 
        });
    }

    const query = { month, year };
    if (meterSerialNumber) {
      query.meterSerialNumber = meterSerialNumber;
    }

    const data = await Model.findOne(query).lean();

    if (!data) {
      return res.status(404).json({ 
        success: false, 
        message: `No ${formatType} data found for ${month} ${year}` 
      });
    }

    // Return first 5 entries for debugging
    const sampleEntries = data.dataEntries ? data.dataEntries.slice(0, 5) : [];

    res.status(200).json({ 
      success: true, 
      meterSerialNumber: data.meterSerialNumber,
      totalEntries: data.dataEntries ? data.dataEntries.length : 0,
      sampleEntries: sampleEntries,
      firstTimestamp: sampleEntries[0]?.meterTimestamp || sampleEntries[0]?.entryTimestamp || 'N/A'
    });

  } catch (error) {
    console.error("Debug error:", error);
    res.status(500).json({ 
      success: false, 
      message: error.message || 'Error fetching debug data' 
    });
  }
};


// Delete format data
exports.deleteFormatData = async (req, res) => {
  try {
    const { month, year, formatType, meterSerialNumber } = req.query;

    if (!month || !year || !formatType) {
      return res.status(400).json({ 
        success: false, 
        message: 'Month, Year, and Format Type are required' 
      });
    }

    let Model;
    switch (formatType) {
      case 'F1':
        Model = FormatDataF1;
        break;
      case 'F2':
        Model = FormatDataF2;
        break;
      case 'F3':
        Model = FormatDataF3;
        break;
      case 'F4':
        Model = FormatDataF4;
        break;
      case 'F5':
        Model = FormatDataF5;
        break;
      default:
        return res.status(400).json({ 
          success: false, 
          message: 'Invalid format type' 
        });
    }

    const query = { month, year };
    if (meterSerialNumber) {
      query.meterSerialNumber = meterSerialNumber;
    }

    const result = await Model.deleteMany(query);

    res.status(200).json({ 
      success: true, 
      message: `Deleted ${result.deletedCount} record(s) for ${formatType} ${month} ${year}`,
      deletedCount: result.deletedCount
    });

  } catch (error) {
    console.error("Delete error:", error);
    res.status(500).json({ 
      success: false, 
      message: error.message || 'Error deleting data' 
    });
  }
};
