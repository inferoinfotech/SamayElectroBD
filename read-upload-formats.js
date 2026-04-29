const xlsx = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, '..', 'SamayElectroFD', 'RESULT FILE -H v4 2026-01-15 FINAL 4 Format (1) - Copy.xlsx');

console.log('Reading Upload Formats from Excel file\n');

try {
  const workbook = xlsx.readFile(filePath);
  
  // Read F1 Upload Format
  console.log('\n========== F1 UPLOAD FORMAT ==========\n');
  const f1Sheet = workbook.Sheets['F1'];
  if (f1Sheet) {
    const f1Data = xlsx.utils.sheet_to_json(f1Sheet, { header: 1, defval: '' });
    console.log('Total Rows:', f1Data.length);
    console.log('\n--- Header Row (Row 1) ---');
    if (f1Data[0]) {
      f1Data[0].forEach((cell, idx) => {
        if (cell) console.log(`Column ${idx}: "${cell}"`);
      });
    }
    console.log('\n--- Sample Data Rows (2-5) ---');
    f1Data.slice(1, 5).forEach((row, idx) => {
      console.log(`\nRow ${idx + 2}:`, JSON.stringify(row.slice(0, 20)));
    });
  }
  
  // Read F2 Upload Format
  console.log('\n\n========== F2 UPLOAD FORMAT ==========\n');
  const f2Sheet = workbook.Sheets['F2'];
  if (f2Sheet) {
    const f2Data = xlsx.utils.sheet_to_json(f2Sheet, { header: 1, defval: '' });
    console.log('Total Rows:', f2Data.length);
    console.log('\n--- Header Row (Row 1) ---');
    if (f2Data[0]) {
      f2Data[0].forEach((cell, idx) => {
        if (cell) console.log(`Column ${idx}: "${cell}"`);
      });
    }
    console.log('\n--- Sample Data Rows (2-5) ---');
    f2Data.slice(1, 5).forEach((row, idx) => {
      console.log(`\nRow ${idx + 2}:`, JSON.stringify(row.slice(0, 35)));
    });
  }
  
  // Read F3 Upload Format
  console.log('\n\n========== F3 UPLOAD FORMAT ==========\n');
  const f3Sheet = workbook.Sheets['F3'];
  if (f3Sheet) {
    const f3Data = xlsx.utils.sheet_to_json(f3Sheet, { header: 1, defval: '' });
    console.log('Total Rows:', f3Data.length);
    console.log('\n--- Metadata Section (Rows 1-17) ---');
    f3Data.slice(0, 17).forEach((row, idx) => {
      console.log(`Row ${idx + 1}:`, JSON.stringify(row.slice(0, 5)));
    });
    console.log('\n--- Data Header Row (Row 18) ---');
    if (f3Data[17]) {
      f3Data[17].forEach((cell, idx) => {
        if (cell) console.log(`Column ${idx}: "${cell}"`);
      });
    }
    console.log('\n--- Sample Data Rows (19-22) ---');
    f3Data.slice(18, 22).forEach((row, idx) => {
      console.log(`\nRow ${idx + 19}:`, JSON.stringify(row.slice(0, 15)));
    });
  }
  
  // Read F4 Upload Format
  console.log('\n\n========== F4 UPLOAD FORMAT ==========\n');
  const f4Sheet = workbook.Sheets['F4'];
  if (f4Sheet) {
    const f4Data = xlsx.utils.sheet_to_json(f4Sheet, { header: 1, defval: '' });
    console.log('Total Rows:', f4Data.length);
    console.log('\n--- Title Row (Row 1) ---');
    console.log('Row 1:', JSON.stringify(f4Data[0]));
    console.log('\n--- Header Row (Row 3) ---');
    if (f4Data[2]) {
      f4Data[2].forEach((cell, idx) => {
        if (cell) console.log(`Column ${idx}: "${cell}"`);
      });
    }
    console.log('\n--- Sample Data Rows (4-7) ---');
    f4Data.slice(3, 7).forEach((row, idx) => {
      console.log(`\nRow ${idx + 4}:`, JSON.stringify(row.slice(0, 20)));
    });
  }
  
  // Read Download Format
  console.log('\n\n========== DOWNLOAD FORMAT (Format Sheet) ==========\n');
  const formatSheet = workbook.Sheets['Format'];
  if (formatSheet) {
    const formatData = xlsx.utils.sheet_to_json(formatSheet, { header: 1, defval: '' });
    console.log('Total Rows:', formatData.length);
    console.log('\n--- Fixed Header Lines (Rows 1-7) ---');
    formatData.slice(0, 7).forEach((row, idx) => {
      console.log(`Row ${idx + 1}:`, JSON.stringify(row.slice(0, 5)));
    });
    console.log('\n--- Column Headers (Row 8) ---');
    if (formatData[7]) {
      formatData[7].forEach((cell, idx) => {
        if (cell) console.log(`Column ${idx}: "${cell}"`);
      });
    }
    console.log('\n--- Sample Data Rows (9-12) ---');
    formatData.slice(8, 12).forEach((row, idx) => {
      console.log(`\nRow ${idx + 9}:`, JSON.stringify(row.slice(0, 17)));
    });
  }
  
} catch (error) {
  console.error('Error:', error.message);
}
