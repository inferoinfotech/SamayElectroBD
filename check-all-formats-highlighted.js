const xlsx = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, '..', 'SamayElectroFD', 'RESULT FILE -H v4 2026-01-15 FINAL 4 Format (1) - Copy.xlsx');

console.log('Checking highlighted columns for F2, F3, F4 upload formats\n');

try {
  const workbook = xlsx.readFile(filePath);
  
  // Check F2
  console.log('\n========== F2 UPLOAD FORMAT ==========\n');
  const f2Sheet = workbook.Sheets['F2'];
  if (f2Sheet) {
    const f2Data = xlsx.utils.sheet_to_json(f2Sheet, { header: 1, defval: '' });
    console.log('F2 Headers:');
    f2Data[0].forEach((cell, idx) => {
      if (cell) console.log(`  Column ${idx}: "${cell}"`);
    });
    console.log('\nSample data row:', f2Data[1].slice(0, 32));
  }
  
  // Check F3
  console.log('\n\n========== F3 UPLOAD FORMAT ==========\n');
  const f3Sheet = workbook.Sheets['F3'];
  if (f3Sheet) {
    const f3Data = xlsx.utils.sheet_to_json(f3Sheet, { header: 1, defval: '' });
    console.log('F3 Data Headers (Row 18):');
    if (f3Data[17]) {
      f3Data[17].forEach((cell, idx) => {
        if (cell) console.log(`  Column ${idx}: "${cell}"`);
      });
    }
    console.log('\nSample data row:', f3Data[18]);
  }
  
  // Check F4
  console.log('\n\n========== F4 UPLOAD FORMAT ==========\n');
  const f4Sheet = workbook.Sheets['F4'];
  if (f4Sheet) {
    const f4Data = xlsx.utils.sheet_to_json(f4Sheet, { header: 1, defval: '' });
    console.log('F4 Headers (Row 3):');
    if (f4Data[2]) {
      f4Data[2].forEach((cell, idx) => {
        if (cell) console.log(`  Column ${idx}: "${cell}"`);
      });
    }
    console.log('\nSample data row:', f4Data[3]);
  }
  
  // Now check what columns have data in Format sheet for comparison
  console.log('\n\n========== COMPARING WITH FORMAT SHEET ==========\n');
  const formatSheet = workbook.Sheets['Format'];
  if (formatSheet) {
    const formatData = xlsx.utils.sheet_to_json(formatSheet, { header: 1, defval: '' });
    
    console.log('Format sheet columns with data:');
    const sampleRow = formatData[8]; // Row 9
    sampleRow.forEach((cell, idx) => {
      const header = formatData[7][idx];
      const hasData = cell !== '' && cell !== null && cell !== undefined;
      console.log(`  Column ${idx}: "${header}" - ${hasData ? 'HAS DATA ✓' : 'EMPTY'}`);
    });
  }
  
} catch (error) {
  console.error('Error:', error.message);
}
