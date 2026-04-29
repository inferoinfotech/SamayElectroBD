const xlsx = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, '..', 'SamayElectroFD', 'RESULT FILE -H v4 2026-01-15 FINAL 4 Format (1) - Copy.xlsx');

console.log('Checking which columns are highlighted in upload formats\n');

try {
  const workbook = xlsx.readFile(filePath);
  
  // Check F1 Upload Format
  console.log('\n========== F1 UPLOAD - Checking Cell Styles ==========\n');
  const f1Sheet = workbook.Sheets['F1'];
  if (f1Sheet) {
    const f1Data = xlsx.utils.sheet_to_json(f1Sheet, { header: 1, defval: '' });
    
    console.log('F1 Header Row (Row 1):');
    if (f1Data[0]) {
      f1Data[0].forEach((cell, idx) => {
        if (cell) {
          // Check cell address
          const cellAddr = xlsx.utils.encode_cell({ r: 0, c: idx });
          const cellObj = f1Sheet[cellAddr];
          console.log(`Column ${idx} (${String.fromCharCode(65 + idx)}): "${cell}"`);
          if (cellObj && cellObj.s) {
            console.log(`  Style:`, JSON.stringify(cellObj.s));
          }
        }
      });
    }
    
    console.log('\n--- Based on typical format, likely highlighted columns for F1 ---');
    console.log('Columns that should be in download:');
    console.log('- Column 4: Meter Data Capture Timestamp');
    console.log('- Column 6: KWH-Import');
    console.log('- Column 7: Block Energy-KWh(Exp)');
    console.log('- Column 8: Block Energy-KVArh Q1');
    console.log('- Column 9: Block Energy-KVArh Q2');
    console.log('- Column 10: Block Energy-KVArh Q3');
    console.log('- Column 11: Block Energy-KVArh Q4');
    console.log('- Column 12: KVAH-Import');
    console.log('- Column 13: Block Energy-KVah(Exp)');
    console.log('- Column 14: Net Active Energy');
    console.log('- Column 15: Average Phase Voltages');
    console.log('- Column 16: Average - Line currents');
    console.log('- Column 5: Block Frequency(Hz)');
  }
  
  // Check Format sheet to see download structure
  console.log('\n\n========== FORMAT SHEET (Download Template) ==========\n');
  const formatSheet = workbook.Sheets['Format'];
  if (formatSheet) {
    const formatData = xlsx.utils.sheet_to_json(formatSheet, { header: 1, defval: '' });
    
    console.log('Download Column Headers (Row 8):');
    if (formatData[7]) {
      formatData[7].forEach((cell, idx) => {
        if (cell) {
          console.log(`Column ${idx}: "${cell}"`);
        }
      });
    }
    
    console.log('\n--- Sample Data Row (Row 9) ---');
    if (formatData[8]) {
      console.log('Row 9:', JSON.stringify(formatData[8].slice(0, 17)));
    }
  }
  
} catch (error) {
  console.error('Error:', error.message);
}
