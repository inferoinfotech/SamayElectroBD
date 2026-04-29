const xlsx = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, '..', 'SamayElectroFD', 'RESULT FILE -H v4 2026-01-15 FINAL 4 Format (1) - Copy.xlsx');

console.log('Analyzing Format sheet to identify highlighted vs non-highlighted columns\n');

try {
  const workbook = xlsx.readFile(filePath);
  const formatSheet = workbook.Sheets['Format'];
  
  if (formatSheet) {
    const formatData = xlsx.utils.sheet_to_json(formatSheet, { header: 1, defval: '' });
    
    console.log('=== DOWNLOAD FORMAT ANALYSIS ===\n');
    console.log('Row 8 (Headers):', formatData[7]);
    console.log('\n--- Checking which columns have data in sample rows ---\n');
    
    // Check rows 9-20 to see which columns typically have data
    const columnHasData = {};
    
    for (let rowIdx = 8; rowIdx < Math.min(20, formatData.length); rowIdx++) {
      const row = formatData[rowIdx];
      row.forEach((cell, colIdx) => {
        if (cell !== '' && cell !== null && cell !== undefined) {
          if (!columnHasData[colIdx]) {
            columnHasData[colIdx] = 0;
          }
          columnHasData[colIdx]++;
        }
      });
    }
    
    console.log('Columns with data (highlighted):');
    Object.keys(columnHasData).forEach(colIdx => {
      const header = formatData[7][colIdx];
      console.log(`  Column ${colIdx}: "${header}" - has data in ${columnHasData[colIdx]} rows`);
    });
    
    console.log('\n--- Sample rows to verify ---');
    for (let i = 8; i < 13; i++) {
      console.log(`\nRow ${i + 1}:`, formatData[i]);
    }
  }
  
} catch (error) {
  console.error('Error:', error.message);
}
