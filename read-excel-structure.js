const xlsx = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, '..', 'SamayElectroFD', 'RESULT FILE -H v4 2026-01-15 FINAL 4 Format (1) - Copy.xlsx');

console.log('Reading Excel file:', filePath);
console.log('\n===========================================\n');

try {
  const workbook = xlsx.readFile(filePath);
  
  console.log('Sheet Names:', workbook.SheetNames);
  console.log('\n===========================================\n');
  
  workbook.SheetNames.forEach((sheetName, index) => {
    console.log(`\n\n========== SHEET ${index + 1}: ${sheetName} ==========\n`);
    
    const sheet = workbook.Sheets[sheetName];
    const jsonData = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    
    console.log(`Total Rows: ${jsonData.length}`);
    console.log('\n--- First 20 Rows ---\n');
    
    jsonData.slice(0, 20).forEach((row, rowIndex) => {
      console.log(`Row ${rowIndex + 1}:`, JSON.stringify(row));
    });
    
    // Show column headers if they exist
    if (jsonData.length > 0) {
      console.log('\n--- Column Structure (First Row) ---');
      jsonData[0].forEach((cell, colIndex) => {
        const colLetter = String.fromCharCode(65 + colIndex);
        console.log(`Column ${colLetter} (${colIndex}): "${cell}"`);
      });
    }
    
    // Show a sample data row
    if (jsonData.length > 1) {
      console.log('\n--- Sample Data Row (Row 2) ---');
      jsonData[1].forEach((cell, colIndex) => {
        const colLetter = String.fromCharCode(65 + colIndex);
        console.log(`Column ${colLetter} (${colIndex}): "${cell}"`);
      });
    }
    
    console.log('\n===========================================');
  });
  
} catch (error) {
  console.error('Error reading Excel file:', error.message);
  console.error(error.stack);
}
