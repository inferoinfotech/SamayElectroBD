const fs = require('fs');
const path = require('path');

const filePath = path.join(__dirname, '..', 'SamayElectroFD', 'Format5.CSV');

console.log('Analyzing Format5.CSV file\n');
console.log('='.repeat(60));

try {
  const content = fs.readFileSync(filePath, 'utf-8');
  const lines = content.split('\n');
  
  console.log(`\nTotal Lines: ${lines.length}`);
  console.log('\n--- First 30 Lines ---\n');
  
  lines.slice(0, 30).forEach((line, idx) => {
    console.log(`Line ${idx + 1}: ${line}`);
  });
  
  console.log('\n--- Analyzing Structure ---\n');
  
  // Check if it's comma-separated
  const firstDataLine = lines[0];
  const columns = firstDataLine.split(',');
  console.log(`Number of columns: ${columns.length}`);
  console.log('\nColumn Headers:');
  columns.forEach((col, idx) => {
    console.log(`  Column ${idx}: "${col.trim()}"`);
  });
  
  // Show sample data rows
  console.log('\n--- Sample Data Rows (2-5) ---\n');
  for (let i = 1; i < Math.min(5, lines.length); i++) {
    const cols = lines[i].split(',');
    console.log(`\nRow ${i + 1}:`);
    cols.forEach((col, idx) => {
      if (col.trim()) {
        console.log(`  Column ${idx}: "${col.trim()}"`);
      }
    });
  }
  
} catch (error) {
  console.error('Error reading file:', error.message);
}
