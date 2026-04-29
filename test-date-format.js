// Test script to check Excel date conversion
const xlsx = require('xlsx');

// Helper function to convert Excel date
const excelDateToFormattedString = (excelDate) => {
  if (!excelDate) return '';
  
  // If already a string, return as is
  if (typeof excelDate === 'string') {
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
      return `${day}-${month}-${year} ${hours}:${minutes}`;
    }
  }
  
  return excelDate.toString();
};

// Test with sample Excel serial number
const excelSerial = 45747.0104166667; // This is 01-04-2025 00:15
console.log('Excel Serial:', excelSerial);
console.log('Converted:', excelDateToFormattedString(excelSerial));

// Test with string
const stringDate = '01-04-2025 00:15';
console.log('\nString Date:', stringDate);
console.log('Converted:', excelDateToFormattedString(stringDate));
