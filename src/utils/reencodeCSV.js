const fs = require('fs');
const stripBomBuffer = require('strip-bom-buffer');

/**
 * Re-encode a CSV file to UTF-8 without BOM (in-place)
 * @param {string} filePath - Absolute path to the CSV file
 * @returns {Promise<void>}
 */
const reencodeCSVtoUTF8 = (filePath) => {
  return new Promise((resolve, reject) => {
    fs.readFile(filePath, (readErr, data) => {
      if (readErr) return reject(readErr);

      const cleanBuffer = stripBomBuffer(data);
      fs.writeFile(filePath, cleanBuffer, 'utf8', (writeErr) => {
        if (writeErr) return reject(writeErr);
        resolve();
      });
    });
  });
};

module.exports = { reencodeCSVtoUTF8 };
