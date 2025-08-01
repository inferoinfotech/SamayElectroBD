const fs = require('fs');
const readline = require('readline');

/**
 * Prepare CSV stream starting from the first valid header row (e.g. 'Date,...')
 * @param {string} filePath
 * @param {(line: string) => boolean} matchHeaderFn
 * @returns {Promise<fs.ReadStream>}
 */
const prepareCSVStream = async (filePath, matchHeaderFn) => {
  return new Promise((resolve, reject) => {
    const inputStream = fs.createReadStream(filePath);
    const rl = readline.createInterface({ input: inputStream });

    const tempFilePath = `${filePath}.tmp`;
    const tempStream = fs.createWriteStream(tempFilePath);

    let headerFound = false;

    rl.on('line', (line) => {
      if (headerFound || matchHeaderFn(line)) {
        headerFound = true;
        tempStream.write(line + '\n');
      }
    });

    rl.on('close', () => {
      tempStream.end();
      if (!headerFound) {
        reject(new Error('❌ No valid CSV header row found.'));
      } else {
        resolve(fs.createReadStream(tempFilePath));
      }
    });

    rl.on('error', reject);
  });
};

module.exports = { prepareCSVStream };
