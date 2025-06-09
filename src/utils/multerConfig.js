const multer = require('multer');
const path = require('path');
const fs = require('fs');

// Check if upload directory exists, if not, create it
const uploadDir = path.join(__dirname, '../../uploads');
if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
}

// Set up multer storage engine
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, './uploads');  // Folder for saving uploaded files
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + '-' + file.originalname);  // Unique filename
    }
});

// File filter to ensure only CSV files are accepted
const fileFilter = (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext !== '.csv') {
        return cb(new Error('Only CSV files are allowed'), false);
    }
    cb(null, true);
};

// Configure Multer with storage, file size limit (e.g., 5MB), and file filter
const upload = multer({ storage });

module.exports = upload;
