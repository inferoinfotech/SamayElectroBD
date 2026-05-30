const mongoose = require('mongoose');

const emailDirectorySchema = new mongoose.Schema(
  {
    key: { type: String, default: 'global', unique: true },
    entries: [
      {
        displayName: { type: String, required: true, trim: true },
        email: { type: String, required: true, trim: true, lowercase: true },
      },
    ],
    fileName: { type: String },
    uploadedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { timestamps: true }
);

module.exports = mongoose.model('EmailDirectory', emailDirectorySchema);
