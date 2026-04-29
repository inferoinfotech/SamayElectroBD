const mongoose = require('mongoose');

const formatDataF4Schema = new mongoose.Schema({
  month: { type: String, required: true },
  year: { type: String, required: true },
  meterSerialNumber: { type: String, required: true },
  title: { type: String }, 
  dataEntries: [{
    time: { type: String },
    kwhImport: { type: Number },
    kwhExport: { type: Number },
    kvahImport: { type: Number },
    kvahExport: { type: Number },
    q1: { type: Number },
    q2: { type: Number },
    q3: { type: Number },
    q4: { type: Number },
    volR: { type: String },
    volY: { type: String },
    volB: { type: String },
    curR: { type: String },
    curY: { type: String },
    curB: { type: String },
    frequency: { type: Number },
    pfImport: { type: String },
    pfExport: { type: String }
  }]
}, { timestamps: true });

formatDataF4Schema.index({ meterSerialNumber: 1, month: 1, year: 1 }, { unique: true });

module.exports = mongoose.model('FormatDataF4', formatDataF4Schema);
