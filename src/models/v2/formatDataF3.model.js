const mongoose = require('mongoose');

const formatDataF3Schema = new mongoose.Schema({
  month: { type: String, required: true },
  year: { type: String, required: true },
  meterSerialNumber: { type: String, required: true },
  dataDumpTime: { type: String },
  dataReadTimeMRI: { type: String },
  dataReadTimeMeter: { type: String },
  meterType: { type: String },
  location: { type: String },
  logInterval: { type: Number },
  emf: { type: Number },
  emfApplied: { type: String },
  installedCTRatio: { type: Number },
  installedPTRatio: { type: Number },
  commissionCTRatio: { type: Number },
  commissionPTRatio: { type: Number },
  lsDays: { type: Number },
  startDate: { type: String },
  endDate: { type: String },
  dataEntries: [{
    date: { type: String },
    intervalStart: { type: String },
    intervalEnd: { type: String },
    kwI: { type: Number },
    kwE: { type: Number },
    kvaI: { type: Number },
    kvaE: { type: Number },
    kwhI: { type: Number },
    kwhE: { type: Number },
    kwhNet: { type: Number },
    kvahI: { type: Number },
    kvahE: { type: Number }
  }]
}, { timestamps: true });

formatDataF3Schema.index({ meterSerialNumber: 1, month: 1, year: 1 }, { unique: true });

module.exports = mongoose.model('FormatDataF3', formatDataF3Schema);
