const mongoose = require('mongoose');

const formatDataF1Schema = new mongoose.Schema({
  month: { type: String, required: true },
  year: { type: String, required: true },
  meterSerialNumber: { type: String, required: true },
  dataEntries: [{
    meterTimestamp: { type: String },
    entryTimestamp: { type: String },
    meterDataCaptureTimestamp: { type: String },
    blockFrequency: { type: Number },
    kwhImport: { type: Number },
    blockEnergyKWhExp: { type: Number },
    blockEnergyKVArhQ1: { type: Number },
    blockEnergyKVArhQ2: { type: Number },
    blockEnergyKVArhQ3: { type: Number },
    blockEnergyKVArhQ4: { type: Number },
    kvahImport: { type: Number },
    blockEnergyKVahExp: { type: Number },
    netActiveEnergy: { type: Number },
    averagePhaseVoltages: { type: Number },
    averageLineCurrents: { type: Number }
  }]
}, { timestamps: true });

formatDataF1Schema.index({ meterSerialNumber: 1, month: 1, year: 1 }, { unique: true });

module.exports = mongoose.model('FormatDataF1', formatDataF1Schema);
