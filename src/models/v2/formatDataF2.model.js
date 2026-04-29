const mongoose = require('mongoose');

const formatDataF2Schema = new mongoose.Schema({
  month: { type: String, required: true },
  year: { type: String, required: true },
  meterSerialNumber: { type: String, required: true },
  dataEntries: [{
    meterTimestamp: { type: String },
    entryTimestamp: { type: String },
    meterDataCaptureTimestamp: { type: String },
    kwhImport: { type: Number },
    kvahImport: { type: Number },
    fundamentalEnergyImportActiveEnergy: { type: Number },
    blockFrequency: { type: Number },
    avgVoltageVRN: { type: Number },
    avgVoltageVYN: { type: Number },
    avgVoltageVBN: { type: Number },
    netActiveEnergy: { type: Number },
    blockEnergyKWhExp: { type: Number },
    blockEnergyKVArhQ1: { type: Number },
    blockEnergyKVArhQ2: { type: Number },
    blockEnergyKVArhQ3: { type: Number },
    blockEnergyKVArhQ4: { type: Number },
    blockEnergyKVahExp: { type: Number },
    avgCurrentIR: { type: Number },
    avgCurrentIY: { type: Number },
    avgCurrentIB: { type: Number },
    reactiveEnergyHigh: { type: Number },
    reactiveEnergyLow: { type: Number },
    exportActiveEnergy: { type: Number },
    totalPowerFactor: { type: Number },
    exportPowerFactor: { type: Number },
    codedFrequency: { type: Number },
    averageVoltage: { type: Number },
    netKVARH: { type: Number },
    loadSurveyTamperSnap: { type: String }, // Fixed: Can be "-" or "POWER FAILURE"
    billingAverageVoltage: { type: Number }
  }]
}, { timestamps: true });

formatDataF2Schema.index({ meterSerialNumber: 1, month: 1, year: 1 }, { unique: true });

module.exports = mongoose.model('FormatDataF2', formatDataF2Schema);
