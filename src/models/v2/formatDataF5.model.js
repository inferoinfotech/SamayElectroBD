const mongoose = require('mongoose');

const formatDataF5Schema = new mongoose.Schema({
  month: { type: String, required: true },
  year: { type: String, required: true },
  meterSerialNumber: { type: String, required: true },
  dataEntries: [{
    date: { type: String },
    intervalStart: { type: String },
    intervalEnd: { type: String },
    kwImp: { type: Number },
    kwExp: { type: Number },
    kvaImp: { type: Number },
    kvaExp: { type: Number },
    kwhImp: { type: Number },
    kwhExp: { type: Number },
    netKwh: { type: Number },
    kvahImp: { type: Number },
    kvahExp: { type: Number },
    kvarhLgDurWhImp: { type: Number },
    kvarhLdDurWhImp: { type: Number },
    kvarhLgDurWhExp: { type: Number },
    kvarhLdDurWhExp: { type: Number },
    avgRPhVolt: { type: Number },
    avgYPhVolt: { type: Number },
    avgBPhVolt: { type: Number },
    averageVolt: { type: Number },
    avgRPhAmp: { type: Number },
    avgYPhAmp: { type: Number },
    avgBPhAmp: { type: Number },
    averageAmp: { type: Number },
    codedFreq: { type: String },
    avgFreq: { type: Number },
    powerOffMinutes: { type: Number },
    kvarhHighImport: { type: Number },
    kvarhHighExport: { type: Number },
    kvarhLowImport: { type: Number },
    kvarhLowExport: { type: Number },
    kvarhNetReacHigh: { type: Number },
    kvarhNetReacLow: { type: Number },
    powerFactorImp: { type: String },
    powerFactorExp: { type: Number }
  }]
}, { timestamps: true });

formatDataF5Schema.index({ meterSerialNumber: 1, month: 1, year: 1 }, { unique: true });

module.exports = mongoose.model('FormatDataF5', formatDataF5Schema);
