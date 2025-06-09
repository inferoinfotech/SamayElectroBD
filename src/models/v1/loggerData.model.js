const mongoose = require('mongoose');

const loggerDataSchema = new mongoose.Schema({
  month: { type: String, required: true },
  year: { type: String, required: true },
  subClient: [{ // Array of sub-client data
    subClientId: { type: mongoose.Schema.Types.ObjectId, ref: 'SubClient', required: true },
    subClientName: { type: String, required: true },
    meterNumber: { type: String, required: true },
    meterType: { type: String, enum: ['abtMainMeter', 'abtCheckMeter'], required: true },
    loggerEntries: [{
      date: { type: String, required: true },
      data: { type: Number, required: true }
    }]
  }],
}, { timestamps: true });

module.exports = mongoose.model('LoggerData', loggerDataSchema);
