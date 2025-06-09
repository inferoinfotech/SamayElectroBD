const mongoose = require('mongoose');

const meterDataSchema = new mongoose.Schema({
    meterNumber: { type: String, required: true },
    clientType: { type: String, enum: ['MainClient', 'SubClient'], required: true },
    meterType: { type: String, enum: ['abtMainMeter', 'abtCheckMeter'], required: true },
    client: { type: mongoose.Schema.Types.ObjectId, required: true, refPath: 'clientType' },

    month: { type: Number, required: true },
    year: { type: Number, required: true },

    dataEntries: [{
        date: { type: Date, required: true },
        intervalStart: { type: String, required: true },
        intervalEnd: { type: String, required: true },
        parameters: mongoose.Schema.Types.Mixed // Dynamic fields
    }],

}, { timestamps: true });

meterDataSchema.index({ meterNumber: 1, month: 1, year: 1 }, { unique: true });

module.exports = mongoose.model('MeterData', meterDataSchema);
