const mongoose = require('mongoose');

const historySchema = new mongoose.Schema({
    clientId: { type: mongoose.Schema.Types.ObjectId, ref: 'MainClient', required: true },
    fieldName: { type: String, required: true },
    oldValue: { type: mongoose.Schema.Types.Mixed, required: true },
    newValue: { type: mongoose.Schema.Types.Mixed, required: true },
    updatedAt: { type: Date, default: Date.now }
});

module.exports = mongoose.model('History', historySchema);
