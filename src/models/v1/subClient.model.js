// subClient.model.js
const mongoose = require('mongoose');
const History = require('./history.model'); // Import the History model
const { unix } = require('moment');

// Sub Client Schema to store and track client information
const subClientSchema = new mongoose.Schema({
    name: { type: String, required: true },
    divisionName: { type: String },
    discom : { type: String },
    consumerNo: { type: String, unique: true },
    // modemSrNo: { type: String },

    // ABT MAIN METER Details
    abtMainMeter: {
        meterNumber: { type: String, require : true, unique: true },
        modemNumber: { type: String },
        mobileNumber: { type: Number },
        simNumber: { type: String }
    },

    // ABT CHECK METER Details
    abtCheckMeter: {
        meterNumber: { type: String, unique: true },
        modemNumber: { type: String },
        mobileNumber: { type: Number },
        simNumber: { type: String }
    },
    voltageLevel: { type: String },
    ctptSrNo: { type: String },
    ctRatio: { type: String },
    ptRatio: { type: String },
    mf: { type: Number },
    pn: { type: Number ,enum: [1,-1] },
    acCapacityKw: { type: Number },
    dcCapacityKwp: { type: Number },
    dcAcRatio: { type: Number },
    noOfModules: { type: Number },
    moduleCapacityWp: { type: Number },
    inverterCapacityKw: { type: Number },
    numberOfInverters: { type: Number },
    makeOfInverter: { type: String },
    sharingPercentage: { type: String },
    contactNo: { type: String },
    email: { type: String },
    REtype: { type: String , enum: ['Solar', 'Wind', 'Hybrid'] },

    // Reference to MainClient
    mainClient: { type: mongoose.Schema.Types.ObjectId, ref: 'MainClient', required: true },

    // History for fields
    history: [{ type: mongoose.Schema.Types.ObjectId, ref: 'History' }], // Referencing History model
}, { timestamps: true });

// Method to update a field and track history
subClientSchema.methods.updateField = async function (fieldName, newValue) {
    const oldValue = this[fieldName];  // Get the current value of the field

    // Ensure the field exists on the document
    if (this[fieldName] === undefined) {
        throw new Error(`Field ${fieldName} does not exist on this sub client`);
    }

    // Save history to the History collection
    const historyEntry = new History({
        clientId: this._id,
        fieldName: fieldName,
        oldValue: oldValue,
        newValue: newValue,
        updatedAt: new Date()
    });

    await historyEntry.save(); // Save history

    // Update the field with the new value in the sub client document
    this[fieldName] = newValue;

    // Save the document after the update
    await this.save();
};

module.exports = mongoose.model('SubClient', subClientSchema);
