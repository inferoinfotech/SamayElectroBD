// mainClient.model.js
const mongoose = require('mongoose');
const History = require('./history.model'); // Import the History model
const e = require('express');

// Main Client Schema to store and track client information.
const mainClientSchema = new mongoose.Schema({
    name: { type: String, required: true },
    subTitle: { type: String },

    // ABT MAIN METER Details
    abtMainMeter: {
        meterNumber: { type: String, required: true,unique: true },
        modemNumber: { type: String },
        mobileNumber: { type: Number },
        simNumber: { type: String }
    },

    // ABT CHECK METER Details
    abtCheckMeter: {
        meterNumber: { type: String , unique: true },
        modemNumber: { type: String },
        mobileNumber: { type: Number },
        simNumber: { type: String }
    },

    // Voltage and Capacity Data
    voltageLevel: { type: String },
    ctptSrNo: { type: String },
    ctRatio: { type: String },
    ptRatio: { type: String },
    mf: { type: Number },
    pn: { type: Number ,enum: [1,-1] },
    REtype: { type: String, enum: ['Solar', 'Wind', 'Hybrid'] },

    // Additional fields for capacity
    acCapacityKw: { type: Number },
    dcCapacityKwp: { type: Number },
    dcAcRatio: { type: Number },
 
    // Additional fields for modules and inverters
    noOfModules: { type: Number },
    numberOfInverters: { type: Number },
    sharingPercentage: { type: String },

    // Other fields for client information
    contactNo: { type: String },
    email: { type: String },

    // History reference
    history: [{ type: mongoose.Schema.Types.ObjectId, ref: 'History' }]  // Reference to History collection

}, { timestamps: true });

// Add the updateField method as an instance method to the schema
mainClientSchema.methods.updateField = async function(fieldName, newValue) {
    const oldValue = this[fieldName];  // Get the current value of the field

    // Ensure that the field exists on the document
    if (this[fieldName] === undefined) {
        throw new Error(`Field ${fieldName} does not exist on this client`);
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

    // Update the field with the new value in the main client document
    this[fieldName] = newValue;

    // Save the document after the update
    await this.save();
};

module.exports = mongoose.model('MainClient', mainClientSchema);
