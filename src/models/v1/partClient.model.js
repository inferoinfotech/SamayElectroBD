const mongoose = require('mongoose');
const History = require('./history.model'); // Import the History model

// Part Client Schema to store and track part client information
const partClientSchema = new mongoose.Schema({
    sharingPercentage: { type: String, required: true },  // Percentage Sharing
    divisionName: { type: String, required: true },      // Division Name
    consumerNo: { type: String, required: true, unique: true },        // Consumer No.
    discom: { type: String, required: true },            // Distribution Company

    // Reference to SubClient
    subClient: { type: mongoose.Schema.Types.ObjectId, ref: 'SubClient', required: true },

    // History for fields (Store references to History documents)
    history: [{ type: mongoose.Schema.Types.ObjectId, ref: 'History' }], // References to History documents

}, { timestamps: true });

// Method to update a field and track history for Part Client
partClientSchema.methods.updateField = async function (fieldName, newValue) {
    const oldValue = this[fieldName];  // Get the current value of the field

    // Ensure the field exists on the document
    if (this[fieldName] === undefined) {
        throw new Error(`Field ${fieldName} does not exist on this part client`);
    }

    // Create history entry (referencing the History model)
    const historyEntry = new History({
        clientId: this._id,
        fieldName: fieldName,
        oldValue: oldValue,
        newValue: newValue,
        updatedAt: new Date()
    });

    // Save history entry in the History collection
    await historyEntry.save();

    // Add the history reference to this PartClient's history field
    this.history.push(historyEntry._id);

    // Update the field with the new value in the part client document
    this[fieldName] = newValue;

    // Save the document after the update
    await this.save();
};

module.exports = mongoose.model('PartClient', partClientSchema);
