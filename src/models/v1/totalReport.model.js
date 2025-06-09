const mongoose = require('mongoose');

const TotalreportSchema = new mongoose.Schema({
    month: { type: Number, required: true },
    year: { type: Number, required: true },
    clients: [{
        mainClientId: { type: mongoose.Schema.Types.ObjectId, required: true, ref: 'MainClient' }, // Reference to Main Client
        mainClientDetail: {
            name: { type: String },
            subTitle: { type: String },
            abtMainMeter: { type: mongoose.Schema.Types.Mixed },
            abtCheckMeter: { type: mongoose.Schema.Types.Mixed },
            voltageLevel: { type: String },
            acCapacityKw: { type: Number },
            dcCapacityKwp: { type: Number },
            noOfModules: { type: Number },
            sharingPercentage: { type: String },
            contactNo: { type: String },
            email: { type: String },
            mf: { type: Number },
        },
        abtMainMeter: {
            meterNumber: { type: String, required: true },
            grossInjectedUnits: { type: Number, required: true },
            grossDrawlUnits: { type: Number, required: true },
            totalimport: { type: Number, required: true },
        },
        abtCheckMeter: {
            meterNumber: { type: String, required: true },
            grossInjectedUnits: { type: Number, required: true },
            grossDrawlUnits: { type: Number, required: true },
            totalimport: { type: Number, required: true },
        },
        difference:{
            grossInjectedUnits: { type: Number, required: true },
            grossDrawlUnits: { type: Number, required: true },
        }
    }]


}, { timestamps: true });

module.exports = mongoose.model('Totalreport', TotalreportSchema);
