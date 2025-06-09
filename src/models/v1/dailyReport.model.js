const mongoose = require('mongoose');

const DailyreportSchema = new mongoose.Schema({
    // Main Client Details
    mainClientId: { type: mongoose.Schema.Types.ObjectId, required: true, ref: 'MainClient' }, // Reference to Main Client
    month: { type: Number, required: true },
    year: { type: Number, required: true },

    mainClient: {
        meterNumber: { type: String, required: true },
        meterType: { type: String, required: true },

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
            email: { type: String }
        },
        totalexport: { type: Number, required: true },
        totalimport: { type: Number, required: true },
        loggerdatas: [
            {
                date: { type: String, required: true },
                export: { type: Number, required: true },
                import: { type: Number, required: true }
            }],
    },

    // Sub Client Details (multiple sub-clients under the main client)
    subClient: [{
        name: { type: String },
        divisionName: { type: String },
        consumerNo: { type: String },
        contactNo: { type: String },
        email: { type: String },
        subClientId: { type: mongoose.Schema.Types.ObjectId, ref: 'SubClient' },
        meterNumber: { type: String },
        meterType: { type: String },
        totalexport: { type: Number, required: true },
        totalimport: { type: Number, required: true },
        totalloggerdata: { type: Number, required: true },
        totalinternallosse: { type: Number, required: true },
        totallossinparsantege: { type: Number, required: true },
        loggerdatas: [
            {
                date: { type: String, required: true },
                export: { type: Number, required: true },
                import: { type: Number, required: true },
                loggerdata: { type: Number, required: true },
                internallosse: { type: Number, required: true },
                lossinparsantege: { type: Number, required: true },
            }],
    }],

    aclinelossdiffrence: {

        totalexport: { type: Number, required: true },
        totalimport: { type: Number, required: true },
        totallossinparsantegeexport: { type: Number, required: true },
        totallossinparsantegeimport: { type: Number, required: true },
        loggerdatas: [
            {
                date: { type: String, required: true },
                export: { type: Number, required: true },
                import: { type: Number, required: true },
                lossinparsantegeexport: { type: Number, required: true },
                lossinparsantegeimport: { type: Number, required: true },
            }],
       
    },

}, { timestamps: true });

module.exports = mongoose.model('Dailyreport', DailyreportSchema);
