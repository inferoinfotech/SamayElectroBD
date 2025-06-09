const mongoose = require('mongoose');

const lossesCalculationSchema = new mongoose.Schema({
  // Main Client Details
  mainClientId: { type: mongoose.Schema.Types.ObjectId, required: true, ref: 'MainClient' }, // Reference to Main Client
  month: { type: Number, required: true },
  year: { type: Number, required: true },
  SLDCGROSSINJECTION: { type: Number },
  SLDCGROSSDRAWL: { type: Number },

  // Main Client Meter Details
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
      email: { type: String },
      ctptSrNo: { type: String },
      ctRatio: { type: String },
      ptRatio: { type: String },
      mf: { type: Number },
    },

    // Gross Injection and Drawl in MWH
    grossInjectionMWH: { type: Number, required: true },
    drawlMWH: { type: Number, required: true },
    netInjectionMWH: { type: Number, required: true },
    asperApprovedbySLDCGROSSINJECTION: { type: Number },
    asperApprovedbySLDCGROSSDRAWL: { type: Number },

    // Main Client Meter Details (Date, Time, Gross Injected Units Total)
    mainClientMeterDetails: [{
      date: { type: String, required: true },
      time: { type: String, required: true },
      grossInjectedUnitsTotal: { type: Number, required: true }
    }],
  },

  // Sub Client Details (multiple sub-clients under the main client)
  subClient: [{
    name: { type: String },
    divisionName: { type: String },
    consumerNo: { type: String },
    contactNo: { type: String },
    email: { type: String },
    discom : { type: String },
    subClientId: { type: mongoose.Schema.Types.ObjectId, ref: 'SubClient' },
    meterNumber: { type: String },
    meterType: { type: String },
    ctptSrNo: { type: String },
    ctRatio: { type: String },
    ptRatio: { type: String },
    mf: { type: Number },
    voltageLevel: { type: String },
    acCapacityKw: { type: Number },

    // Sub Client Meter Data (GROSS INJECTION, DRAWL, NET INJECTION, etc.)
    subClientsData: {
      grossInjectionMWH: { type: Number, required: true },
      drawlMWH: { type: Number, required: true },
      netInjectionMWH: { type: Number, required: true },
      grossInjectionMWHAfterLosses: { type: Number, required: true },
      drawlMWHAfterLosses: { type: Number, required: true },
      netInjectionMWHAfterLosses: { type: Number, required: true },
      weightageGrossInjecting: { type: Number, required: true },
      weightageGrossDrawl: { type: Number, required: true },
      lossesInjectedUnits: { type: Number, required: true },
      inPercentageOfLossesInjectedUnits: { type: Number, required: true },
      lossesDrawlUnits: { type: Number, required: true },
      inPercentageOfLossesDrawlUnits: { type: Number, required: true },
      partclient: [{
        divisionName: { type: String },
        consumerNo: { type: String },
        sharingPercentage: { type: Number },
        grossInjectionMWHAfterLosses: { type: Number },
        drawlMWHAfterLosses: { type: Number },
        netInjectionMWHAfterLosses: { type: Number },
        grossInjectionMWH: { type: Number},
        drawlMWH: { type: Number},
        netInjectionMWH: { type: Number},
        weightageGrossInjecting: { type: Number},
        weightageGrossDrawl: { type: Number},
        lossesInjectedUnits: { type: Number},
        inPercentageOfLossesInjectedUnits: { type: Number },
        lossesDrawlUnits: { type: Number },
        inPercentageOfLossesDrawlUnits: { type: Number},
      }],

      // Sub Client Meter Data (Date, Time, Gross Injected Units Total, NET Total after Losses)
      subClientMeterData: [{
        date: { type: String, required: true },
        time: { type: String, required: true },
        grossInjectedUnitsTotal: { type: Number, required: true },
        netTotalAfterLosses: { type: Number, required: true },
        partclient: [{
          divisionName: { type: String },
          netTotalAfterLosses: { type: Number },
        }],
      }],
    },
  }],

  // Sub Client Overall (sum of all sub-client data)
  subClientoverall: {
    overallGrossInjectedUnits: { type: Number, required: true },
    grossDrawlUnits: { type: Number, required: true },
  },

  // Difference between Main Client and Sub Clients
  difference: {
    diffInjectedUnits: { type: Number, required: true },
    diffDrawlUnits: { type: Number, required: true }
  },

}, { timestamps: true });

// Exporting the model
module.exports = mongoose.model('LossesCalculationData', lossesCalculationSchema);
