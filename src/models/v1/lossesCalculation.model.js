// models/LossesCalculationData.js
const mongoose = require('mongoose');

/**
 * Optional helper/audit stored per interval for traceability against Excel.
 * Remove this sub-schema if you want leaner documents.
 */
const HelperAuditSchema = new mongoose.Schema(
  {
    raw: { type: Number },            // helper col-1 equivalent (I_E * mf * pn / 1000)
    allocatedGroup: { type: Number }, // placeholder if you later add time-slice allocation
    discomScaled: { type: Number },   // after monthly DISCOM scaling (feeds losses)
  },
  { _id: false }
);

const PartClientMeterRowSchema = new mongoose.Schema(
  {
    divisionName: { type: String },
    netTotalAfterLosses: { type: Number },
  },
  { _id: false }
);

const SubClientMeterRowSchema = new mongoose.Schema(
  {
    date: { type: String, required: true },
    time: { type: String, required: true },
    grossInjectedUnitsTotal: { type: Number, required: true }, // signed MWh (+ inj, - drawl)
    netTotalAfterLosses: { type: Number }, // signed MWh after losses
    partclient: [PartClientMeterRowSchema],
    helper: { type: HelperAuditSchema },   // optional audit
  },
  { _id: false }
);

const PartClientSchema = new mongoose.Schema(
  {
    divisionName: { type: String },
    consumerNo: { type: String },
    sharingPercentage: { type: Number }, // e.g., 63.16

    // Before losses
    grossInjectionMWH: { type: Number },
    drawlMWH: { type: Number },
    netInjectionMWH: { type: Number },

    // After losses
    grossInjectionMWHAfterLosses: { type: Number },
    drawlMWHAfterLosses: { type: Number },
    netInjectionMWHAfterLosses: { type: Number },

    // Weightages & losses inheritance
    weightageGrossInjecting: { type: Number },
    weightageGrossDrawl: { type: Number },
    lossesInjectedUnits: { type: Number },
    inPercentageOfLossesInjectedUnits: { type: Number },
    lossesDrawlUnits: { type: Number },
    inPercentageOfLossesDrawlUnits: { type: Number },
  },
  { _id: false }
);

const SubClientDataSchema = new mongoose.Schema(
  {
    // Month totals BEFORE losses (after DISCOM scaling)
    grossInjectionMWH: { type: Number, required: true },
    drawlMWH: { type: Number, required: true },
    netInjectionMWH: { type: Number, required: true },

    // Month totals AFTER losses
    grossInjectionMWHAfterLosses: { type: Number, required: true },
    drawlMWHAfterLosses: { type: Number, required: true },
    netInjectionMWHAfterLosses: { type: Number, required: true },

    // Weightages & losses percentages
    weightageGrossInjecting: { type: Number, required: true },
    weightageGrossDrawl: { type: Number, required: true },
    lossesInjectedUnits: { type: Number, required: true },
    inPercentageOfLossesInjectedUnits: { type: Number, required: true },
    lossesDrawlUnits: { type: Number, required: true },
    inPercentageOfLossesDrawlUnits: { type: Number, required: true },

    // Part clients split
    partclient: [PartClientSchema],

    // Interval rows (these feed both totals and losses)
    subClientMeterData: [SubClientMeterRowSchema],
  },
  { _id: false }
);

const SubClientSchema = new mongoose.Schema(
  {
    name: { type: String },
    divisionName: { type: String },
    consumerNo: { type: String },
    contactNo: { type: String },
    email: { type: String },
    discom: { type: String }, // DGVCL, MGVCL, PGVCL, UGVCL, TAECO, TSECO, TEL
    subClientId: { type: mongoose.Schema.Types.ObjectId, ref: 'SubClient' },
    meterNumber: { type: String },
    meterType: { type: String },
    ctptSrNo: { type: String },
    ctRatio: { type: String },
    ptRatio: { type: String },
    mf: { type: Number },
    voltageLevel: { type: String },
    acCapacityKw: { type: Number },

    subClientsData: { type: SubClientDataSchema, required: true },
  },
  { _id: false }
);

const MainClientMeterRowSchema = new mongoose.Schema(
  {
    date: { type: String, required: true },
    time: { type: String, required: true },
    grossInjectedUnitsTotal: { type: Number, required: true }, // signed MWh (+ inj, - drawl)
  },
  { _id: false }
);

const MainClientSchema = new mongoose.Schema(
  {
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

    // Month totals BEFORE losses (after main-entry scaling)
    grossInjectionMWH: { type: Number, required: true },
    drawlMWH: { type: Number, required: true },
    netInjectionMWH: { type: Number, required: true },

    // Main-entry diffs (info)
    asperApprovedbySLDCGROSSINJECTION: { type: Number },
    asperApprovedbySLDCGROSSDRAWL: { type: Number },

    mainClientMeterDetails: [MainClientMeterRowSchema],
  },
  { _id: false }
);

const AuditDiscomScaleItemSchema = new mongoose.Schema(
  { fPos: Number, fNeg: Number },
  { _id: false }
);

const lossesCalculationSchema = new mongoose.Schema(
  {
    // Main Client Context
    mainClientId: { type: mongoose.Schema.Types.ObjectId, required: true, ref: 'MainClient' },
    month: { type: Number, required: true },
    year: { type: Number, required: true },

    // MAIN ENTRY (SLDC approved)
    SLDCGROSSINJECTION: { type: Number }, // B9 (+)
    SLDCGROSSDRAWL: { type: Number },     // L9 (-)

    // Optional per-DISCOM fields (store the inputs/targets for easy queries)
    DGVCL: { type: Number },
    MGVCL: { type: Number },
    PGVCL: { type: Number },
    UGVCL: { type: Number },
    TAECO: { type: Number },
    TSECO: { type: Number },
    TEL: { type: Number },

    // Main client block
    mainClient: { type: MainClientSchema, required: true },

    // All sub clients
    subClient: [SubClientSchema],

    // Overall subs (before losses)
    subClientoverall: {
      overallGrossInjectedUnits: { type: Number, required: true },
      grossDrawlUnits: { type: Number, required: true },
    },

    // Differences (overall subs vs main)
    difference: {
      diffInjectedUnits: { type: Number, required: true },
      diffDrawlUnits: { type: Number, required: true },
    },

    // Outputs for the screenshot section
    perDiscomTotals: {
      type: Map, // e.g., { DGVCL: 356.883, ... } after allocation
      of: Number,
    },
    excessInjectionPPA: { type: Number },        // SLDC approved inj - total credited
    energyDrawnFromDiscom: { type: Number },     // equals approved drawl (or main drawl if not provided)

    // Optional audits (factors and sums used to replicate Excel)
    audit: {
      subsPositiveSum: Number, // p
      subsNegativeSum: Number, // n
      mainScale: {
        fPos: Number,
        fNeg: Number,
      },
      discomScale: {
        type: Map,
        of: AuditDiscomScaleItemSchema, // { fPos, fNeg } per DISCOM
      },
    },

    // Clients that used abtCheckMeter as fallback
    clientsUsingCheckMeter: [String],
  },
  { timestamps: true }
);

module.exports = mongoose.model('LossesCalculationData', lossesCalculationSchema);
