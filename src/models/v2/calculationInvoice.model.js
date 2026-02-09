// calculationInvoice.model.js
const mongoose = require('mongoose');

// Calculation Invoice Schema to store policy calculation data
const calculationInvoiceSchema = new mongoose.Schema({
    subClientId: {
        type: mongoose.Schema.Types.ObjectId,
        ref: 'SubClient',
        required: true
    },
    
    policyId: {
        type: mongoose.Schema.Types.ObjectId,
        ref: 'Policy',
        required: true
    },
    
    // Selected policies with apply status
    selectedPolicies: [{
        policyItemId: {
            type: mongoose.Schema.Types.ObjectId,
            required: true
        },
        apply: {
            type: Boolean,
            default: true
        }
    }],
    
    // Solar Generation Date (legacy - kept for backward compatibility)
    solarGenerationMonth: {
        type: Number,
        required: false,
        min: 1,
        max: 12
    },
    
    solarGenerationYear: {
        type: Number,
        required: false
    },
    
    // Solar Generation Months (new structure - supports multiple months)
    solarGenerationMonths: [{
        id: { type: String },
        month: { type: Number, min: 1, max: 12 },
        year: { type: Number },
        policyId: {
            type: mongoose.Schema.Types.ObjectId,
            ref: 'Policy'
        },
        selectedPolicies: [{
            policyItemId: {
                type: mongoose.Schema.Types.ObjectId,
                required: true
            },
            apply: {
                type: Boolean,
                default: true
            }
        }],
        setOffEntry: {
            generationUnit: { type: Number },
            drawlUnit: { type: Number },
            bankingCharges: { type: Number },
            energyCharge: { type: Number },
            totalSetoff: { type: Number }
        },
        todEntry: {
            rate: { type: Number },
            amount: { type: Number },
            unit: { type: Number }
        },
        roofTop: {
            billingCharge: { type: Number },
            totalGeneration: { type: Number },
            unit: { type: Number },
            roofTopBanking: { type: Number },
            anyOtherCredit: { type: Number },
            anyOtherDebit: { type: Number }
        },
        windFarm: {
            unit: { type: Number },
            amount: { type: Number },
            rate: { type: Number }
        },
        anyOther: {
            credit: { type: Number },
            debit: { type: Number },
            tcsTds: { type: Number }
        },
        calculationTable: { type: mongoose.Schema.Types.Mixed } // Each month has its own calculation table
    }],
    
    // Adjustment Billing Date
    adjustmentBillingMonth: {
        type: Number,
        required: true,
        min: 1,
        max: 12
    },
    
    adjustmentBillingYear: {
        type: Number,
        required: true
    },
    
    // DGVCL CREDIT (shared across all months)
    dgvclCredit: {
        type: Number
    },
    
    // 5.1 Set Off Entry (legacy - kept for backward compatibility)
    setOffEntry: {
        generationUnit: { type: Number },
        drawlUnit: { type: Number },
        bankingCharges: { type: Number },
        energyCharge: { type: Number },
        totalSetoff: { type: Number },
        dgvclCredit: { type: Number }
    },
    
    // 5.2 TOD Entry
    todEntry: {
        rate: { type: Number },
        amount: { type: Number },
        unit: { type: Number }
    },
    
    // 5.3 Roof Top
    roofTop: {
        billingCharge: { type: Number },
        totalGeneration: { type: Number },
        unit: { type: Number },
        roofTopBanking: { type: Number },
        anyOtherCredit: { type: Number },
        anyOtherDebit: { type: Number }
    },
    
    // 5.4 Wind Farm (legacy - kept for backward compatibility)
    windFarm: {
        unit: { type: Number },
        amount: { type: Number },
        rate: { type: Number }
    },
    
    // 5.5 Any Other (legacy - kept for backward compatibility)
    anyOther: {
        credit: { type: Number },
        debit: { type: Number },
        tcsTds: { type: Number }
    },
    
    // 6. Manual Entry
    manualEntry: {
        totalConsumption: { type: Number },
        contractDemand: { type: Number },
        contractDemand85Percent: { type: Number },
        maxUtilizationContractDemand: { type: Number },
        pfPercentage: { type: Number },
        pfRebateAmount: { type: Number },
        totalCurrentMonthBillAmount: { type: Number },
        totalAdjustmentAmount: { type: Number },
        settlementThreshold: { type: Number },
        anyOther: { type: Number },
        // Dynamic custom fields
        customFields: [{
            fieldName: { type: String, required: true },
            value: { type: mongoose.Schema.Types.Mixed, required: true }
        }]
    },
    
    // 7. Calculation Table
    calculationTable: {
        // Section 1: As Per RE Solar Policy-2023
        section1: {
            "1.1": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: {
                    electricityDuty: {
                        unitsInKwh: { type: Number },
                        rate: { type: Number },
                        creditAmount: { type: Number },
                        debitAmount: { type: Number },
                        remark: { type: String }
                    },
                    drawlFromDiscom: {
                        unitsInKwh: { type: Number },
                        rate: { type: Number },
                        creditAmount: { type: Number },
                        debitAmount: { type: Number },
                        remark: { type: String }
                    }
                }
            },
            "1.2": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: { type: mongoose.Schema.Types.Mixed }
            },
            "1.3": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: { type: mongoose.Schema.Types.Mixed }
            },
            "1.4": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: {
                    electricityDuty: {
                        unitsInKwh: { type: Number },
                        rate: { type: Number },
                        creditAmount: { type: Number },
                        debitAmount: { type: Number },
                        remark: { type: String }
                    },
                    bankedEnergyPercent: {
                        unitsInKwh: { type: Number },
                        rate: { type: Number },
                        creditAmount: { type: Number },
                        debitAmount: { type: Number },
                        remark: { type: String }
                    }
                }
            },
            "1.5": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: { type: mongoose.Schema.Types.Mixed }
            },
            "1.6": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: { type: mongoose.Schema.Types.Mixed }
            },
            "1.7": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: { type: mongoose.Schema.Types.Mixed }
            },
            "1.8": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: { type: mongoose.Schema.Types.Mixed }
            },
            "1.9": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: { type: mongoose.Schema.Types.Mixed }
            },
            "1.10": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: { type: mongoose.Schema.Types.Mixed }
            }
        },
        // Section 2: Other Credit (if any)
        section2: {
            "2.1": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: { type: mongoose.Schema.Types.Mixed }
            },
            "2.2": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: { type: mongoose.Schema.Types.Mixed }
            },
            "2.3": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: { type: mongoose.Schema.Types.Mixed }
            },
            "2.4": {
                unitsInKwh: { type: Number },
                rate: { type: Number },
                creditAmount: { type: Number },
                debitAmount: { type: Number },
                remark: { type: String },
                showElectricityDuty: { type: Boolean, default: false },
                subRows: { type: mongoose.Schema.Types.Mixed }
            }
        },
        // Section 3: Other (dynamic array)
        section3: [{
            id: { type: String, required: true },
            particulars: { type: String },
            unitsInKwh: { type: Number },
            rate: { type: Number },
            creditAmount: { type: Number },
            debitAmount: { type: Number },
            remark: { type: String },
            showElectricityDuty: { type: Boolean, default: false },
            subRows: { type: mongoose.Schema.Types.Mixed }
        }],
        // Final Section
        finalSection: {
            row1: {
                creditAmount: { type: Number },
                debitAmount: { type: Number }
            },
            row2: {
                mergedValue: { type: Number }
            },
            row3: {
                mergedValue: { type: Number }
            },
            row4: {
                mergedValue: { type: Number },
                status: { type: String }
            }
        }
    },
    
    // Optional metadata
    notes: {
        type: String
    },
    
    isActive: {
        type: Boolean,
        default: true
    }

}, { timestamps: true });

// Index for faster queries
calculationInvoiceSchema.index({ subClientId: 1, policyId: 1 });
calculationInvoiceSchema.index({ solarGenerationYear: 1, solarGenerationMonth: 1 });
calculationInvoiceSchema.index({ adjustmentBillingYear: 1, adjustmentBillingMonth: 1 });

module.exports = mongoose.model('CalculationInvoice', calculationInvoiceSchema);

