// clientPolicy.model.js
const mongoose = require('mongoose');
const Policy = require('./policy.model');

// Client Policy Schema to map clients to policies and track which sub-policies apply
const clientPolicySchema = new mongoose.Schema({
    subClientId: { 
        type: mongoose.Schema.Types.ObjectId, 
        ref: 'SubClient', 
        required: true 
    },
    
    policyId: { 
        type: mongoose.Schema.Types.ObjectId, 
        ref: 'Policy', 
        required: true 
        // Reference to the main policy (e.g., "jan-2025")
    },
    
    // Array of sub-policies with apply status
    policies: [{
        policyItemId: { 
            type: mongoose.Schema.Types.ObjectId, 
            required: true 
            // Reference to the sub-policy item from the Policy model
            // This is the _id of the item in the policies array
        },
        apply: { 
            type: Boolean, 
            default: true 
            // true if this sub-policy applies to the client, false otherwise
        },
        customValue: {
            type: mongoose.Schema.Types.Mixed
            // Optional: if client has a custom value for this policy item
        }
    }],
    
    // Optional metadata
    effectiveDate: {
        type: Date,
        default: Date.now
    },
    
    isActive: {
        type: Boolean,
        default: true
    }

}, { timestamps: true });

// Index for faster queries
clientPolicySchema.index({ subClientId: 1, policyId: 1 });
clientPolicySchema.index({ subClientId: 1, isActive: 1 });
clientPolicySchema.index({ policyId: 1 });

// Compound unique index to ensure one client-policy mapping per client-policy combination
clientPolicySchema.index({ subClientId: 1, policyId: 1 }, { unique: true });

module.exports = mongoose.model('ClientPolicy', clientPolicySchema);

