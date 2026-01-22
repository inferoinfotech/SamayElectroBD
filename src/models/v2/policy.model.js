// policy.model.js
const mongoose = require('mongoose');

// Policy Schema to store main policies with sub-policies
const policySchema = new mongoose.Schema({
    name: { 
        type: String, 
        required: true,
        unique: true,
        trim: true
        // Example: "jan-2025", "feb-2025", etc.
    },
    
    // Array of sub-policies
    policies: [{
        key: { 
            type: String, 
            required: true,
            trim: true
            // Example: "unit price", "tax rate", "discount", etc.
        },
        value: { 
            type: mongoose.Schema.Types.Mixed, 
            required: true
            // Can be String, Number, Boolean, etc.
            // Example: "25", 25, true, etc.
        },
        description: {
            type: String,
            trim: true
            // Optional description for the policy
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
policySchema.index({ name: 1 });
policySchema.index({ isActive: 1 });

module.exports = mongoose.model('Policy', policySchema);

