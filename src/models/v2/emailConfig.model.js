// emailConfig.model.js
const mongoose = require('mongoose');

// Email Configuration Schema for templates and settings
const emailConfigSchema = new mongoose.Schema({
    // Configuration type
    configType: {
        type: String,
        enum: ['weekly', 'monthly', 'general'],
        required: true,
        unique: true
    },
    
    // Email template
    template: {
        subject: { type: String, required: true },
        body: { type: String, required: true }, // HTML supported
        isActive: { type: Boolean, default: true }
    },
    
    // Recipients configuration
    recipients: {
        // Main recipients (TO) with their specific CC emails
        clients: [{
            clientId: {
                type: mongoose.Schema.Types.ObjectId,
                ref: 'MainClient',
                required: true
            },
            clientName: { type: String },
            consumerNo: { type: String },
            email: { type: String },
            // Client-specific CC emails
            ccEmails: [{
                email: { type: String, required: true },
                name: { type: String }
            }]
        }],
        
        // Global CC recipients (deprecated - keeping for backward compatibility)
        ccEmails: [{
            email: { type: String, required: true },
            name: { type: String }
        }]
    },
    
    // Last updated info
    updatedBy: {
        type: mongoose.Schema.Types.ObjectId,
        ref: 'User'
    }

}, { timestamps: true });

// Index for faster queries
emailConfigSchema.index({ configType: 1 });

module.exports = mongoose.model('EmailConfig', emailConfigSchema);
