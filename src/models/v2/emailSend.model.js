// emailSend.model.js
const mongoose = require('mongoose');

// Email Send Schema to track email sending records
const emailSendSchema = new mongoose.Schema({
    // Type: Weekly, Monthly, or General
    sendType: {
        type: String,
        enum: ['weekly', 'monthly', 'general'],
        required: true
    },
    
    // Period information (optional for general type)
    period: {
        week: { type: Number }, // Week number (1-52)
        month: { type: Number, min: 1, max: 12 }, // Month (1-12)
        year: { type: Number }
    },
    
    // Recipients with their meter details
    recipients: [{
        mainClientId: {
            type: mongoose.Schema.Types.ObjectId,
            ref: 'MainClient',
            required: true
        },
        mainClientName: { type: String },
        meterNumber: { type: String, required: true },
        email: { type: String, required: true },
        
        // Status checks
        csvCheck: { type: Boolean, default: false },
        cdfCheck: { type: Boolean, default: false },
        mailSent: { type: Boolean, default: false },
        
        // Email details
        sentAt: { type: Date },
        emailSubject: { type: String },
        emailBody: { type: String },
        
        // Attachments info
        attachments: [{
            filename: { type: String },
            path: { type: String },
            size: { type: Number }
        }],
        
        // Error tracking
        error: { type: String }
    }],
    
    // Uploaded files for this batch
    uploadedFiles: [{
        filename: { type: String },
        originalName: { type: String },
        path: { type: String },
        size: { type: Number },
        uploadedAt: { type: Date, default: Date.now }
    }],
    
    // Batch information
    totalRecipients: { type: Number, default: 0 },
    successCount: { type: Number, default: 0 },
    failureCount: { type: Number, default: 0 },
    
    // Status
    status: {
        type: String,
        enum: ['draft', 'processing', 'completed', 'failed'],
        default: 'draft'
    },
    
    // Notes
    notes: { type: String },
    
    // Created by user
    createdBy: {
        type: mongoose.Schema.Types.ObjectId,
        ref: 'User'
    }

}, { timestamps: true });

// Indexes for faster queries
emailSendSchema.index({ sendType: 1, 'period.year': 1, 'period.month': 1 });
emailSendSchema.index({ 'recipients.meterNumber': 1 });
emailSendSchema.index({ status: 1 });
emailSendSchema.index({ createdAt: -1 });

module.exports = mongoose.model('EmailSend', emailSendSchema);
