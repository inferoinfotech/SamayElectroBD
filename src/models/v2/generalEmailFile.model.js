// generalEmailFile.model.js - Store uploaded files for General Monthly Email
const mongoose = require('mongoose');

const generalEmailFileSchema = new mongoose.Schema({
    clientId: {
        type: mongoose.Schema.Types.ObjectId,
        ref: 'SubClient',
        required: true
    },
    clientName: {
        type: String,
        required: true
    },
    year: {
        type: Number,
        required: true
    },
    month: {
        type: Number,
        required: true,
        min: 1,
        max: 12
    },
    files: [{
        filename: {
            type: String,
            required: true
        },
        originalName: {
            type: String,
            required: true
        },
        path: {
            type: String,
            required: true
        },
        size: {
            type: Number,
            required: true
        },
        mimetype: {
            type: String,
            required: true
        },
        uploadedAt: {
            type: Date,
            default: Date.now
        }
    }],
    lastEmailSentAt: {
        type: Date
    },
    emailSentCount: {
        type: Number,
        default: 0
    }
}, {
    timestamps: true
});

// Index for faster queries
generalEmailFileSchema.index({ clientId: 1, year: 1, month: 1 });
generalEmailFileSchema.index({ year: 1, month: 1 });

const GeneralEmailFile = mongoose.model('GeneralEmailFile', generalEmailFileSchema);

module.exports = GeneralEmailFile;
