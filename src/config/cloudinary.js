// cloudinary.js - Cloudinary Configuration
const cloudinary = require('cloudinary').v2;
const logger = require('../utils/logger');

// Configure Cloudinary
cloudinary.config({
    cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
    api_key: process.env.CLOUDINARY_API_KEY,
    api_secret: process.env.CLOUDINARY_API_SECRET
});

// Test connection
const testConnection = async () => {
    try {
        const result = await cloudinary.api.ping();
        logger.info('Cloudinary connected successfully');
        return true;
    } catch (error) {
        logger.error(`Cloudinary connection failed: ${error.message}`);
        return false;
    }
};

// Upload file to Cloudinary
const uploadToCloudinary = async (filePath, options = {}) => {
    try {
        const result = await cloudinary.uploader.upload(filePath, {
            folder: options.folder || 'samay-electro/email-attachments',
            resource_type: 'auto', // Automatically detect file type
            use_filename: true,
            unique_filename: true,
            ...options
        });

        logger.info(`File uploaded to Cloudinary: ${result.secure_url}`);
        return {
            success: true,
            url: result.secure_url,
            publicId: result.public_id,
            format: result.format,
            size: result.bytes,
            resourceType: result.resource_type
        };
    } catch (error) {
        logger.error(`Cloudinary upload failed: ${error.message}`);
        return {
            success: false,
            error: error.message
        };
    }
};

// Delete file from Cloudinary
const deleteFromCloudinary = async (publicId, resourceType = 'auto') => {
    try {
        const result = await cloudinary.uploader.destroy(publicId, {
            resource_type: resourceType
        });

        logger.info(`File deleted from Cloudinary: ${publicId}`);
        return {
            success: true,
            result
        };
    } catch (error) {
        logger.error(`Cloudinary delete failed: ${error.message}`);
        return {
            success: false,
            error: error.message
        };
    }
};

// Get file URL from Cloudinary
const getCloudinaryUrl = (publicId, options = {}) => {
    return cloudinary.url(publicId, {
        secure: true,
        ...options
    });
};

module.exports = {
    cloudinary,
    testConnection,
    uploadToCloudinary,
    deleteFromCloudinary,
    getCloudinaryUrl
};
