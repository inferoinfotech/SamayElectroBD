// emailSendController.js
const EmailSend = require('../../models/v2/emailSend.model');
const EmailConfig = require('../../models/v2/emailConfig.model');
const GeneralEmailFile = require('../../models/v2/generalEmailFile.model');
const MainClient = require('../../models/v1/mainClient.model');
const SubClient = require('../../models/v1/subClient.model');
const logger = require('../../utils/logger');
const mongoose = require('mongoose');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');
const csv = require('csv-parser');
const { reencodeCSVtoUTF8 } = require('../../utils/reencodeCSV');
const { prepareCSVStream } = require('../../utils/prepareCSVStream');
const archiver = require('archiver');
const { uploadToCloudinary, deleteFromCloudinary } = require('../../config/cloudinary');

// Create email transporter
const createTransporter = () => {
    return nodemailer.createTransport({
        host: process.env.EMAIL_HOST,
        port: process.env.EMAIL_PORT,
        secure: false,
        auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASS
        }
    });
};

// Helper function: Sanitize filename for Windows/Unix compatibility
const sanitizeFilename = (filename) => {
    // Remove or replace invalid characters: / \ : * ? " < > |
    return filename
        .replace(/[\/\\:*?"<>|]/g, '_')  // Replace invalid chars with underscore
        .replace(/\s+/g, '_')             // Replace spaces with underscore
        .replace(/_+/g, '_')              // Replace multiple underscores with single
        .trim();
};

// Helper function: Download file from Cloudinary URL
const downloadFromCloudinary = async (cloudinaryUrl, localPath) => {
    return new Promise((resolve, reject) => {
        const https = require('https');
        const file = fs.createWriteStream(localPath);
        
        https.get(cloudinaryUrl, (response) => {
            response.pipe(file);
            file.on('finish', () => {
                file.close();
                logger.info(`File downloaded from Cloudinary: ${localPath}`);
                resolve(localPath);
            });
        }).on('error', (err) => {
            fs.unlink(localPath, () => {}); // Delete incomplete file
            logger.error(`Cloudinary download error: ${err.message}`);
            reject(err);
        });
    });
};

// Helper function: Create ZIP file from multiple files
const createZipFile = async (filePaths, outputPath) => {
    return new Promise((resolve, reject) => {
        // Ensure output directory exists
        const outputDir = path.dirname(outputPath);
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        const output = fs.createWriteStream(outputPath);
        const archive = archiver('zip', {
            zlib: { level: 9 } // Maximum compression
        });

        output.on('close', () => {
            logger.info(`ZIP file created: ${outputPath} (${archive.pointer()} bytes)`);
            resolve(outputPath);
        });

        archive.on('error', (err) => {
            logger.error(`ZIP creation error: ${err.message}`);
            reject(err);
        });

        archive.pipe(output);

        // Add files to ZIP
        filePaths.forEach(filePath => {
            if (fs.existsSync(filePath)) {
                archive.file(filePath, { name: path.basename(filePath) });
            }
        });

        archive.finalize();
    });
};

// Helper function: Extract Meter Number from filename
const extractMeterNumber = (fileName) => {
    const match = fileName.match(/Load Survey - (\w+)/);
    return match ? match[1].trim() : null;
};

// Helper function: Find client by meter number
const findClientByMeterNumber = async (meterNumber) => {
    let client, meterType, clientType;

    if (meterNumber.startsWith('GJ')) {
        client = await MainClient.findOne({ 'abtMainMeter.meterNumber': meterNumber });
        meterType = client ? 'abtMainMeter' : null;

        if (!client) {
            client = await MainClient.findOne({ 'abtCheckMeter.meterNumber': meterNumber });
            meterType = client ? 'abtCheckMeter' : null;
        }

        clientType = 'MainClient';
    } else if (meterNumber.startsWith('DG')) {
        client = await SubClient.findOne({ 'abtMainMeter.meterNumber': meterNumber });
        meterType = client ? 'abtMainMeter' : null;

        if (!client) {
            client = await SubClient.findOne({ 'abtCheckMeter.meterNumber': meterNumber });
            meterType = client ? 'abtCheckMeter' : null;
        }

        clientType = 'SubClient';
    }

    return client ? { client, clientType, meterType } : null;
};

// Process uploaded CSV/CDF files and extract recipients
exports.processEmailFiles = async (req, res) => {
    try {
        const { sendType, year, month, week } = req.body;
        const files = req.files;

        if (!files || files.length === 0) {
            return res.status(400).json({ message: 'No files uploaded' });
        }

        if (!sendType || !year) {
            return res.status(400).json({ message: 'Send type and year are required' });
        }

        if (sendType === 'monthly' && !month) {
            return res.status(400).json({ message: 'Month is required for monthly emails' });
        }

        if (sendType === 'weekly' && !week) {
            return res.status(400).json({ message: 'Week is required for weekly emails' });
        }

        // For general type, year is optional
        if (sendType === 'general') {
            // No period validation needed
        }

        const processedFiles = [];
        const recipients = [];
        const errors = [];

        for (const file of files) {
            const filePath = path.resolve(file.path);
            
            try {
                // Re-encode CSV to UTF-8
                await reencodeCSVtoUTF8(filePath);

                // Extract meter number from filename
                const meterNumber = extractMeterNumber(file.originalname);
                
                if (!meterNumber) {
                    errors.push({
                        fileName: file.originalname,
                        reason: 'Could not extract meter number from filename'
                    });
                    fs.unlinkSync(filePath);
                    continue;
                }

                // Find client by meter number
                const clientData = await findClientByMeterNumber(meterNumber);
                
                if (!clientData) {
                    errors.push({
                        fileName: file.originalname,
                        reason: `Client not found for meter number: ${meterNumber}`
                    });
                    fs.unlinkSync(filePath);
                    continue;
                }

                const { client } = clientData;

                // Log for debugging
                logger.info(`Found client: ${client.name} (ID: ${client._id}) for meter: ${meterNumber}`);

                // Check if recipient already exists in the list
                const existingRecipient = recipients.find(r => r.meterNumber === meterNumber);
                
                if (existingRecipient) {
                    // Add file to existing recipient
                    existingRecipient.csvCheck = true;
                } else {
                    // Add new recipient (no email validation needed)
                    recipients.push({
                        mainClientId: client._id,
                        mainClientName: client.name,
                        meterNumber: meterNumber,
                        email: client.email || 'No email',
                        consumerNo: client.consumerNo || 'N/A',
                        csvCheck: true,
                        cdfCheck: false,
                        mailSent: false
                    });
                }

                // Upload file to Cloudinary
                const cloudinaryResult = await uploadToCloudinary(filePath, {
                    folder: `samay-electro/email-files/${sendType}/${year}`,
                    resource_type: 'raw' // For CSV/CDF/Excel files
                });

                if (!cloudinaryResult.success) {
                    errors.push({
                        fileName: file.originalname,
                        reason: `Cloudinary upload failed: ${cloudinaryResult.error}`
                    });
                    if (fs.existsSync(filePath)) {
                        fs.unlinkSync(filePath);
                    }
                    continue;
                }

                // Store processed file info with Cloudinary URL
                processedFiles.push({
                    filename: file.filename,
                    originalName: file.originalname,
                    path: file.path, // Keep local path for backward compatibility
                    cloudinaryUrl: cloudinaryResult.url,
                    cloudinaryPublicId: cloudinaryResult.publicId,
                    size: file.size,
                    meterNumber: meterNumber,
                    uploadedAt: new Date()
                });

                // Delete local file after successful Cloudinary upload
                if (fs.existsSync(filePath)) {
                    fs.unlinkSync(filePath);
                    logger.info(`Local file deleted after Cloudinary upload: ${filePath}`);
                }

                logger.info(`Processed file: ${file.originalname} for meter: ${meterNumber}`);

            } catch (error) {
                errors.push({
                    fileName: file.originalname,
                    reason: `Processing error: ${error.message}`
                });
                if (fs.existsSync(filePath)) {
                    fs.unlinkSync(filePath);
                }
            }
        }

        if (recipients.length === 0) {
            return res.status(400).json({
                message: 'No valid recipients found. Make sure meter numbers in filenames match existing clients.',
                errors,
                hint: `Processed ${files.length} file(s), but no clients found for the meter numbers`
            });
        }

        // Create email batch
        const period = {
            year: year ? parseInt(year) : new Date().getFullYear(),
            ...(sendType === 'monthly' ? { month: parseInt(month) } : {}),
            ...(sendType === 'weekly' ? { week: parseInt(week) } : {})
        };

        const emailBatch = new EmailSend({
            sendType,
            period,
            recipients,
            uploadedFiles: processedFiles,
            totalRecipients: recipients.length,
            notes: `${sendType === 'weekly' ? 'Weekly' : 'Monthly'} email batch`,
            createdBy: req.userId,
            status: 'draft'
        });

        await emailBatch.save();

        logger.info(`Email batch created: ${emailBatch._id} with ${recipients.length} recipients`);
        
        res.status(201).json({
            message: 'Files processed successfully',
            batch: emailBatch,
            summary: {
                totalFiles: files.length,
                successful: processedFiles.length,
                failed: errors.length
            },
            errors: errors.length > 0 ? errors : undefined
        });

    } catch (error) {
        logger.error(`Error processing email files: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Create new email batch (legacy - kept for compatibility)
exports.createEmailBatch = async (req, res) => {
    try {
        const { sendType, period, recipients, uploadedFiles, notes } = req.body;

        if (!sendType || !period || !recipients || recipients.length === 0) {
            return res.status(400).json({
                message: 'Send type, period, and recipients are required'
            });
        }

        // Validate period based on sendType
        if (sendType === 'monthly' && (!period.month || !period.year)) {
            return res.status(400).json({
                message: 'Month and year are required for monthly emails'
            });
        }

        if (sendType === 'weekly' && (!period.week || !period.year)) {
            return res.status(400).json({
                message: 'Week and year are required for weekly emails'
            });
        }

        const emailBatch = new EmailSend({
            sendType,
            period,
            recipients,
            uploadedFiles: uploadedFiles || [],
            totalRecipients: recipients.length,
            notes,
            createdBy: req.userId // From auth middleware
        });

        await emailBatch.save();

        logger.info(`Email batch created: ${emailBatch._id}`);
        res.status(201).json({
            message: 'Email batch created successfully',
            batch: emailBatch
        });
    } catch (error) {
        logger.error(`Error creating email batch: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Get all email batches
exports.getAllEmailBatches = async (req, res) => {
    try {
        const { sendType, status, year, month, week } = req.query;
        const query = {};

        if (sendType) query.sendType = sendType;
        if (status) query.status = status;
        if (year) query['period.year'] = parseInt(year);
        if (month) query['period.month'] = parseInt(month);
        if (week) query['period.week'] = parseInt(week);

        logger.info(`Fetching batches with query: ${JSON.stringify(query)}`);

        const batches = await EmailSend.find(query)
            .populate('createdBy', 'name email')
            .sort({ createdAt: -1 })
            .lean();

        logger.info(`Retrieved ${batches.length} email batches`);
        res.status(200).json({ batches });
    } catch (error) {
        logger.error(`Error retrieving email batches: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Get single email batch by ID
exports.getEmailBatchById = async (req, res) => {
    try {
        const { batchId } = req.params;

        if (!mongoose.Types.ObjectId.isValid(batchId)) {
            return res.status(400).json({ message: 'Invalid batch ID format' });
        }

        const batch = await EmailSend.findById(batchId)
            .populate('createdBy', 'name email')
            .populate('recipients.mainClientId', 'name email')
            .lean();

        if (!batch) {
            logger.warn(`Email batch not found: ${batchId}`);
            return res.status(404).json({ message: 'Email batch not found' });
        }

        logger.info(`Retrieved email batch: ${batchId}`);
        res.status(200).json({ batch });
    } catch (error) {
        logger.error(`Error retrieving email batch: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Update recipient status (CSV check, CDF check)
exports.updateRecipientStatus = async (req, res) => {
    try {
        const { batchId, recipientId } = req.params;
        const { csvCheck, cdfCheck } = req.body;

        if (!mongoose.Types.ObjectId.isValid(batchId)) {
            return res.status(400).json({ message: 'Invalid batch ID format' });
        }

        const batch = await EmailSend.findById(batchId);
        if (!batch) {
            return res.status(404).json({ message: 'Email batch not found' });
        }

        const recipient = batch.recipients.id(recipientId);
        if (!recipient) {
            return res.status(404).json({ message: 'Recipient not found' });
        }

        if (csvCheck !== undefined) recipient.csvCheck = csvCheck;
        if (cdfCheck !== undefined) recipient.cdfCheck = cdfCheck;

        await batch.save();

        logger.info(`Recipient status updated: ${recipientId}`);
        res.status(200).json({
            message: 'Recipient status updated',
            recipient
        });
    } catch (error) {
        logger.error(`Error updating recipient status: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Send email to single recipient
exports.sendEmailToRecipient = async (req, res) => {
    try {
        const { batchId, recipientId } = req.params;
        const { subject, body, attachmentPaths } = req.body;

        if (!mongoose.Types.ObjectId.isValid(batchId)) {
            return res.status(400).json({ message: 'Invalid batch ID format' });
        }

        const batch = await EmailSend.findById(batchId);
        if (!batch) {
            return res.status(404).json({ message: 'Email batch not found' });
        }

        const recipient = batch.recipients.id(recipientId);
        if (!recipient) {
            return res.status(404).json({ message: 'Recipient not found' });
        }

        // Get email configuration for this send type to fetch CC emails
        const emailConfig = await EmailConfig.findOne({ configType: batch.sendType })
            .populate('recipients.clients.clientId', '_id name consumerNo email');

        // Find client-specific CC emails from config - match by client name
        let clientCCEmails = [];
        let clientEmail = null;
        
        if (emailConfig && emailConfig.recipients.clients) {
            const configClient = emailConfig.recipients.clients.find(c => {
                // Match by client name (most reliable since meter number leads to same client)
                return c.clientName === recipient.mainClientName;
            });
            
            if (configClient) {
                // Get client email from config
                clientEmail = configClient.email;
                
                // Get CC emails
                if (configClient.ccEmails && configClient.ccEmails.length > 0) {
                    clientCCEmails = configClient.ccEmails.map(cc => cc.email);
                    logger.info(`Found ${clientCCEmails.length} CC email(s) for client ${recipient.mainClientName}: ${clientCCEmails.join(', ')}`);
                }
            }
        }

        // Validate: Must have client email (mandatory)
        if (!clientEmail || clientEmail === 'No email' || clientEmail === '' || !clientEmail.includes('@')) {
            const errorMsg = `No valid email configured for ${recipient.mainClientName}. Please add email in Mail Configuration.`;
            logger.error(errorMsg);
            
            // Update recipient with error
            recipient.error = errorMsg;
            recipient.mailSent = false;
            batch.failureCount += 1;
            await batch.save();
            
            return res.status(400).json({ message: errorMsg });
        }

        // CC emails are optional - no validation needed

        // Create ZIP file for Weekly and Monthly emails
        let attachments = [];
        let zipFilePath = null;
        let tempDownloadedFiles = []; // Track temporary downloaded files
        
        if ((batch.sendType === 'weekly' || batch.sendType === 'monthly') && batch.uploadedFiles && batch.uploadedFiles.length > 0) {
            // Sanitize client name for filename
            const sanitizedClientName = sanitizeFilename(recipient.mainClientName);
            
            // Create ZIP file
            const zipFileName = `${sanitizedClientName}_${batch.sendType}_${batch.period.year}${batch.period.month ? '_' + batch.period.month : ''}${batch.period.week ? '_Week' + batch.period.week : ''}.zip`;
            zipFilePath = path.join(__dirname, '../../../uploads', zipFileName);
            
            try {
                // Download files from Cloudinary to temp location
                const tempDir = path.join(__dirname, '../../../uploads/temp');
                if (!fs.existsSync(tempDir)) {
                    fs.mkdirSync(tempDir, { recursive: true });
                }

                for (const uploadedFile of batch.uploadedFiles) {
                    if (uploadedFile.cloudinaryUrl) {
                        const tempFilePath = path.join(tempDir, uploadedFile.originalName);
                        await downloadFromCloudinary(uploadedFile.cloudinaryUrl, tempFilePath);
                        tempDownloadedFiles.push(tempFilePath);
                    }
                }

                // Create ZIP from downloaded files
                if (tempDownloadedFiles.length > 0) {
                    await createZipFile(tempDownloadedFiles, zipFilePath);
                    
                    // Add ZIP as attachment
                    attachments.push({
                        filename: zipFileName,
                        path: zipFilePath
                    });
                    
                    logger.info(`ZIP file created for ${recipient.mainClientName}: ${zipFileName}`);
                }
            } catch (zipError) {
                logger.error(`Failed to create ZIP file: ${zipError.message}`);
                // Fallback: Download and send individual files
                try {
                    const tempDir = path.join(__dirname, '../../../uploads/temp');
                    if (!fs.existsSync(tempDir)) {
                        fs.mkdirSync(tempDir, { recursive: true });
                    }

                    for (const uploadedFile of batch.uploadedFiles) {
                        if (uploadedFile.cloudinaryUrl) {
                            const tempFilePath = path.join(tempDir, uploadedFile.originalName);
                            await downloadFromCloudinary(uploadedFile.cloudinaryUrl, tempFilePath);
                            attachments.push({
                                filename: uploadedFile.originalName,
                                path: tempFilePath
                            });
                            tempDownloadedFiles.push(tempFilePath);
                        }
                    }
                } catch (downloadError) {
                    logger.error(`Failed to download files: ${downloadError.message}`);
                }
            }
        } else {
            // For other types or no attachments, download individual files from Cloudinary
            if (batch.uploadedFiles && batch.uploadedFiles.length > 0) {
                try {
                    const tempDir = path.join(__dirname, '../../../uploads/temp');
                    if (!fs.existsSync(tempDir)) {
                        fs.mkdirSync(tempDir, { recursive: true });
                    }

                    for (const uploadedFile of batch.uploadedFiles) {
                        if (uploadedFile.cloudinaryUrl) {
                            const tempFilePath = path.join(tempDir, uploadedFile.originalName);
                            await downloadFromCloudinary(uploadedFile.cloudinaryUrl, tempFilePath);
                            attachments.push({
                                filename: uploadedFile.originalName,
                                path: tempFilePath
                            });
                            tempDownloadedFiles.push(tempFilePath);
                        }
                    }
                } catch (downloadError) {
                    logger.error(`Failed to download files: ${downloadError.message}`);
                }
            }
        }

        // Build recipient list: Client email (mandatory) + CC emails (optional)
        const allRecipients = [clientEmail];
        if (clientCCEmails && clientCCEmails.length > 0) {
            allRecipients.push(...clientCCEmails);
        }

        // Send email to all recipients
        const transporter = createTransporter();
        const mailOptions = {
            from: process.env.EMAIL_USER,
            to: allRecipients.join(','), // Client email + CC emails as recipients
            subject: subject || `${batch.sendType === 'weekly' ? 'Weekly' : 'Monthly'} Report - ${recipient.meterNumber}`,
            html: body || `<p>Dear ${recipient.mainClientName},</p><p>Please find attached your ${batch.sendType} report.</p>`,
            attachments
        };

        await transporter.sendMail(mailOptions);

        // Clean up ZIP file and temp downloaded files after sending
        if (zipFilePath && fs.existsSync(zipFilePath)) {
            try {
                fs.unlinkSync(zipFilePath);
                logger.info(`ZIP file deleted after sending: ${zipFilePath}`);
            } catch (cleanupError) {
                logger.error(`Failed to delete ZIP file: ${cleanupError.message}`);
            }
        }

        // Clean up temporary downloaded files
        for (const tempFile of tempDownloadedFiles) {
            if (fs.existsSync(tempFile)) {
                try {
                    fs.unlinkSync(tempFile);
                    logger.info(`Temp file deleted: ${tempFile}`);
                } catch (cleanupError) {
                    logger.error(`Failed to delete temp file: ${cleanupError.message}`);
                }
            }
        }

        // Update recipient status
        recipient.mailSent = true;
        recipient.sentAt = new Date();
        recipient.emailSubject = mailOptions.subject;
        recipient.emailBody = body;
        recipient.ccEmails = clientCCEmails; // Store CC emails used (may be empty)
        recipient.error = null; // Clear any previous error
        recipient.attachments = attachments.map(att => ({
            filename: att.filename,
            path: att.path,
            size: fs.existsSync(att.path) ? fs.statSync(att.path).size : 0
        }));

        batch.successCount += 1;
        await batch.save();

        logger.info(`Email sent to ${allRecipients.length} recipient(s): ${allRecipients.join(', ')}`);
        res.status(200).json({
            message: 'Email sent successfully',
            recipient,
            sentTo: allRecipients
        });
    } catch (error) {
        logger.error(`Error sending email: ${error.message}`);
        
        // Update recipient with error
        try {
            const batch = await EmailSend.findById(req.params.batchId);
            const recipient = batch.recipients.id(req.params.recipientId);
            recipient.error = error.message;
            recipient.mailSent = false;
            batch.failureCount += 1;
            await batch.save();
        } catch (updateError) {
            logger.error(`Error updating recipient error: ${updateError.message}`);
        }

        res.status(500).json({ message: error.message });
    }
};

// Send email to all recipients in batch
exports.sendEmailToAll = async (req, res) => {
    try {
        const { batchId } = req.params;
        const { subject, body, attachmentPaths } = req.body;

        if (!mongoose.Types.ObjectId.isValid(batchId)) {
            return res.status(400).json({ message: 'Invalid batch ID format' });
        }

        const batch = await EmailSend.findById(batchId);
        if (!batch) {
            return res.status(404).json({ message: 'Email batch not found' });
        }

        // Get email configuration for this send type
        const emailConfig = await EmailConfig.findOne({ configType: batch.sendType })
            .populate('recipients.clients.clientId', '_id name consumerNo email');
        
        // Use template from config if available, otherwise use provided or default
        const emailSubject = subject || emailConfig?.template?.subject || `${batch.sendType === 'weekly' ? 'Weekly' : 'Monthly'} Report`;
        const emailBody = body || emailConfig?.template?.body || `<p>Dear Client,</p><p>Please find attached your ${batch.sendType} report.</p>`;

        batch.status = 'processing';
        await batch.save();

        const transporter = createTransporter();
        let successCount = 0;
        let failureCount = 0;

        // Send emails to all recipients
        for (const recipient of batch.recipients) {
            if (recipient.mailSent) {
                continue; // Skip already sent
            }

            try {
                // Find client-specific CC emails from config
                let clientCCEmails = [];
                if (emailConfig && emailConfig.recipients.clients) {
                    const configClient = emailConfig.recipients.clients.find(c => {
                        const clientId = c.clientId?._id || c.clientId;
                        return clientId.toString() === recipient.mainClientId.toString();
                    });
                    
                    if (configClient && configClient.ccEmails && configClient.ccEmails.length > 0) {
                        clientCCEmails = configClient.ccEmails.map(cc => cc.email);
                        logger.info(`Found ${clientCCEmails.length} CC email(s) for client ${recipient.mainClientName}: ${clientCCEmails.join(', ')}`);
                    }
                }

                // Prepare attachments
                const attachments = [];
                if (attachmentPaths && attachmentPaths.length > 0) {
                    for (const filePath of attachmentPaths) {
                        const fullPath = path.join(__dirname, '../../../', filePath);
                        if (fs.existsSync(fullPath)) {
                            attachments.push({
                                filename: path.basename(filePath),
                                path: fullPath
                            });
                        }
                    }
                }

                const mailOptions = {
                    from: process.env.EMAIL_USER,
                    to: recipient.email,
                    cc: clientCCEmails.length > 0 ? clientCCEmails.join(',') : undefined,
                    subject: emailSubject,
                    html: emailBody,
                    attachments
                };

                await transporter.sendMail(mailOptions);

                // Update recipient
                recipient.mailSent = true;
                recipient.sentAt = new Date();
                recipient.emailSubject = mailOptions.subject;
                recipient.emailBody = emailBody;
                recipient.ccEmails = clientCCEmails; // Store CC emails used
                recipient.attachments = attachments.map(att => ({
                    filename: att.filename,
                    path: att.path,
                    size: fs.existsSync(att.path) ? fs.statSync(att.path).size : 0
                }));

                successCount++;
                logger.info(`Email sent to: ${recipient.email}${clientCCEmails.length > 0 ? ` with CC: ${clientCCEmails.join(', ')}` : ''}`);
            } catch (error) {
                recipient.error = error.message;
                failureCount++;
                logger.error(`Failed to send email to ${recipient.email}: ${error.message}`);
            }
        }

        batch.successCount = successCount;
        batch.failureCount = failureCount;
        batch.status = failureCount === 0 ? 'completed' : 'failed';
        await batch.save();

        logger.info(`Bulk email completed: ${successCount} success, ${failureCount} failed`);
        res.status(200).json({
            message: 'Bulk email process completed',
            successCount,
            failureCount,
            batch
        });
    } catch (error) {
        logger.error(`Error sending bulk emails: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Delete email batch
exports.deleteEmailBatch = async (req, res) => {
    try {
        const { batchId } = req.params;

        if (!mongoose.Types.ObjectId.isValid(batchId)) {
            return res.status(400).json({ message: 'Invalid batch ID format' });
        }

        const batch = await EmailSend.findByIdAndDelete(batchId);
        if (!batch) {
            return res.status(404).json({ message: 'Email batch not found' });
        }

        logger.info(`Email batch deleted: ${batchId}`);
        res.status(200).json({ message: 'Email batch deleted successfully' });
    } catch (error) {
        logger.error(`Error deleting email batch: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Send general email (for General Monthly Email tab)
exports.sendGeneralEmail = async (req, res) => {
    try {
        const { clientId, clientName, clientEmail, ccEmails, subject, body, files } = req.body;

        if (!clientEmail) {
            return res.status(400).json({ message: 'Client email is required' });
        }

        if (!subject || !body) {
            return res.status(400).json({ message: 'Subject and body are required' });
        }

        logger.info(`Sending general email to: ${clientName} (${clientEmail})`);

        // Download files from Cloudinary if provided
        const attachments = [];
        const tempDownloadedFiles = [];
        
        if (files && files.length > 0) {
            try {
                const tempDir = path.join(__dirname, '../../../uploads/temp');
                if (!fs.existsSync(tempDir)) {
                    fs.mkdirSync(tempDir, { recursive: true });
                }

                for (const file of files) {
                    if (file.cloudinaryUrl) {
                        const tempFilePath = path.join(tempDir, file.originalName);
                        await downloadFromCloudinary(file.cloudinaryUrl, tempFilePath);
                        attachments.push({
                            filename: file.originalName,
                            path: tempFilePath
                        });
                        tempDownloadedFiles.push(tempFilePath);
                        logger.info(`Downloaded file from Cloudinary: ${file.originalName}`);
                    }
                }
            } catch (downloadError) {
                logger.error(`Failed to download files from Cloudinary: ${downloadError.message}`);
            }
        }

        // Build recipient list: Client email (mandatory) + CC emails (optional)
        const allRecipients = [clientEmail];
        if (ccEmails && ccEmails.length > 0) {
            allRecipients.push(...ccEmails);
        }

        // Send email
        const transporter = createTransporter();
        const mailOptions = {
            from: process.env.EMAIL_USER,
            to: allRecipients.join(','),
            subject: subject,
            html: body,
            attachments
        };

        await transporter.sendMail(mailOptions);

        // Clean up temporary downloaded files
        for (const tempFile of tempDownloadedFiles) {
            if (fs.existsSync(tempFile)) {
                try {
                    fs.unlinkSync(tempFile);
                    logger.info(`Temp file deleted: ${tempFile}`);
                } catch (cleanupError) {
                    logger.error(`Failed to delete temp file: ${cleanupError.message}`);
                }
            }
        }

        logger.info(`General email sent to ${allRecipients.length} recipient(s): ${allRecipients.join(', ')}`);
        
        res.status(200).json({
            message: 'Email sent successfully',
            sentTo: allRecipients,
            clientName: clientName
        });
    } catch (error) {
        logger.error(`Error sending general email: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Upload files for email batch
exports.uploadEmailFiles = async (req, res) => {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ message: 'No files uploaded' });
        }

        const uploadedFiles = req.files.map(file => ({
            filename: file.filename,
            originalName: file.originalname,
            path: file.path,
            size: file.size,
            uploadedAt: new Date()
        }));

        logger.info(`${uploadedFiles.length} files uploaded for email`);
        res.status(200).json({
            message: 'Files uploaded successfully',
            files: uploadedFiles
        });
    } catch (error) {
        logger.error(`Error uploading files: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Get batch with recipients for table display
exports.getBatchRecipients = async (req, res) => {
    try {
        const { batchId } = req.params;

        if (!mongoose.Types.ObjectId.isValid(batchId)) {
            return res.status(400).json({ message: 'Invalid batch ID format' });
        }

        const batch = await EmailSend.findById(batchId)
            .populate('createdBy', 'name email')
            .lean();

        if (!batch) {
            logger.warn(`Email batch not found: ${batchId}`);
            return res.status(404).json({ message: 'Email batch not found' });
        }

        logger.info(`Retrieved batch recipients: ${batchId}`);
        res.status(200).json({ 
            batch,
            recipients: batch.recipients 
        });
    } catch (error) {
        logger.error(`Error retrieving batch recipients: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Save uploaded files for General Email (with year/month tracking)
exports.saveGeneralEmailFiles = async (req, res) => {
    try {
        const { clientId, clientName, year, month } = req.body;

        if (!clientId || !clientName || !year || !month) {
            return res.status(400).json({ message: 'Client ID, name, year, and month are required' });
        }

        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ message: 'No files uploaded' });
        }

        // Upload files to Cloudinary and prepare files array
        const filesData = [];
        const errors = [];

        for (const file of req.files) {
            const filePath = path.resolve(file.path);
            
            try {
                // Upload to Cloudinary
                const cloudinaryResult = await uploadToCloudinary(filePath, {
                    folder: `samay-electro/general-files/${year}/${month}`,
                    resource_type: 'auto' // Auto-detect file type (PDF, images, Excel, etc.)
                });

                if (!cloudinaryResult.success) {
                    errors.push({
                        fileName: file.originalname,
                        reason: `Cloudinary upload failed: ${cloudinaryResult.error}`
                    });
                    // Delete local file
                    if (fs.existsSync(filePath)) {
                        fs.unlinkSync(filePath);
                    }
                    continue;
                }

                // Store file info with Cloudinary URL
                filesData.push({
                    filename: file.filename,
                    originalName: file.originalname,
                    path: file.path, // Keep for backward compatibility
                    cloudinaryUrl: cloudinaryResult.url,
                    cloudinaryPublicId: cloudinaryResult.publicId,
                    size: file.size,
                    mimetype: file.mimetype,
                    uploadedAt: new Date()
                });

                // Delete local file after successful Cloudinary upload
                if (fs.existsSync(filePath)) {
                    fs.unlinkSync(filePath);
                    logger.info(`Local file deleted after Cloudinary upload: ${filePath}`);
                }

                logger.info(`File uploaded to Cloudinary: ${file.originalname}`);
            } catch (uploadError) {
                errors.push({
                    fileName: file.originalname,
                    reason: `Upload error: ${uploadError.message}`
                });
                // Delete local file
                if (fs.existsSync(filePath)) {
                    fs.unlinkSync(filePath);
                }
            }
        }

        if (filesData.length === 0) {
            return res.status(400).json({
                message: 'No files uploaded successfully',
                errors
            });
        }

        // Find existing record or create new one
        let generalEmailFile = await GeneralEmailFile.findOne({
            clientId,
            year: parseInt(year),
            month: parseInt(month)
        });

        if (generalEmailFile) {
            // Add new files to existing record
            generalEmailFile.files.push(...filesData);
            generalEmailFile.clientName = clientName; // Update name if changed
            await generalEmailFile.save();
            logger.info(`Added ${filesData.length} files to existing record for ${clientName} (${year}-${month})`);
        } else {
            // Create new record
            generalEmailFile = new GeneralEmailFile({
                clientId,
                clientName,
                year: parseInt(year),
                month: parseInt(month),
                files: filesData
            });
            await generalEmailFile.save();
            logger.info(`Created new record with ${filesData.length} files for ${clientName} (${year}-${month})`);
        }

        res.status(200).json({
            message: 'Files saved successfully',
            fileCount: generalEmailFile.files.length,
            files: filesData,
            errors: errors.length > 0 ? errors : undefined
        });
    } catch (error) {
        logger.error(`Error saving general email files: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Get uploaded files for a client (by year/month)
exports.getGeneralEmailFiles = async (req, res) => {
    try {
        const { clientId, year, month } = req.query;

        if (!clientId) {
            return res.status(400).json({ message: 'Client ID is required' });
        }

        const query = { clientId };
        if (year) query.year = parseInt(year);
        if (month) query.month = parseInt(month);

        const records = await GeneralEmailFile.find(query)
            .sort({ year: -1, month: -1, createdAt: -1 })
            .lean();

        // Calculate total file count
        const totalFiles = records.reduce((sum, record) => sum + record.files.length, 0);

        logger.info(`Retrieved ${records.length} record(s) with ${totalFiles} file(s) for client ${clientId}`);
        
        res.status(200).json({
            records,
            totalFiles,
            fileCount: totalFiles // For backward compatibility
        });
    } catch (error) {
        logger.error(`Error retrieving general email files: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Get file counts for all clients (for table display)
exports.getAllClientsFileCounts = async (req, res) => {
    try {
        const { year, month } = req.query;

        const query = {};
        if (year) query.year = parseInt(year);
        if (month) query.month = parseInt(month);

        const records = await GeneralEmailFile.find(query)
            .select('clientId clientName year month files')
            .lean();

        // Group by clientId and calculate file counts
        const clientFileCounts = {};
        records.forEach(record => {
            const clientIdStr = record.clientId.toString();
            if (!clientFileCounts[clientIdStr]) {
                clientFileCounts[clientIdStr] = {
                    clientId: record.clientId,
                    clientName: record.clientName,
                    fileCount: 0
                };
            }
            clientFileCounts[clientIdStr].fileCount += record.files.length;
        });

        logger.info(`Retrieved file counts for ${Object.keys(clientFileCounts).length} client(s)`);
        
        res.status(200).json({
            clientFileCounts: Object.values(clientFileCounts)
        });
    } catch (error) {
        logger.error(`Error retrieving client file counts: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Update last email sent timestamp
exports.updateGeneralEmailSent = async (req, res) => {
    try {
        const { clientId, year, month } = req.body;

        if (!clientId || !year || !month) {
            return res.status(400).json({ message: 'Client ID, year, and month are required' });
        }

        const record = await GeneralEmailFile.findOneAndUpdate(
            {
                clientId,
                year: parseInt(year),
                month: parseInt(month)
            },
            {
                lastEmailSentAt: new Date(),
                $inc: { emailSentCount: 1 }
            },
            { new: true }
        );

        if (!record) {
            return res.status(404).json({ message: 'Record not found' });
        }

        logger.info(`Updated email sent timestamp for ${record.clientName} (${year}-${month})`);
        
        res.status(200).json({
            message: 'Email sent timestamp updated',
            record
        });
    } catch (error) {
        logger.error(`Error updating email sent timestamp: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Delete a file from general email files
exports.deleteGeneralEmailFile = async (req, res) => {
    try {
        const { clientId, year, month, filename } = req.body;

        if (!clientId || !year || !month || !filename) {
            return res.status(400).json({ message: 'Client ID, year, month, and filename are required' });
        }

        const record = await GeneralEmailFile.findOne({
            clientId,
            year: parseInt(year),
            month: parseInt(month)
        });

        if (!record) {
            return res.status(404).json({ message: 'Record not found' });
        }

        // Find and remove the file
        const fileIndex = record.files.findIndex(f => f.filename === filename);
        if (fileIndex === -1) {
            return res.status(404).json({ message: 'File not found' });
        }

        const file = record.files[fileIndex];
        
        // Delete physical file
        const fullPath = path.join(__dirname, '../../../', file.path);
        if (fs.existsSync(fullPath)) {
            fs.unlinkSync(fullPath);
        }

        // Remove from database
        record.files.splice(fileIndex, 1);
        await record.save();

        logger.info(`Deleted file ${filename} for ${record.clientName} (${year}-${month})`);
        
        res.status(200).json({
            message: 'File deleted successfully',
            fileCount: record.files.length
        });
    } catch (error) {
        logger.error(`Error deleting general email file: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Download a file from general email files
exports.downloadGeneralEmailFile = async (req, res) => {
    try {
        const { clientId, year, month, filename } = req.query;

        if (!clientId || !year || !month || !filename) {
            return res.status(400).json({ message: 'Client ID, year, month, and filename are required' });
        }

        const record = await GeneralEmailFile.findOne({
            clientId,
            year: parseInt(year),
            month: parseInt(month)
        });

        if (!record) {
            return res.status(404).json({ message: 'Record not found' });
        }

        // Find the file
        const file = record.files.find(f => f.filename === filename);
        if (!file) {
            return res.status(404).json({ message: 'File not found' });
        }

        // Get full file path
        const fullPath = path.join(__dirname, '../../../', file.path);
        
        // Check if file exists
        if (!fs.existsSync(fullPath)) {
            return res.status(404).json({ message: 'File not found on server' });
        }

        // Set headers for download
        res.setHeader('Content-Disposition', `attachment; filename="${file.originalName}"`);
        res.setHeader('Content-Type', file.mimetype);

        // Stream the file
        const fileStream = fs.createReadStream(fullPath);
        fileStream.pipe(res);

        logger.info(`File downloaded: ${file.originalName} for client ${clientId}`);
    } catch (error) {
        logger.error(`Error downloading file: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};
