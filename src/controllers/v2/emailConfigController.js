// emailConfigController.js
const EmailConfig = require('../../models/v2/emailConfig.model');
const MainClient = require('../../models/v1/mainClient.model');
const logger = require('../../utils/logger');
const mongoose = require('mongoose');
const fs = require('fs');
const path = require('path');

// Helper function for default email template
const getDefaultEmailTemplate = () => {
    // Get backend URL from environment or use default
    const backendUrl = process.env.BACKEND_URL || 'http://localhost:8000';
    
    // Check which logo file exists (png, jpg, jpeg, svg)
    const logoExtensions = ['png', 'jpg', 'jpeg', 'svg'];
    const publicPath = path.join(__dirname, '../../../public/images');
    
    let logoFileName = 'logo.png'; // default fallback
    
    // Check if public/images directory exists
    if (fs.existsSync(publicPath)) {
        // Find which logo file exists
        for (const ext of logoExtensions) {
            const logoPath = path.join(publicPath, `logo.${ext}`);
            if (fs.existsSync(logoPath)) {
                logoFileName = `logo.${ext}`;
                logger.info(`Logo file detected: ${logoFileName}`);
                break;
            }
        }
    }
    
    const logoUrl = `${backendUrl}/public/images/${logoFileName}`;
    
    return `<div style="font-family: Arial, sans-serif; color: #333;">
<p>PFA</p>

<p>Regards,</p>

<p><strong>Manish Thummar</strong><br/>
Mobile : +91 82385 85535</p>

<div style="margin: 20px 0;">
    <img src="${logoUrl}" alt="Samay Electro Service" style="max-width: 300px;" />
</div>

<p><strong>A-203, 2nd Floor, Dev Residency,<br/>
Near Verachha Co-Op. Bank, Yogichowk, Punagam,<br/>
Surat-395010, Gujarat, India.</strong></p>

<p>E-mail: <a href="mailto:info@samayelectro.com">info@samayelectro.com</a> | <a href="mailto:admin@samayelectro.com">admin@samayelectro.com</a></p>

<p><strong>MSME No.: UDYAM-GJ-22-0293351</strong></p>

<p><strong>GSTIN : 24AJTPT1949D1ZU</strong></p>

<p><strong>Working Hours (IST): 10:00 am to 6:00 pm, Sunday Week off</strong></p>

<p style="color: #4CAF50; font-size: 12px;"><em>please consider the environment before printing this email</em></p>
</div>`;
};

// Helper function to create default config
const createDefaultConfig = (configType) => {
    return new EmailConfig({
        configType,
        template: {
            subject: `${configType === 'weekly' ? 'Weekly' : 'Monthly'} Generation Report`,
            body: getDefaultEmailTemplate(),
            isActive: true
        },
        recipients: { clients: [], ccEmails: [] }
    });
};

// Get all main clients for dropdown
exports.getAllClients = async (req, res) => {
    try {
        const clients = await MainClient.find({}).lean();

        if (!clients || clients.length === 0) {
            logger.warn('No Clients found for email config');
            return res.status(200).json({ clients: [] });
        }

        const formattedClients = clients.map(client => ({
            _id: client._id,
            name: client.name,
            consumerNo: client.abtMainMeter?.meterNumber || 'N/A',
            email: client.email,
            mainClientName: client.name,
            ...client
        }));

        logger.info(`Retrieved ${formattedClients.length} clients for email config`);
        res.status(200).json({ clients: formattedClients });
    } catch (error) {
        logger.error(`Error retrieving clients: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Get email configuration by type (weekly/monthly)
exports.getEmailConfig = async (req, res) => {
    try {
        const { configType } = req.params;

        if (!['weekly', 'monthly', 'general'].includes(configType)) {
            return res.status(400).json({ message: 'Invalid config type. Must be weekly, monthly, or general' });
        }

        let config = await EmailConfig.findOne({ configType })
            .populate('recipients.clients.clientId', 'name consumerNo email')
            .populate('updatedBy', 'name email');

        // If config doesn't exist, create default with professional signature
        if (!config) {
            const defaultConfig = createDefaultConfig(configType);
            await defaultConfig.save();
            config = defaultConfig;
        }

        // Convert to plain object and ensure all fields are included
        const configObj = config.toObject();
        
        // Make sure client email field is included in response
        if (configObj.recipients && configObj.recipients.clients) {
            configObj.recipients.clients = configObj.recipients.clients.map(client => ({
                ...client,
                clientId: client.clientId._id || client.clientId,
                clientName: client.clientName,
                consumerNo: client.consumerNo,
                email: client.email || '', // Ensure email field is always present
                ccEmails: client.ccEmails || []
            }));
        }

        logger.info(`Retrieved email config: ${configType}`);
        res.status(200).json({ config: configObj });
    } catch (error) {
        logger.error(`Error retrieving email config: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Update email configuration
exports.updateEmailConfig = async (req, res) => {
    try {
        const { configType } = req.params;
        const { template, recipients } = req.body;

        if (!['weekly', 'monthly', 'general'].includes(configType)) {
            return res.status(400).json({ message: 'Invalid config type. Must be weekly, monthly, or general' });
        }

        let config = await EmailConfig.findOne({ configType });

        if (!config) {
            // Create new config
            config = createDefaultConfig(configType);
            if (template) {
                config.template = { ...config.template, ...template };
            }
            if (recipients) {
                config.recipients = recipients;
            }
            config.updatedBy = req.userId;
        } else {
            // Update existing config
            if (template) {
                config.template = {
                    ...config.template,
                    ...template
                };
            }
            if (recipients) {
                config.recipients = recipients;
            }
            config.updatedBy = req.userId;
        }

        await config.save();

        logger.info(`Email config updated: ${configType}`);
        res.status(200).json({
            message: 'Email configuration updated successfully',
            config
        });
    } catch (error) {
        logger.error(`Error updating email config: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Add client to configuration
exports.addClientToConfig = async (req, res) => {
    try {
        const { configType } = req.params;
        const { clientId } = req.body;

        if (!['weekly', 'monthly', 'general'].includes(configType)) {
            return res.status(400).json({ message: 'Invalid config type' });
        }

        if (!mongoose.Types.ObjectId.isValid(clientId)) {
            return res.status(400).json({ message: 'Invalid client ID' });
        }

        // Find client
        const client = await MainClient.findById(clientId).lean();
        if (!client) {
            return res.status(404).json({ message: 'Client not found' });
        }

        // Remove email validation - allow clients without email
        // if (!client.email) {
        //     return res.status(400).json({ message: 'Client does not have an email address' });
        // }

        let config = await EmailConfig.findOne({ configType });
        if (!config) {
            config = createDefaultConfig(configType);
        }

        // Check if client already exists
        const exists = config.recipients.clients.some(
            c => c.clientId.toString() === clientId
        );

        if (exists) {
            return res.status(400).json({ message: 'Client already added to configuration' });
        }

        // Add client
        config.recipients.clients.push({
            clientId: client._id,
            clientName: client.name,
            consumerNo: client.abtMainMeter?.meterNumber || 'N/A',
            email: '' // Empty email - user will add manually
        });

        config.updatedBy = req.userId;
        await config.save();

        logger.info(`Client added to ${configType} config: ${client.name}`);
        res.status(200).json({
            message: 'Client added successfully',
            config
        });
    } catch (error) {
        logger.error(`Error adding client to config: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Remove client from configuration
exports.removeClientFromConfig = async (req, res) => {
    try {
        const { configType, clientId } = req.params;

        if (!['weekly', 'monthly', 'general'].includes(configType)) {
            return res.status(400).json({ message: 'Invalid config type' });
        }

        const config = await EmailConfig.findOne({ configType });
        if (!config) {
            return res.status(404).json({ message: 'Configuration not found' });
        }

        // Log before removal
        logger.info(`Attempting to remove client ${clientId} from ${configType} config`);
        logger.info(`Current clients count: ${config.recipients.clients.length}`);

        // Remove client - handle both string and ObjectId comparison
        const initialLength = config.recipients.clients.length;
        config.recipients.clients = config.recipients.clients.filter(
            c => {
                const cId = c.clientId?.toString() || c.clientId;
                const targetId = clientId.toString();
                return cId !== targetId;
            }
        );

        const finalLength = config.recipients.clients.length;
        
        if (initialLength === finalLength) {
            logger.warn(`Client ${clientId} not found in configuration`);
            return res.status(404).json({ message: 'Client not found in configuration' });
        }

        config.updatedBy = req.userId;
        await config.save();

        logger.info(`Client removed from ${configType} config: ${clientId}`);
        logger.info(`Remaining clients count: ${finalLength}`);
        
        res.status(200).json({
            message: 'Client removed successfully',
            config
        });
    } catch (error) {
        logger.error(`Error removing client from config: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Add CC email to specific client
exports.addCCEmailToClient = async (req, res) => {
    try {
        const { configType, clientId } = req.params;
        const { email, name } = req.body;

        if (!['weekly', 'monthly', 'general'].includes(configType)) {
            return res.status(400).json({ message: 'Invalid config type' });
        }

        if (!email) {
            return res.status(400).json({ message: 'Email is required' });
        }

        // Validate email format
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(email)) {
            return res.status(400).json({ message: 'Invalid email format' });
        }

        const config = await EmailConfig.findOne({ configType });
        if (!config) {
            return res.status(404).json({ message: 'Configuration not found' });
        }

        // Find the client
        const client = config.recipients.clients.find(
            c => (c.clientId?._id || c.clientId).toString() === clientId
        );

        if (!client) {
            return res.status(404).json({ message: 'Client not found in configuration' });
        }

        // Initialize ccEmails array if not exists
        if (!client.ccEmails) {
            client.ccEmails = [];
        }

        // Check if email already exists for this client
        const exists = client.ccEmails.some(cc => cc.email === email);
        if (exists) {
            return res.status(400).json({ message: 'Email already added to this client\'s CC list' });
        }

        // Add CC email to client
        client.ccEmails.push({ email, name: name || '' });
        config.updatedBy = req.userId;
        await config.save();

        logger.info(`CC email added to client ${clientId} in ${configType} config: ${email}`);
        res.status(200).json({
            message: 'CC email added successfully',
            config
        });
    } catch (error) {
        logger.error(`Error adding CC email to client: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Remove CC email from specific client
exports.removeCCEmailFromClient = async (req, res) => {
    try {
        const { configType, clientId, email } = req.params;

        if (!['weekly', 'monthly', 'general'].includes(configType)) {
            return res.status(400).json({ message: 'Invalid config type' });
        }

        const config = await EmailConfig.findOne({ configType });
        if (!config) {
            return res.status(404).json({ message: 'Configuration not found' });
        }

        // Find the client
        const client = config.recipients.clients.find(
            c => (c.clientId?._id || c.clientId).toString() === clientId
        );

        if (!client) {
            return res.status(404).json({ message: 'Client not found in configuration' });
        }

        if (!client.ccEmails) {
            return res.status(404).json({ message: 'No CC emails found for this client' });
        }

        // Remove CC email from client
        client.ccEmails = client.ccEmails.filter(cc => cc.email !== email);

        config.updatedBy = req.userId;
        await config.save();

        logger.info(`CC email removed from client ${clientId} in ${configType} config: ${email}`);
        res.status(200).json({
            message: 'CC email removed successfully',
            config
        });
    } catch (error) {
        logger.error(`Error removing CC email from client: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Update CC email for specific client
exports.updateCCEmailForClient = async (req, res) => {
    try {
        const { configType, clientId, oldEmail } = req.params;
        const { email: newEmail, name } = req.body;

        if (!['weekly', 'monthly', 'general'].includes(configType)) {
            return res.status(400).json({ message: 'Invalid config type' });
        }

        if (!newEmail) {
            return res.status(400).json({ message: 'New email is required' });
        }

        // Validate email format
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(newEmail)) {
            return res.status(400).json({ message: 'Invalid email format' });
        }

        const config = await EmailConfig.findOne({ configType });
        if (!config) {
            return res.status(404).json({ message: 'Configuration not found' });
        }

        // Find the client
        const client = config.recipients.clients.find(
            c => (c.clientId?._id || c.clientId).toString() === clientId
        );

        if (!client) {
            return res.status(404).json({ message: 'Client not found in configuration' });
        }

        if (!client.ccEmails) {
            return res.status(404).json({ message: 'No CC emails found for this client' });
        }

        // Find and update the CC email
        const ccEmail = client.ccEmails.find(cc => cc.email === oldEmail);
        if (!ccEmail) {
            return res.status(404).json({ message: 'CC email not found' });
        }

        ccEmail.email = newEmail;
        if (name !== undefined) ccEmail.name = name;

        config.updatedBy = req.userId;
        await config.save();

        logger.info(`CC email updated for client ${clientId} in ${configType} config: ${oldEmail} -> ${newEmail}`);
        res.status(200).json({
            message: 'CC email updated successfully',
            config
        });
    } catch (error) {
        logger.error(`Error updating CC email for client: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Add CC email
exports.addCCEmail = async (req, res) => {
    try {
        const { configType } = req.params;
        const { email, name } = req.body;

        if (!['weekly', 'monthly', 'general'].includes(configType)) {
            return res.status(400).json({ message: 'Invalid config type' });
        }

        if (!email) {
            return res.status(400).json({ message: 'Email is required' });
        }

        // Validate email format
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(email)) {
            return res.status(400).json({ message: 'Invalid email format' });
        }

        let config = await EmailConfig.findOne({ configType });
        if (!config) {
            config = createDefaultConfig(configType);
        }

        // Check if email already exists
        const exists = config.recipients.ccEmails.some(cc => cc.email === email);
        if (exists) {
            return res.status(400).json({ message: 'Email already added to CC list' });
        }

        // Add CC email
        config.recipients.ccEmails.push({ email, name: name || '' });
        config.updatedBy = req.userId;
        await config.save();

        logger.info(`CC email added to ${configType} config: ${email}`);
        res.status(200).json({
            message: 'CC email added successfully',
            config
        });
    } catch (error) {
        logger.error(`Error adding CC email: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Remove CC email
exports.removeCCEmail = async (req, res) => {
    try {
        const { configType, email } = req.params;

        if (!['weekly', 'monthly', 'general'].includes(configType)) {
            return res.status(400).json({ message: 'Invalid config type' });
        }

        const config = await EmailConfig.findOne({ configType });
        if (!config) {
            return res.status(404).json({ message: 'Configuration not found' });
        }

        // Remove CC email
        config.recipients.ccEmails = config.recipients.ccEmails.filter(
            cc => cc.email !== email
        );

        config.updatedBy = req.userId;
        await config.save();

        logger.info(`CC email removed from ${configType} config: ${email}`);
        res.status(200).json({
            message: 'CC email removed successfully',
            config
        });
    } catch (error) {
        logger.error(`Error removing CC email: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Update template
exports.updateTemplate = async (req, res) => {
    try {
        const { configType } = req.params;
        const { subject, body, isActive } = req.body;

        if (!['weekly', 'monthly', 'general'].includes(configType)) {
            return res.status(400).json({ message: 'Invalid config type' });
        }

        let config = await EmailConfig.findOne({ configType });
        if (!config) {
            config = new EmailConfig({
                configType,
                template: { subject: '', body: '', isActive: true },
                recipients: { clients: [], ccEmails: [] }
            });
        }

        if (subject !== undefined) config.template.subject = subject;
        if (body !== undefined) config.template.body = body;
        if (isActive !== undefined) config.template.isActive = isActive;

        config.updatedBy = req.userId;
        await config.save();

        logger.info(`Template updated for ${configType} config`);
        res.status(200).json({
            message: 'Template updated successfully',
            config
        });
    } catch (error) {
        logger.error(`Error updating template: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Reset configuration to default template (useful for testing)
exports.resetConfigToDefault = async (req, res) => {
    try {
        const { configType } = req.params;

        if (!['weekly', 'monthly', 'general'].includes(configType)) {
            return res.status(400).json({ message: 'Invalid config type. Must be weekly, monthly, or general' });
        }

        // Delete existing config
        await EmailConfig.findOneAndDelete({ configType });

        // Create new default config
        const defaultConfig = createDefaultConfig(configType);
        defaultConfig.updatedBy = req.userId;
        await defaultConfig.save();

        logger.info(`Configuration reset to default for ${configType}`);
        res.status(200).json({
            message: 'Configuration reset to default successfully',
            config: defaultConfig
        });
    } catch (error) {
        logger.error(`Error resetting config: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Update client email in configuration
exports.updateClientEmail = async (req, res) => {
    try {
        const { configType, clientId } = req.params;
        const { email } = req.body;

        if (!['weekly', 'monthly', 'general'].includes(configType)) {
            return res.status(400).json({ message: 'Invalid config type' });
        }

        if (!email) {
            return res.status(400).json({ message: 'Email is required' });
        }

        // Basic email validation
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(email)) {
            return res.status(400).json({ message: 'Invalid email format' });
        }

        const config = await EmailConfig.findOne({ configType });
        if (!config) {
            return res.status(404).json({ message: 'Configuration not found' });
        }

        // Find the client
        const client = config.recipients.clients.find(c => {
            const cId = c.clientId?.toString() || c.clientId;
            return cId === clientId.toString();
        });

        if (!client) {
            return res.status(404).json({ message: 'Client not found in configuration' });
        }

        // Update client email
        client.email = email;

        config.updatedBy = req.userId;
        await config.save();

        logger.info(`Client email updated for ${clientId} in ${configType} config: ${email}`);
        res.status(200).json({
            message: 'Client email updated successfully',
            config
        });
    } catch (error) {
        logger.error(`Error updating client email: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};


