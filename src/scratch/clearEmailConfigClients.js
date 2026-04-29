const mongoose = require('mongoose');
require('dotenv').config({ path: './.env' }); // adjust path to .env if needed
const EmailConfig = require('../models/v2/emailConfig.model');

// Connect to MongoDB
mongoose.connect(process.env.MONGO_URI || 'mongodb+srv://krupal02:Babu1234@cluster0.pxto7.mongodb.net/samay', {
    useNewUrlParser: true,
    useUnifiedTopology: true
}).then(async () => {
    console.log("Connected to DB");
    
    // Clear all clients from email configs
    const result = await EmailConfig.updateMany({}, { $set: { "recipients.clients": [] } });
    console.log(`Updated ${result.modifiedCount} configs to remove old sub-clients.`);
    
    mongoose.connection.close();
}).catch(err => {
    console.error("Error connecting to DB", err);
    mongoose.connection.close();
});
