const dns = require('dns');
const mongoose = require('mongoose');
require('dotenv').config();

// Local dev only: Windows DNS may refuse SRV lookups for mongodb+srv URLs.
if (process.env.NODE_ENV !== 'production') {
  dns.setServers(['8.8.8.8', '8.8.4.4']);
}

/**
 * Connects to the MongoDB database using the connection URL from the environment variables.
 */
const connectDB = async () => {
  try {
    await mongoose.connect(process.env.MONGODB_URL, {
      useNewUrlParser: true,
      useUnifiedTopology: true,
    });
    console.log("DB connected successfully");
  } catch (err) {
    console.error("DB connection failed", err);
    // Exit the process with failure code
    process.exit(1);
  }
};

// Attempt to connect to the database with retry logic
connectDB();

// Optional: For retry logic
mongoose.connection.on('disconnected', () => {
  console.log('MongoDB disconnected! Retrying...');
  setTimeout(connectDB, 5000);  // Retry after 5 seconds
});
