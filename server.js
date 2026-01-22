const express = require('express');
const cors = require('cors');
require('dotenv').config();
require('./src/config/db');  // Database connection setup

const authRoutes = require('./src/routes/v1/authRoutes');
const mainClientRoutes = require('./src/routes/v1/mainClientRoutes');  // Main Client routes
const subClientRoutes = require('./src/routes/v1/subClientRoutes');
const partClientRoutes = require('./src/routes/v1/partClientRoutes');  // Part Client routes
const meterDataRoutes = require('./src/routes/v1/meterDataRoutes');  // Meter Data routes
const lossesCalculationRoutes = require('./src/routes/v1/lossesCalculationRoutes');  // Losses Calculation routes
const loggerDataRoutes = require('./src/routes/v1/loggerDataRoutes');  // Logger Data routes
const dailyReportRoutes = require('./src/routes/v1/dailyReportRoutes');  // Daily Report routes
const totalReportRoutes = require('./src/routes/v1/totalReportRoutes');  // Total Report routes
const clientProgressRoutes = require('./src/routes/v1/clientProgressRoutes');  // Client Progress routes
const totalDataYearlyRoutes = require('./src/routes/v1/totalDataYearlyRoutes');  // Total Data Yearly routes
const ClientProgressField = require('./src/routes/v1/ClientProgressFIledRoutes')

// V2 Routes
const policyRoutes = require('./src/routes/v2/policyRoutes');  // Policy routes
const clientPolicyRoutes = require('./src/routes/v2/clientPolicyRoutes');  // Client Policy routes
const calculationInvoiceRoutes = require('./src/routes/v2/calculationInvoiceRoutes');  // Calculation Invoice routes

const app = express();
app.use(express.json());  // Body parsing middleware
const allowedOrigins = ["http://localhost:5173", "https://samayelectro.vercel.app", "https://www.samayelectro.com", "https://samay-electro-fd-pd63.vercel.app"];

app.get('/', (req, res) => {
  res.send('API is running...');
})

app.use(
  cors({
    origin: allowedOrigins,
    credentials: true,
    methods: "GET,POST,PUT,DELETE, PATCH",
    allowedHeaders: "Content-Type,Authorization",
  })
);
// Routes
app.use('/api/v1/auth', authRoutes);  // Authentication routes
app.use('/api/v1/mainClient', mainClientRoutes);  // Main Client routes
app.use('/api/v1/subClient', subClientRoutes);
app.use('/api/v1/partClient', partClientRoutes);  // Part Client routes
app.use('/api/v1/meter-data',meterDataRoutes);  // Meter Data routes);
app.use('/api/v1/losses-calculation', lossesCalculationRoutes);  // Losses Calculation routes
app.use('/api/v1/logger-data', loggerDataRoutes);  // Losses Calculation routes
app.use('/api/v1/daily-report',dailyReportRoutes);  // Daily Report routes
app.use('/api/v1/total-report', totalReportRoutes);  // Total Report routes
app.use('/api/v1/client-progress', clientProgressRoutes);
app.use('/api/v1/client-progress-filed', ClientProgressField);
app.use('/api/v1/totalDataYearily',totalDataYearlyRoutes);

// V2 Routes
app.use('/api/v1/policy', policyRoutes);  // Policy routes
app.use('/api/v1/client-policy', clientPolicyRoutes);  // Client Policy routes
app.use('/api/v1/calculation-invoice', calculationInvoiceRoutes);  // Calculation Invoice routes

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
