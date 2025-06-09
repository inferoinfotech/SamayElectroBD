const User = require('../../models/v1/auth.model');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const sendEmail = require('../../services/v1/emailService');
const crypto = require('crypto');  // To generate OTP
const logger = require('../../utils/logger');
require('dotenv').config();

let otpStore = {};  // Temporary store for OTPs (should be replaced with a database in production)

// Step 1: Register User
exports.register = async (req, res) => {
    try {
        const { email, password, name } = req.body;

        // Check if the user already exists
        const existingUser = await User.findOne({ email });
        if (existingUser) {
            logger.warn(`User registration attempt with existing email: ${email}`);
            return res.status(400).json({ message: "Email already exists." });
        }

        // Hash password
        const hashedPass = await bcrypt.hash(password, 10);
        const user = await User.create({ email, password: hashedPass ,name});

        // Log the successful registration
        logger.info(`User registered successfully: ${email}`);
        res.status(201).json({ message: "User registered successfully", user });
    } catch (error) {
        logger.error(`Error registering user ${req.body.email}: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Step 2: Login User
exports.login = async (req, res) => {
    try {
        // Extract email and password from the request body
        const { email, password } = req.body;

        // Find the user by email
        const user = await User.findOne({ email });
        if (!user) {
            // Log the failed login attempt and return a 404 response
            logger.warn(`Failed login attempt, user not found: ${email}`);
            return res.status(404).json({ message: "User not found" });
        }

        // Compare the hashed password with the one provided by the user
        const isMatch = await bcrypt.compare(password, user.password);
        if (!isMatch) {
            // Log the failed login attempt due to incorrect password and return a 400 response
            logger.warn(`Failed login attempt due to incorrect password: ${email}`);
            return res.status(400).json({ message: "Invalid credentials" });
        }

        // Generate a JWT token that expires in 7 days
        const token = jwt.sign(
            { id: user._id, email: user.email ,name: user.name}, // Include email in token
            process.env.JWT_SECRET,
            { expiresIn: "1h" }
        );

        // Log the successful login event
        logger.info(`User logged in successfully: ${email}`);

        // Respond with the generated JWT token
        res.status(200).json({ message: "Login successful", token });
    } catch (error) {
        // Log any errors that occur during the login process
        logger.error(`Error logging in user ${req.body.email}: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Step 3: Forget Password (Generate OTP and Send Email)
exports.forgetPassword = async (req, res) => {
    try {
        const { email } = req.body;
        const user = await User.findOne({ email });
        if (!user) {
            logger.warn(`Password reset attempt for non-existing user: ${email}`);
            return res.status(404).json({ message: "User not found" });
        }

        // Generate OTP
        const otp = crypto.randomInt(100000, 999999).toString();
        otpStore[email] = otp; // Store OTP temporarily

        // Send OTP via email
        await sendEmail(email, "Password Reset OTP", `Your OTP for password reset is: ${otp}`);
        logger.info(`OTP sent to user for password reset: ${email}`);

        res.status(200).json({ message: "OTP sent to your email." });
    } catch (error) {
        logger.error(`Error sending OTP for password reset ${req.body.email}: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Step 4: Verify OTP
exports.verifyOTP = async (req, res) => {
    try {
        const { email, otp } = req.body;

        // Check if the OTP matches
        if (otpStore[email] !== otp) {
            logger.warn(`Invalid OTP attempt for email: ${email}`);
            return res.status(400).json({ message: "Invalid OTP" });
        }

        // If OTP is valid, remove it from store (can be used only once)
        delete otpStore[email];

        logger.info(`OTP verified successfully for user: ${email}`);
        res.status(200).json({ message: "OTP verified successfully." });
    } catch (error) {
        logger.error(`Error verifying OTP for user ${req.body.email}: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Step 5: Reset Password
exports.resetPassword = async (req, res) => {
    try {
        const { email, newPassword } = req.body;

        const user = await User.findOne({ email });
        if (!user) {
            logger.warn(`User not found during password reset: ${email}`);
            return res.status(404).json({ message: "User not found" });
        }

        // Hash the new password
        const hashedPass = await bcrypt.hash(newPassword, 10);
        user.password = hashedPass;
        await user.save();

        logger.info(`Password updated successfully for user: ${email}`);
        res.status(200).json({ message: "Password updated successfully." });
    } catch (error) {
        logger.error(`Error resetting password for user ${req.body.email}: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};

// Step 6: Reset Password Using Old Password
exports.resetPasswordWithOldPassword = async (req, res) => {
    try {
        // Extract token from headers
        const token = req.headers.authorization?.split(" ")[1]; // "Bearer <token>"
        if (!token) {
            return res.status(401).json({ message: "Unauthorized: No token provided" });
        }

        // Verify token and extract user info
        const decoded = jwt.verify(token, process.env.JWT_SECRET);
        const email = decoded.email; // Extract email from token

        // Find user by extracted email
        const user = await User.findOne({ email });
        if (!user) {
            logger.warn(`User not found during password reset with old password: ${email}`);
            return res.status(404).json({ message: "User not found" });
        }

        const { oldPassword, newPassword } = req.body;

        // Verify old password
        const isMatch = await bcrypt.compare(oldPassword, user.password);
        if (!isMatch) {
            logger.warn(`Incorrect old password attempt for user: ${email}`);
            return res.status(400).json({ message: "Old password is incorrect" });
        }

        // Ensure new password is not the same as old password
        const isSamePassword = await bcrypt.compare(newPassword, user.password);
        if (isSamePassword) {
            logger.warn(`User tried to set the same password: ${email}`);
            return res.status(400).json({ message: "New password must be different from the old password" });
        }

        // Hash the new password
        const hashedPass = await bcrypt.hash(newPassword, 10);
        user.password = hashedPass;
        await user.save();

        logger.info(`Password updated successfully for user: ${email}`);
        res.status(200).json({ message: "Password updated successfully." });
    } catch (error) {
        logger.error(`Error resetting password for user: ${error.message}`);
        res.status(500).json({ message: error.message });
    }
};