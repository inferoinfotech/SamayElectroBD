const express = require('express');
const router = express.Router();
const authController = require('../../controllers/v1/authController');

// Register route
router.post('/register', authController.register);

// Login route
router.post('/login', authController.login);

// Forget Password (Send OTP)
router.post('/forget-password', authController.forgetPassword);

// Verify OTP
router.post('/verify-otp', authController.verifyOTP);

// Reset Password
router.post('/reset-password', authController.resetPassword);

// Reset Password Using Old Password
router.post('/reset-password-with-old-password', authController.resetPasswordWithOldPassword);

module.exports = router;
