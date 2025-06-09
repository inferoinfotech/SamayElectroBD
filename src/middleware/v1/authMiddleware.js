// authMiddleware.js
const jwt = require('jsonwebtoken');
const dotenv = require('dotenv');
const logger = require('../../utils/logger');
dotenv.config();

// Middleware to check if the user is authenticated
exports.verifyToken = (req, res, next) => {
    const token = req.headers['authorization']?.split(' ')[1]; // "Bearer <token>"

    if (!token) {
        logger.warn('No token provided in request.');
        return res.status(403).json({ message: 'No token provided' });
    }

    // Verify the token using the JWT_SECRET
    jwt.verify(token, process.env.JWT_SECRET, (err, decoded) => {
        if (err) {
            if (err.name === 'TokenExpiredError') {
                logger.error(`Token expired for user: ${req.body.userId || 'unknown'}`);
                return res.status(401).json({ message: 'Token expired. Please log in again.' });
            }
            logger.error('Unauthorized access attempt: Invalid token.');
            return res.status(401).json({ message: 'Unauthorized access' });
        }

        // Store the user ID from the token in the request for later use
        req.userId = decoded.id;

        logger.info(`Token verified for user: ${decoded.id}`);

        next(); // Proceed to the next middleware or route handler
    });
};
