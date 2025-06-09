const winston = require('winston');
require('winston-daily-rotate-file');

// Set log level based on environment
const logLevel = process.env.NODE_ENV === 'production' ? 'warn' : 'debug';

// Create a Winston logger with different log levels and output formats
const logger = winston.createLogger({
  level: logLevel,  // Default level based on environment
  format: winston.format.combine(
    winston.format.colorize(), // Adds colors to log levels
    winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }), // Timestamp
    winston.format.printf(({ timestamp, level, message }) => {
      return `${timestamp} ${level}: ${message}`;
    })
  ),
  transports: [
    // Console log output
    new winston.transports.Console({
      level: logLevel,
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.simple()
      ),
    }),
    // File log output (in production)
    new winston.transports.DailyRotateFile({
      filename: 'logs/app-%DATE%.log',
      datePattern: 'YYYY-MM-DD',
      level: 'warn',  // Only log warnings and above to file
      maxFiles: '7d',  // Keep logs for 7 days
    }),
  ],
  exceptionHandlers: [
    new winston.transports.Console({
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.simple()
      ),
    }),
    new winston.transports.DailyRotateFile({
      filename: 'logs/exceptions-%DATE%.log',
      datePattern: 'YYYY-MM-DD',
      level: 'error',
      maxFiles: '30d',  // Keep exception logs for 30 days
    }),
  ],
  rejectionHandlers: [
    new winston.transports.Console({
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.simple()
      ),
    }),
    new winston.transports.DailyRotateFile({
      filename: 'logs/rejections-%DATE%.log',
      datePattern: 'YYYY-MM-DD',
      level: 'error',
      maxFiles: '30d',  // Keep rejection logs for 30 days
    }),
  ],
});

// Example usage of logging
logger.info('Application started successfully!');
logger.warn('This is a warning message.');
logger.error('An error occurred.');

// Export the logger
module.exports = logger;
