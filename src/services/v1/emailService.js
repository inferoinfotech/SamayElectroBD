const nodemailer = require("nodemailer");
require("dotenv").config();

/**
 * Utility function for sending emails via SMTP.
 * @param {string} to - Recipient email address.
 * @param {string} subject - Email subject.
 * @param {string} text - Email body content.
 */
const sendEmail = async (to, subject, text) => {
  const transporter = nodemailer.createTransport({
    host: process.env.EMAIL_HOST,
    port: process.env.EMAIL_PORT,
    secure: process.env.EMAIL_PORT === '465', // Use SSL for port 465, otherwise TLS
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS,
    },
  });

  const mailOptions = {
    from: `"Support" <${process.env.EMAIL_USER}>`,
    to,
    subject,
    text,
  };

  try {
    const info = await transporter.sendMail(mailOptions);
    console.log("Email sent: ", info.response);
  } catch (error) {
    if (error.response) {
      console.error("SMTP Error: ", error.response);
    } else {
      console.error("Error sending email: ", error.message);
    }
    throw new Error("Failed to send email");
  }
};

module.exports = sendEmail;
