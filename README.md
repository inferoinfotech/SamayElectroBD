# Samay Electro - Backend API

Backend API for Samay Electro email management system with Cloudinary integration.

## 🚀 Features

- ✅ Email sending with ZIP attachments (Weekly/Monthly)
- ✅ Cloudinary cloud storage integration
- ✅ Automatic file cleanup
- ✅ JWT authentication
- ✅ MongoDB database
- ✅ Email templates management
- ✅ Client management
- ✅ File upload and processing

## 📋 Prerequisites

- Node.js (v16 or higher)
- MongoDB Atlas account
- Cloudinary account
- Gmail account with App Password

## 🛠️ Installation

1. Clone the repository:
```bash
git clone https://github.com/Kashyap297/samayelectroBD.git
cd samayelectroBD
```

2. Install dependencies:
```bash
npm install
```

3. Create `.env` file (copy from `.env.example`):
```bash
cp .env.example .env
```

4. Update `.env` with your credentials:
```env
MONGODB_URL=your_mongodb_connection_string
PORT=8000
JWT_SECRET=your_secret_key
EMAIL_HOST=smtp.gmail.com
EMAIL_PORT=587
EMAIL_USER=your_email@gmail.com
EMAIL_PASS=your_gmail_app_password
BACKEND_URL=http://localhost:8000
CLOUDINARY_CLOUD_NAME=your_cloud_name
CLOUDINARY_API_KEY=your_api_key
CLOUDINARY_API_SECRET=your_api_secret
```

## 🏃 Running the Application

### Development Mode:
```bash
npm run dev
```

### Production Mode:
```bash
npm start
```

Server will run on: `http://localhost:8000`

## 📦 Dependencies

- **express** - Web framework
- **mongoose** - MongoDB ODM
- **cloudinary** - Cloud storage
- **archiver** - ZIP file creation
- **nodemailer** - Email sending
- **multer** - File upload
- **jsonwebtoken** - JWT authentication
- **bcryptjs** - Password hashing
- **winston** - Logging

## 🌐 API Endpoints

### Authentication
- `POST /api/v1/auth/register` - Register user
- `POST /api/v1/auth/login` - Login user

### Email Send
- `POST /api/v1/email-send/process` - Process and upload files
- `POST /api/v1/email-send/batch/:batchId/recipient/:recipientId/send` - Send email
- `GET /api/v1/email-send/batches` - Get all batches

### Email Config
- `GET /api/v1/email-config/:configType` - Get email configuration
- `POST /api/v1/email-config/:configType/client` - Add client to config
- `PUT /api/v1/email-config/:configType/template` - Update template

## 📁 Project Structure

```
SamayElectroBD/
├── src/
│   ├── config/
│   │   ├── db.js              # Database connection
│   │   └── cloudinary.js      # Cloudinary configuration
│   ├── controllers/
│   │   ├── v1/                # Version 1 controllers
│   │   └── v2/                # Version 2 controllers
│   ├── models/                # Mongoose models
│   ├── routes/                # API routes
│   ├── middleware/            # Custom middleware
│   └── utils/                 # Utility functions
├── uploads/                   # Temporary file storage
├── logs/                      # Application logs
├── .env                       # Environment variables
├── .env.example              # Environment template
├── server.js                 # Entry point
└── package.json              # Dependencies
```

## 🔒 Security

- JWT token authentication
- Password hashing with bcryptjs
- Environment variables for sensitive data
- CORS configuration
- Input validation

## 📝 Logging

Logs are stored in `logs/` directory:
- `app-YYYY-MM-DD.log` - Application logs
- `exceptions-YYYY-MM-DD.log` - Exception logs
- `rejections-YYYY-MM-DD.log` - Rejection logs

## 🚀 Deployment

### Vercel/Railway/Render:

1. Push code to GitHub
2. Connect repository to platform
3. Add environment variables
4. Deploy

### Environment Variables for Production:
```env
MONGODB_URL=your_production_mongodb_url
PORT=8000
JWT_SECRET=strong_secret_key
EMAIL_HOST=smtp.gmail.com
EMAIL_PORT=587
EMAIL_USER=production_email@gmail.com
EMAIL_PASS=production_app_password
BACKEND_URL=https://your-backend-url.com
CLOUDINARY_CLOUD_NAME=your_cloud_name
CLOUDINARY_API_KEY=your_api_key
CLOUDINARY_API_SECRET=your_api_secret
```

## 🐛 Troubleshooting

### Cloudinary Connection Failed:
- Check API credentials in `.env`
- Verify internet connection
- Check Cloudinary dashboard

### Email Not Sending:
- Verify Gmail App Password
- Check EMAIL_USER and EMAIL_PASS
- Enable "Less secure app access" (if needed)

### MongoDB Connection Error:
- Check MONGODB_URL format
- Verify network access in MongoDB Atlas
- Check database user permissions

## 📄 License

ISC

## 👥 Authors

- Kashyap297

## 🔗 Links

- [GitHub Repository](https://github.com/Kashyap297/samayelectroBD)
- [Frontend Repository](https://github.com/Kashyap297/SamayElectroFD)
