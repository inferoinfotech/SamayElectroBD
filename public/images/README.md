# Logo Setup Instructions

## How to Add Your Company Logo

1. **Get your logo file** (PNG, JPG, or SVG format recommended)
   - Recommended size: 300px width or less
   - File name should be: `logo.png` (or `logo.jpg`)

2. **Place the logo file in this folder:**
   ```
   SamayElectroBD/public/images/logo.png
   ```

3. **Update .env file** (if needed for production):
   ```
   BACKEND_URL=https://your-production-backend-url.com
   ```

4. **For local development:**
   - Logo will be accessible at: `http://localhost:5000/public/images/logo.png`
   
5. **For production:**
   - Logo will be accessible at: `https://your-backend-url/public/images/logo.png`

## Current Status
- ✅ Folder created
- ⏳ Waiting for logo file to be added
- ✅ Backend configured to serve static files
- ✅ Email template updated to use dynamic logo URL

## Testing
After adding the logo, test it by visiting:
- Local: http://localhost:5000/public/images/logo.png
- Production: https://your-backend-url/public/images/logo.png
