# Cloud Family Loader - Vercel Integration

This feature allows you to load Revit families from a cloud-based API hosted on Vercel, in addition to loading from local folders.

## Features

- **Dual Mode**: Toggle between Local (folder-based) and Cloud (Vercel API) sources
- **Cloud API**: Fetch family metadata from Vercel serverless functions
- **Auto Download**: Automatically download family files when loading
- **Same UI**: Use the familiar Family Loader interface for both local and cloud families

## Setup Instructions

### 1. Deploy to Vercel

1. Install Vercel CLI (if not already installed):
   ```bash
   npm install -g vercel
   ```

2. Login to Vercel:
   ```bash
   vercel login
   ```

3. Deploy the project:
   ```bash
   cd /path/to/revit-API_t3lab-lite
   vercel
   ```

4. Follow the prompts:
   - Set up and deploy? `Y`
   - Which scope? Choose your account
   - Link to existing project? `N` (first time) or `Y` (if updating)
   - Project name: `revit-family-api` (or your preferred name)
   - In which directory is your code located? `./`

5. After deployment, Vercel will provide you with a URL like:
   ```
   https://your-project-name.vercel.app
   ```

### 2. Configure the Cloud API URL

1. Open the file:
   ```
   T3Lab_Lite.extension/lib/GUI/FamilyLoaderDialog.py
   ```

2. Find and update the `CLOUD_API_URL` variable (around line 60):
   ```python
   # Cloud API configuration
   CLOUD_API_URL = "https://t3stu-dojk2t66r-tien-thanh-trans-projects.vercel.app/api/families"
   ```

3. **✅ Already configured!** Current deployment URL is set.

### 2.1. Disable Vercel Deployment Protection (Required)

The API needs to be publicly accessible for Family Loader to work:

1. Go to https://vercel.com/dashboard
2. Select your project
3. Navigate to **Settings** → **Deployment Protection**
4. **Disable Protection** or set to **Standard** (no password/authentication)
5. Redeploy if needed: `vercel --prod`

Without this step, the API will return "Authentication Required" error.

### 3. Set Up Your Family Data

The API currently returns sample data. You have several options to customize it:

#### Option A: Update the Sample Data (Quick Start)

Edit `api/families.py` and update the `families_data` dictionary with your own family information.

#### Option B: Connect to a Database

Modify `api/families.py` to fetch data from a database (MongoDB, PostgreSQL, etc.):

```python
# Example with MongoDB
from pymongo import MongoClient

def handler(BaseHTTPRequestHandler):
    def do_GET(self):
        client = MongoClient(os.environ['MONGODB_URI'])
        db = client['revit_families']
        families = list(db.families.find())
        # ... process and return
```

#### Option C: Use Cloud Storage

Store your `.rfa` files and thumbnails in cloud storage (AWS S3, Google Cloud Storage, etc.) and update the URLs in your API response.

### 4. Data Structure

The API should return JSON in this format:

```json
{
  "categories": [
    {
      "name": "Doors",
      "path": "Architecture/Doors",
      "families": [
        {
          "name": "Single-Flush Door",
          "fileName": "Single-Flush.rfa",
          "category": "Doors",
          "size": 245760,
          "downloadUrl": "https://your-storage.com/families/doors/Single-Flush.rfa",
          "thumbnailUrl": "https://your-storage.com/thumbnails/doors/Single-Flush.png",
          "description": "Standard single flush door",
          "version": "2024",
          "tags": ["door", "flush", "single"]
        }
      ]
    }
  ],
  "totalFamilies": 10,
  "lastUpdated": "2024-12-24T00:00:00Z"
}
```

### 5. Environment Variables (Optional)

For production, use environment variables for sensitive data:

1. Create a `.env` file (not committed to git):
   ```
   MONGODB_URI=your_mongodb_connection_string
   AWS_ACCESS_KEY=your_aws_key
   AWS_SECRET_KEY=your_aws_secret
   ```

2. Add to `.gitignore`:
   ```
   .env
   .vercel
   ```

3. Set environment variables in Vercel dashboard:
   - Go to your project settings
   - Navigate to "Environment Variables"
   - Add your variables

## Usage

1. Open Revit and launch the Family Loader tool

2. In the top-right corner, you'll see:
   ```
   Source: ○ Local  ○ Cloud (Vercel)
   ```

3. Select **Cloud (Vercel)** to load families from the cloud API

4. Browse categories and select families as usual

5. Click **Load** - the tool will:
   - Download selected families to a temp folder
   - Load them into your Revit project
   - Clean up temp files automatically

## File Locations

- **API Endpoint**: `/api/families.py`
- **Vercel Config**: `/vercel.json`
- **Python Dependencies**: `/requirements.txt`
- **Family Loader Dialog**: `/T3Lab_Lite.extension/lib/GUI/FamilyLoaderDialog.py`
- **UI Definition**: `/T3Lab_Lite.extension/lib/GUI/FamilyLoader.xaml`

## Troubleshooting

### "Failed to connect to cloud API"

- Check your internet connection
- Verify the `CLOUD_API_URL` is correct
- Check Vercel deployment status: `vercel ls`

### "Download failed"

- Ensure `downloadUrl` in your API response is accessible
- Check if the URL requires authentication
- Verify file permissions on cloud storage

### "Invalid file format"

- Ensure downloaded files are valid `.rfa` files
- Check file size (must be between 1KB and 500MB)
- Verify the download completed successfully

## Local Testing

To test the API locally before deploying:

1. Install Vercel CLI
2. Run locally:
   ```bash
   vercel dev
   ```
3. Update `CLOUD_API_URL` in `FamilyLoaderDialog.py`:
   ```python
   CLOUD_API_URL = "http://localhost:3000/api/families"
   ```

## Security Considerations

- **Authentication**: Consider adding API authentication for production
- **Rate Limiting**: Implement rate limiting to prevent abuse
- **CORS**: The API currently allows all origins (`Access-Control-Allow-Origin: *`)
- **HTTPS**: Always use HTTPS URLs for production

## Next Steps

1. Deploy your actual family library to cloud storage
2. Set up a database to manage family metadata
3. Add authentication if needed
4. Implement caching for better performance
5. Add more metadata fields (tags, descriptions, versions)

## Support

For issues or questions:
- Check the Vercel deployment logs: `vercel logs`
- Review the pyRevit console output in Revit
- Check network requests in the system logs
