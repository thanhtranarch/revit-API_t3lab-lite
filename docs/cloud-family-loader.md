# Cloud Family Loader: Vercel Integration

This feature enables Revit family loading from a cloud-based API hosted on Vercel, providing a centralized alternative to local folder-based storage.

---

## Features

*   **Dual-Source Integration**: Toggle between Local and Cloud sources within a single interface.
*   **Serverless API**: Fetch family metadata dynamically via Vercel functions.
*   **Automated Workflow**: Integrated download and loading process for cloud assets.
*   **Unified UI**: Consistent user experience regardless of the data source.

---

## Deployment & Setup

### 1. Vercel Deployment

1.  **Install Vercel CLI**:
    ```bash
    npm install -g vercel
    ```

2.  **Authentication**:
    ```bash
    vercel login
    ```

3.  **Initialize Deployment**:
    Navigate to the project root and run:
    ```bash
    vercel
    ```

4.  **Configuration Prompts**:
    *   Set up and deploy? `Y`
    *   Link to existing project? `N` (for initial setup)
    *   Project name: `t3lab-family-api`
    *   Directory: `./`

Upon completion, Vercel will provide a production URL (e.g., `https://t3lab-family-api.vercel.app`).

### 2. API Configuration

The Revit plugin must be configured to communicate with your deployment.

1.  **Update Endpoint**:
    Edit `T3Lab.extension/lib/GUI/FamilyLoaderDialog.py` (approx. line 60):
    ```python
    CLOUD_API_URL = "https://your-project.vercel.app/api/families"
    ```

2.  **Deployment Protection (Security)**:
    If Vercel Deployment Protection is enabled, configure a bypass token:
    *   Go to Vercel Dashboard → Settings → Deployment Protection.
    *   Generate a **Bypass Token for Automation**.
    *   Add it to `FamilyLoaderDialog.py` (approx. line 65):
        ```python
        VERCEL_BYPASS_TOKEN = "your-bypass-token"
        ```

---

## Data Management

### 1. Data Structure

The API expects a JSON response structured as follows:

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
          "downloadUrl": "https://storage.provider.com/Single-Flush.rfa",
          "thumbnailUrl": "https://storage.provider.com/Single-Flush.png",
          "version": "2024"
        }
      ]
    }
  ],
  "totalFamilies": 1,
  "lastUpdated": "2024-12-24T00:00:00Z"
}
```

### 2. Customization Options

*   **Static Data**: Update the `families_data` dictionary in `api/families.py`.
*   **Database Integration**: Modify `api/families.py` to fetch from MongoDB or PostgreSQL.
*   **Cloud Storage**: Host `.rfa` files on AWS S3 or Google Cloud Storage.

---

## Usage

1.  Launch the **Family Loader** tool within Revit.
2.  Toggle the source to **Cloud (Vercel)** in the upper-right corner.
3.  Browse the categorized cloud library.
4.  Select assets and click **Load**. The system will:
    *   Download families to a temporary cache.
    *   Load them into the active Revit project.
    *   Perform automatic cleanup of temporary files.

---

## Troubleshooting

| Issue | Resolution |
| :--- | :--- |
| **404 Not Found** | Verify `CLOUD_API_URL` matches your Vercel deployment and endpoint path. |
| **401/403 Forbidden** | Ensure `VERCEL_BYPASS_TOKEN` is correct or disable Deployment Protection. |
| **500 Server Error** | Check Vercel logs using `vercel logs`. |
| **Connection Failed** | Verify internet connectivity and firewall permissions. |
| **Download Failed** | Ensure the `downloadUrl` in the API response is publicly accessible. |

---

## Technical Reference

### File Locations
*   **API Endpoint**: `api/families.py`
*   **Vercel Config**: `vercel.json`
*   **Dependencies**: `requirements.txt`
*   **Plugin Logic**: `T3Lab.extension/lib/GUI/FamilyLoaderDialog.py`

### Security Considerations
*   Always use HTTPS for API communication.
*   Use environment variables in the Vercel dashboard for sensitive credentials.
*   Implement rate limiting for production-scale deployments.
