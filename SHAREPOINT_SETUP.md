# SharePoint Integration Setup Guide

## Prerequisites

1. **Azure AD App Registration**
2. **SharePoint Online Site**
3. **GitHub Account** (for GitHub Actions)
4. **Streamlit Cloud Account** (optional, for web UI)

---

## Step 1: Azure AD App Registration

### 1.1 Create App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Fill in:
   - **Name**: `Netrace-SharePoint-Integration`
   - **Supported account types**: Single tenant
   - **Redirect URI**: Leave blank for now
5. Click **Register**

### 1.2 Note the IDs

After creation, note these values (you'll need them later):
- **Application (client) ID**
- **Directory (tenant) ID**

### 1.3 Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Description: `GitHub Actions Secret`
4. Expires: 24 months (or as per policy)
5. Click **Add**
6. **COPY THE SECRET VALUE NOW** (you won't see it again!)

### 1.4 Configure API Permissions

1. Go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Select **Application permissions**
5. Add these permissions:
   - `Sites.Read.All`
   - `Sites.ReadWrite.All`
   - `Files.Read.All`
   - `Files.ReadWrite.All`
6. Click **Grant admin consent** (requires admin)

---

## Step 2: Get SharePoint Site ID

### Option A: Using Graph Explorer

1. Go to [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Sign in
3. Run this query:
   ```
   GET https://graph.microsoft.com/v1.0/sites?search=YourSiteName
   ```
4. Find your site in results and copy the `id` value

### Option B: Using PowerShell

```powershell
Install-Module PnP.PowerShell
Connect-PnPOnline -Url "https://yourcompany.sharepoint.com/sites/yoursite" -Interactive
$site = Get-PnPSite
$site.Id
```

---

## Step 3: Configure GitHub Secrets

1. Go to your GitHub repository
2. Navigate to **Settings** > **Secrets and variables** > **Actions**
3. Click **New repository secret**
4. Add these secrets:

| Secret Name | Value | Example |
|------------|-------|---------|
| `SHAREPOINT_CLIENT_ID` | Application (client) ID | `12345678-1234-...` |
| `SHAREPOINT_CLIENT_SECRET` | Client secret value | `abC~1dE2fG...` |
| `SHAREPOINT_TENANT_ID` | Directory (tenant) ID | `87654321-4321-...` |
| `SHAREPOINT_SITE_ID` | SharePoint site ID | `yourcompany.sharepoint.com,...` |
| `SHAREPOINT_BASE_FOLDER` | Base folder path | `/Documents` |
| `SHAREPOINT_MASTER_PATH` | Master Excel path | `/Documents/Master.xlsx` |

---

## Step 4: SharePoint Folder Structure

Create this structure in your SharePoint site:

```
Documents/
├── HH1309/
│   ├── PO001.pdf
│   └── PO002.pdf
├── HH1310/
│   └── PO003.pdf
└── Master.xlsx (will be auto-created/updated)
```

---

## Step 5: Enable GitHub Actions

1. Push your code to GitHub:
   ```bash
   git add .
   git commit -m "Add SharePoint integration"
   git push origin main
   ```

2. Go to **Actions** tab in GitHub
3. Enable workflows if prompted
4. The workflow will run:
   - Every 15 minutes (scheduled)
   - On manual trigger
   - On push to main (for testing)

### Manual Trigger

1. Go to **Actions** tab
2. Select **Auto Import PDFs from SharePoint**
3. Click **Run workflow**
4. Select branch: `main`
5. Click **Run workflow**

---

## Step 6: Streamlit Cloud Deployment (Optional)

### 6.1 Deploy to Streamlit Cloud

1. Go to [Streamlit Cloud](https://share.streamlit.io)
2. Sign in with GitHub
3. Click **New app**
4. Select:
   - Repository: `ahmetkayahub/Netrace`
   - Branch: `main`
   - Main file: `app.py`
5. Click **Deploy**

### 6.2 Configure Secrets

1. In Streamlit Cloud app settings
2. Go to **Secrets** section
3. Paste this (with your actual values):

```toml
[sharepoint]
client_id = "your-client-id"
client_secret = "your-client-secret"
tenant_id = "your-tenant-id"
site_id = "your-site-id"
base_folder = "/Documents"
master_path = "/Documents/Master.xlsx"
enabled = true
```

4. Click **Save**

---

## Testing

### Test Local Scripts

```bash
# Set environment variables
export SHAREPOINT_CLIENT_ID="your-client-id"
export SHAREPOINT_CLIENT_SECRET="your-client-secret"
export SHAREPOINT_TENANT_ID="your-tenant-id"
export SHAREPOINT_SITE_ID="your-site-id"

# Test download
python scripts/download_from_sharepoint.py

# Test upload
python scripts/upload_to_sharepoint.py

# Test full auto import
export USE_SHAREPOINT=true
python scripts/auto_import.py
```

### Test GitHub Actions

1. Manually trigger workflow (see Step 5)
2. Check **Actions** tab for results
3. Download artifacts to verify Master.xlsx

### Test Streamlit Cloud

1. Open your Streamlit app URL
2. Go to **SharePoint Integration** section
3. Click **Download PDFs from SharePoint**
4. Verify files are downloaded
5. Click **Upload Master.xlsx to SharePoint**
6. Check SharePoint for updated file

---

## Troubleshooting

### Authentication Error

**Error**: `Authentication failed`

**Solution**:
- Verify Client ID, Secret, Tenant ID
- Check secret hasn't expired
- Ensure admin consent granted for API permissions

### Site ID Not Found

**Error**: `Failed to list folders: 404`

**Solution**:
- Verify Site ID format
- Check SharePoint site URL
- Ensure app has `Sites.Read.All` permission

### Permission Denied

**Error**: `403 Forbidden`

**Solution**:
- Grant admin consent in Azure AD
- Verify API permissions: `Sites.ReadWrite.All`, `Files.ReadWrite.All`
- Wait 5-10 minutes for permissions to propagate

### GitHub Actions Quota

**Error**: `Workflow run exceeded quota`

**Solution**:
- GitHub Actions free tier: 2,000 minutes/month
- Reduce schedule frequency (e.g., hourly instead of every 15 min)
- Or upgrade to GitHub Pro/Team

---

## Security Best Practices

1. ✅ **Never commit secrets** to Git
2. ✅ Use **minimum required permissions** (principle of least privilege)
3. ✅ **Rotate secrets** regularly (every 90 days recommended)
4. ✅ **Monitor usage** in Azure AD audit logs
5. ✅ Use **managed identities** when possible (Azure-hosted scenarios)

---

## Cost Analysis

| Service | Free Tier | Notes |
|---------|-----------|-------|
| **GitHub Actions** | 2,000 min/month | Workflow runs |
| **Streamlit Cloud** | Unlimited | Public apps |
| **Microsoft Graph API** | Unlimited | No additional cost |
| **Azure AD** | Free tier OK | Basic app registration |

**Total Monthly Cost: $0** (within free tiers)

---

## Support

For issues:
1. Check GitHub Actions logs
2. Review Azure AD sign-in logs
3. Test with Graph Explorer
4. Check SharePoint permissions
