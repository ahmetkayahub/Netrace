# JSON-Based Auto-Import Setup Guide

## Overview

This guide explains how to set up the **JSON-based auto-import system** that **does NOT require admin consent** for SharePoint API access.

### How It Works

```
User uploads PDF → Streamlit extracts data → Downloads JSON → 
User uploads JSON to SharePoint → GitHub Actions syncs via OneDrive → 
Processes JSON → Updates Master.xlsx → Uploads back to SharePoint
```

### Advantages

✅ **No Admin Consent Required** - Uses OneDrive sync instead of SharePoint API  
✅ **Fully Automatic** - GitHub Actions runs every 15 minutes  
✅ **Simple Setup** - Only requires OneDrive authorization  
✅ **Cost-Free** - Uses GitHub Actions free tier (2,000 min/month)

---

## Prerequisites

1. **OneDrive/SharePoint account** with access to company SharePoint
2. **GitHub account** (for GitHub Actions)
3. **rclone** installed locally (for initial setup)

---

## Step 1: Get OneDrive Token

### 1.1 Install rclone (Local Computer)

**Linux/macOS:**
```bash
curl https://rclone.org/install.sh | sudo bash
```

**Windows:**
Download from https://rclone.org/downloads/

### 1.2 Authorize OneDrive

Run this command on your local computer:

```bash
rclone authorize "onedrive"
```

**What happens:**
1. Browser opens automatically
2. Sign in with your company Microsoft account
3. Grant permissions
4. Terminal shows a long token string

**Copy the token!** It looks like this:
```json
{"access_token":"eyJ0eXAi...", "refresh_token":"M.C5Y_BAY..."}
```

### 1.3 Get Drive ID

After authorization, find your SharePoint drive ID:

```bash
rclone config create sharepoint onedrive

# Follow prompts:
# - Choose drive type: "business" (OneDrive for Business/SharePoint)
# - Paste the token from step 1.2
# - Choose your SharePoint site from the list

# Get drive ID
rclone listdrives sharepoint:
```

Copy the `drive_id` value.

---

## Step 2: Configure GitHub Secrets

1. Go to your GitHub repository: https://github.com/ahmetkayahub/Netrace

2. Navigate to **Settings** > **Secrets and variables** > **Actions**

3. Click **New repository secret**

4. Add these secrets:

| Secret Name | Value | Example |
|------------|-------|---------|
| `ONEDRIVE_TOKEN` | Full token JSON from Step 1.2 | `{"access_token":"eyJ0...", "refresh_token":"M.C5..."}` |
| `ONEDRIVE_DRIVE_ID` | Drive ID from Step 1.3 | `b!abc123def456...` |

---

## Step 3: SharePoint Folder Structure

Create this structure in your SharePoint site:

```
Documents/
├── HH1309/
│   ├── po_data_20240301_120000.json
│   └── po_data_20240302_130000.json
├── HH1310/
│   └── po_data_20240303_140000.json
└── Master.xlsx (auto-created/updated by system)
```

**Important:**
- Folder names become `SahaID` in Excel
- You can use any folder name (not just HH prefix)
- JSON files must have `.json` extension

---

## Step 4: User Workflow

### 4.1 Extract Data from PDF

1. Go to https://netrace.streamlit.app/
2. Upload your PO PDF
3. Click **"📊 Extract & Export JSON"**
4. Review extracted data in the table
5. Click **"📥 Download JSON"**

### 4.2 Upload to SharePoint

1. Open SharePoint in browser
2. Navigate to `Documents/` folder
3. Create or select a SahaID folder (e.g., `HH1309`)
4. Upload the downloaded JSON file

### 4.3 Wait for Auto-Import

- System runs **every 15 minutes**
- Downloads JSON files from SharePoint
- Processes them into Master.xlsx
- Uploads updated Master.xlsx back to SharePoint

**Check progress:**
- GitHub repository → **Actions** tab
- Look for workflow run: "Auto Import JSON from SharePoint"

---

## Step 5: Manual Trigger (Optional)

To trigger import immediately instead of waiting 15 minutes:

1. Go to **Actions** tab in GitHub
2. Select **"Auto Import JSON from SharePoint"**
3. Click **"Run workflow"**
4. Select branch: `main`
5. Click **"Run workflow"** button

---

## Testing

### Test Locally

```bash
cd /home/ahmet/Masaüstü/pdf-highlighter

# Create test data
mkdir -p data/HH1309

# Create sample JSON file
cat > data/HH1309/po_data_test.json << 'EOF'
{
  "po_number": "PO-2024-001",
  "imported_at": "2024-03-01T12:00:00Z",
  "source_file": "test.pdf",
  "lines": [
    {
      "HizmetKodu": "P123456",
      "OrderQty": "10",
      "Price": "150.00",
      "NetValue": "1500.00"
    },
    {
      "HizmetKodu": "P789012",
      "OrderQty": "25",
      "Price": "200.00",
      "NetValue": "5000.00"
    }
  ]
}
EOF

# Run import
python3 scripts/auto_import_json.py

# Check Master.xlsx
# Should have 2 new rows in SahaKayitlari sheet
```

### Test with SharePoint

1. **Upload test JSON to SharePoint:**
   - Create folder: `Documents/TEST/`
   - Upload the JSON file from local test

2. **Configure rclone locally:**
   ```bash
   # Use the same token as GitHub Secrets
   rclone config create sharepoint onedrive
   # Paste ONEDRIVE_TOKEN value when prompted
   ```

3. **Test download:**
   ```bash
   mkdir -p sync_test
   rclone copy sharepoint:Documents/ sync_test/ --include "*.json"
   ls -R sync_test/
   ```

4. **Trigger GitHub Actions workflow**

5. **Check results:**
   - GitHub Actions logs
   - Download artifact: `master-excel-XXX`
   - Verify rows in Excel

---

## Troubleshooting

### Token Expired

**Error:** `401 Unauthorized` in GitHub Actions logs

**Solution:**
1. Re-run `rclone authorize "onedrive"` locally
2. Update `ONEDRIVE_TOKEN` secret in GitHub
3. Trigger workflow again

### No JSON Files Found

**Error:** `Total JSON files downloaded: 0`

**Solution:**
1. Verify JSON files exist in SharePoint `Documents/` folder
2. Check file extension is `.json` (not `.txt`)
3. Verify folder structure: `Documents/SahaID/file.json`

### Master.xlsx Not Uploading

**Error:** `Master.xlsx not found, skipping upload`

**Solution:**
1. Check Python script logs for errors
2. Verify `utils/excel_writer.py` is working
3. Run locally to debug: `python3 scripts/auto_import_json.py`

### rclone Connection Failed

**Error:** `Failed to create file system`

**Solution:**
1. Check `ONEDRIVE_TOKEN` is valid JSON format
2. Ensure no extra spaces or newlines in secret
3. Re-authorize if token is old (>90 days)

---

## JSON File Format

### Required Fields

```json
{
  "po_number": "PO-2024-001",
  "imported_at": "2024-03-01T12:00:00Z",
  "source_file": "example.pdf",
  "lines": [
    {
      "HizmetKodu": "P123456",
      "OrderQty": "10",
      "Price": "150.00",
      "NetValue": "1500.00"
    }
  ]
}
```

### Field Descriptions

| Field | Type | Description | Required |
|-------|------|-------------|----------|
| `po_number` | string | Purchase Order number | Yes |
| `imported_at` | ISO 8601 datetime | When data was extracted | Yes |
| `source_file` | string | Original PDF filename | Yes |
| `lines` | array | List of order line items | Yes |
| `lines[].HizmetKodu` | string | P-code (e.g., P123456) | Yes |
| `lines[].OrderQty` | string | Quantity ordered | No |
| `lines[].Price` | string | Unit price | No |
| `lines[].NetValue` | string | Total value (Qty × Price) | No |

---

## Monitoring

### GitHub Actions Dashboard

**Check workflow status:**
1. Repository → **Actions** tab
2. See recent runs with status (✅ success, ❌ failed)
3. Click on a run to see detailed logs

### Download Results

**Get Master.xlsx from artifacts:**
1. Open successful workflow run
2. Scroll to **Artifacts** section
3. Download `master-excel-XXX.zip`
4. Extract and open `Master.xlsx`

### Schedule

**Default schedule:** Every 15 minutes

**Change schedule:**
Edit `.github/workflows/auto_import_json.yml`:
```yaml
schedule:
  - cron: '0 * * * *'  # Every hour
  # or
  - cron: '0 0 * * *'  # Once per day at midnight
```

---

## Security Best Practices

1. ✅ **Never commit tokens** to Git
2. ✅ **Use GitHub Secrets** for sensitive data
3. ✅ **Rotate tokens** every 90 days (re-authorize with rclone)
4. ✅ **Monitor workflow logs** for suspicious activity
5. ✅ **Limit SharePoint permissions** to specific folders only

---

## Cost Analysis

| Service | Free Tier | Usage | Cost |
|---------|-----------|-------|------|
| **GitHub Actions** | 2,000 min/month | ~5 min/run × 96 runs/day = 480 min/day | **$0** (within limit) |
| **rclone** | Unlimited | Open source | **$0** |
| **OneDrive Storage** | 1 TB (Microsoft 365) | Minimal (JSON + Excel) | **$0** |

**Total Monthly Cost: $0** (assuming normal usage)

---

## Support

### Common Issues

1. **Token not working?** → Re-authorize with `rclone authorize "onedrive"`
2. **Workflow not running?** → Check GitHub Actions is enabled in repo settings
3. **Data not appearing in Excel?** → Check JSON format and Python script logs
4. **SharePoint sync slow?** → Reduce schedule frequency or check file sizes

### Contact

For issues specific to this setup:
1. Check GitHub Actions logs first
2. Review error messages carefully
3. Test locally with `scripts/auto_import_json.py`
4. Compare JSON format with examples above

---

## Advanced: Multiple SharePoint Sites

To sync from multiple SharePoint sites:

```yaml
# .github/workflows/auto_import_json.yml

- name: ⬇️ Download from Site 1
  run: |
    rclone copy sharepoint1:Documents/ data/site1/ --include "*.json"

- name: ⬇️ Download from Site 2
  run: |
    rclone copy sharepoint2:Documents/ data/site2/ --include "*.json"
```

**Add secrets:**
- `ONEDRIVE_TOKEN_SITE1`
- `ONEDRIVE_DRIVE_ID_SITE1`
- `ONEDRIVE_TOKEN_SITE2`
- `ONEDRIVE_DRIVE_ID_SITE2`

---

## Next Steps

1. ✅ Complete Steps 1-4 above
2. ✅ Test with sample JSON file
3. ✅ Train users on PDF → JSON → SharePoint workflow
4. ✅ Monitor first few runs in GitHub Actions
5. ✅ Set calendar reminder to rotate token in 90 days
