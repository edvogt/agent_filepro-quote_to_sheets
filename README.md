# FilePro to Google Workspace Integration
## Implementation Guide

This package contains everything needed to implement the FilePro quotation sync workflow.

## Contents

1. **FilePro_Google_Workspace_Integration_Proposal.docx** - Complete workflow proposal document
2. **filepro_sync.py** - Python sync application (example implementation)
3. **README.md** - This file

## Quick Start

### Prerequisites

- Python 3.8 or higher installed on your Windows server
- Google Workspace account with admin access
- FilePro accounting system configured to export CSV files
- Network access from local server to Google APIs

### Setup Steps

#### 1. Google Cloud Setup (15 minutes)

1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create a new project (e.g., "FilePro Integration")
3. Enable the following APIs:
   - Google Sheets API
   - Google Drive API
4. Create a Service Account:
   - Go to IAM & Admin → Service Accounts
   - Click "Create Service Account"
   - Name it "filepro-sync"
   - Click "Create and Continue"
   - No roles needed at this stage
   - Click "Done"
5. Create Service Account Key:
   - Click on the service account you just created
   - Go to "Keys" tab
   - Click "Add Key" → "Create new key"
   - Choose JSON format
   - Save the downloaded file as `service-account-key.json`

#### 2. Google Drive Setup (5 minutes)

1. Create a folder structure in Google Drive:
   ```
   Quotations/
   ├── Active/
   ├── Approved/
   ├── Archive/
   └── Templates/
   ```

2. Share the "Quotations" folder with your service account:
   - Right-click the folder → Share
   - Paste the service account email (from the JSON file: `client_email`)
   - Give it "Editor" access

3. Get the folder ID:
   - Open the "Active" folder
   - Copy the ID from the URL: `https://drive.google.com/drive/folders/FOLDER_ID_HERE`

#### 3. Python Environment Setup (10 minutes)

On your Windows server, open Command Prompt or PowerShell:

```bash
# Install required packages
pip install gspread google-auth watchdog pandas --break-system-packages

# Create working directory
mkdir C:\FilePro\Sync
cd C:\FilePro\Sync

# Copy the filepro_sync.py script to this directory
# Copy your service-account-key.json to this directory
```

#### 4. Configure the Script (5 minutes)

Edit `filepro_sync.py` and update the CONFIG section:

```python
CONFIG = {
    # Update these paths
    'service_account_file': r'C:\FilePro\Sync\service-account-key.json',
    'google_drive_folder_id': 'YOUR_FOLDER_ID_FROM_STEP_2',
    'export_directory': r'C:\FilePro\Exports\Quotations',
    'archive_directory': r'C:\FilePro\Exports\Archive',
    
    # These can stay as defaults
    'file_pattern': 'QUOTE_*.csv',
    'check_interval': 60,
    'archive_processed': True,
    'sheet_prefix': 'Quote',
    'log_file': 'filepro_sync.log',
    'log_level': 'INFO'
}
```

#### 5. FilePro Export Configuration

Configure FilePro to export quotation data to CSV format:

**Required CSV Format:**
- Filename: `QUOTE_[NUMBER]_[TIMESTAMP].csv`
- Example: `QUOTE_12345_20250124_143022.csv`
- Location: `C:\FilePro\Exports\Quotations\`

**Suggested CSV Columns:**
- Quote Number
- Client Name
- Client Email
- Date Created
- Status
- Item Description
- Quantity
- Unit Price
- Line Total
- Subtotal
- Tax
- Total Amount
- Terms
- Notes

#### 6. Test the Setup

Run the script manually to test:

```bash
cd C:\FilePro\Sync
python filepro_sync.py
```

You should see:
- Log messages indicating successful Google Sheets connection
- Processing of any existing CSV files
- New Google Sheets created in your Drive folder

Press `Ctrl+C` to stop the script.

#### 7. Schedule Automatic Execution (Windows Task Scheduler)

1. Open Task Scheduler
2. Click "Create Basic Task"
3. Name: "FilePro Quotation Sync"
4. Trigger: "When the computer starts"
5. Action: "Start a program"
6. Program/script: `pythonw.exe` (pythonw runs without console window)
7. Add arguments: `C:\FilePro\Sync\filepro_sync.py`
8. Start in: `C:\FilePro\Sync`
9. Finish the wizard
10. Right-click the task → Properties
11. Check "Run whether user is logged on or not"
12. Check "Run with highest privileges"
13. Click OK

The script will now run automatically in the background!

## Usage

### For Team Members

Once set up, the workflow is automatic:

1. **FilePro Export**: When you export a quotation from FilePro, it creates a CSV file
2. **Automatic Sync**: Within 1-2 minutes, a Google Sheet appears in the "Active" folder
3. **Edit**: Make any needed changes directly in Google Sheets
4. **Share**: Click "Share" to send to clients or team members
5. **Organize**: Move completed quotes to "Approved" or "Archive" folders

### Finding Quotations

**By Number**: Search Google Drive for "Quote 12345"
**By Client**: Search for the client name
**By Date**: Filter folders by year/month
**Recent**: Check the "Active" folder

### Version Control

Google Sheets automatically tracks all changes:
- Click "File" → "Version history" → "See version history"
- View who made changes and when
- Restore previous versions if needed

## Troubleshooting

### Script Not Running

1. Check the log file: `C:\FilePro\Sync\filepro_sync.log`
2. Verify Python is installed: `python --version`
3. Check service account permissions in Google Drive
4. Ensure export directory exists and has CSV files

### Authentication Errors

- Verify the service account JSON file is in the correct location
- Check that the folder is shared with the service account email
- Ensure Google Sheets API and Drive API are enabled in Cloud Console

### Files Not Processing

- Check filename matches pattern: `QUOTE_*.csv`
- Verify CSV format is correct (check first few lines)
- Look for error messages in the log file
- Try processing a single file manually

### Performance Issues

- Reduce `check_interval` in CONFIG for faster processing
- Increase interval if hitting API rate limits
- Consider batch processing for large numbers of quotations

## Customization Ideas

### Adding Email Notifications

Install `yagmail` and add to the script:

```python
import yagmail

def send_notification(quote_number, sheet_url):
    yag = yagmail.SMTP('your-email@gmail.com')
    yag.send(
        to='team@company.com',
        subject=f'New Quotation {quote_number}',
        contents=f'View at: {sheet_url}'
    )
```

### Auto-Organize by Status

Modify the script to move sheets to different folders based on status:

```python
def organize_by_status(spreadsheet, status):
    folder_map = {
        'approved': CONFIG['approved_folder_id'],
        'rejected': CONFIG['archive_folder_id'],
        'pending': CONFIG['active_folder_id']
    }
    
    if status.lower() in folder_map:
        # Move to appropriate folder
        drive.move_file(spreadsheet.id, folder_map[status.lower()])
```

### Template-Based Creation

Create a template sheet with formulas and formatting, then copy it:

```python
def create_from_template(quote_number):
    template = client.open_by_key(CONFIG['template_sheet_id'])
    new_sheet = client.copy(
        template.id,
        title=f"Quote {quote_number}",
        folder_id=CONFIG['folder_id']
    )
    return new_sheet
```

## Support

For issues or questions:

1. Check the log file for error messages
2. Review this README for common solutions
3. Consult the main proposal document for workflow details
4. Contact your implementation partner for technical support

## Maintenance

### Regular Tasks

- **Weekly**: Review log file for errors
- **Monthly**: Archive old quotations to free up space
- **Quarterly**: Review and update Google Sheets templates
- **Annually**: Update Python packages and dependencies

### Updating the Script

```bash
cd C:\FilePro\Sync
python -m pip install --upgrade gspread google-auth watchdog pandas
```

## Security Best Practices

1. **Protect Service Account Key**: Never commit to version control or share publicly
2. **Limit Permissions**: Only give Editor access to specific folders
3. **Regular Audits**: Review who has access to quotation folders
4. **Backup**: Google Drive provides automatic backups, but consider additional archival
5. **Access Logs**: Monitor Google Workspace audit logs for unusual activity

## Next Steps

After successful implementation:

1. Train team members on the new workflow
2. Document company-specific processes
3. Create Google Sheets templates with your branding
4. Set up automated archival workflows
5. Consider integrating with other Google Workspace tools (Gmail, Calendar, etc.)

---

**Document Version**: 1.0  
**Last Updated**: January 2025  
**Created For**: Small Professional Company FilePro Integration
