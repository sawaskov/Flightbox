# Outlook Email Sorter

A Python application that connects to a Microsoft Outlook mailbox via Microsoft Graph API and automatically sorts incoming emails into subfolders based on the sender's email domain.

## Features

- ✅ Connect to Microsoft Outlook mailbox via Microsoft Graph API
- ✅ Read emails from the inbox
- ✅ Automatically create subfolders based on sender email domains
- ✅ Move emails to appropriate folders
- ✅ Dry-run mode to preview changes before executing
- ✅ Comprehensive error handling and logging

## Prerequisites

- Python 3.7 or higher
- Access to the Outlook mailbox: `proofofflightZA@publicis.co.za`
- Azure AD admin access (to register the application)

## Setup Instructions

### Step 1: Install Python Dependencies

```bash
pip install -r requirements.txt
```

### Step 2: Register Application in Azure AD

You need to register an application in Azure Active Directory to get API access credentials.

1. **Go to Azure Portal**: Navigate to [https://portal.azure.com](https://portal.azure.com)

2. **Navigate to Azure Active Directory**:
   - Click on "Azure Active Directory" in the left sidebar
   - Select "App registrations" from the menu
   - Click "New registration"

3. **Register the Application**:
   - **Name**: `Outlook Email Sorter` (or any name you prefer)
   - **Supported account types**: Choose based on your organization's needs
     - For single tenant: "Accounts in this organizational directory only"
     - For multi-tenant: "Accounts in any organizational directory"
   - **Redirect URI**: Leave blank for this application type
   - Click "Register"

4. **Note Down Application Details**:
   - **Application (client) ID**: Copy this value (you'll need it for `CLIENT_ID`)
   - **Directory (tenant) ID**: Copy this value (you'll need it for `TENANT_ID`)

5. **Create a Client Secret**:
   - In your app registration, go to "Certificates & secrets"
   - Click "New client secret"
   - Add a description (e.g., "Email Sorter Secret")
   - Choose expiration period
   - Click "Add"
   - **IMPORTANT**: Copy the secret value immediately (you won't be able to see it again)
   - This is your `CLIENT_SECRET`

6. **Configure API Permissions**:
   - Go to "API permissions" in your app registration
   - Click "Add a permission"
   - Select "Microsoft Graph"
   - Choose "Application permissions" (not Delegated)
   - Add the following permissions:
     - `Mail.Read` - Read mail in all mailboxes
     - `Mail.ReadWrite` - Read and write mail in all mailboxes
     - `User.Read.All` - Read all users' full profiles (if accessing other mailboxes)
   - Click "Add permissions"
   - **IMPORTANT**: Click "Grant admin consent" for your organization
     - This step is required for the permissions to work

### Step 3: Configure Environment Variables

1. **Copy the example configuration file**:
   ```bash
   copy config.example.env .env
   ```
   (On Linux/Mac: `cp config.example.env .env`)

2. **Edit the `.env` file** with your values:
   ```env
   CLIENT_ID=your_application_client_id_from_azure
   CLIENT_SECRET=your_client_secret_value_from_azure
   TENANT_ID=your_tenant_id_from_azure
   MAILBOX_EMAIL=proofofflightZA@publicis.co.za
   GRAPH_API_ENDPOINT=https://graph.microsoft.com/v1.0
   ```

### Step 4: Verify Permissions

Ensure that:
- The application has been granted admin consent for all required permissions
- The mailbox `proofofflightZA@publicis.co.za` exists and is accessible
- Your Azure AD account has permission to access the mailbox

## Usage

### Dry Run (Preview Changes)

Before actually moving emails, test the script in dry-run mode to see what it would do:

```bash
python outlook_email_sorter.py --dry-run
```

or

```bash
python outlook_email_sorter.py -d
```

This will:
- Show which emails would be moved
- Show which folders would be created
- Not actually modify anything

### Run the Email Sorter

Once you're satisfied with the dry-run results, run the script to actually sort emails:

```bash
python outlook_email_sorter.py
```

The script will:
1. Connect to the specified mailbox
2. Retrieve up to 100 messages from the inbox (most recent first)
3. For each email, extract the sender's domain
4. Create a subfolder named after the domain (if it doesn't exist)
5. Move the email to that folder

### Example Output

```
============================================================
Outlook Email Sorter
Mailbox: proofofflightZA@publicis.co.za
Mode: LIVE
============================================================

Retrieved 25 messages from inbox
Found existing folder: example_com (ID: AAMkAG...)
Created new folder: another_domain_org (ID: AAMkAG...)
✓ Moved: 'Important Email' from sender@example.com → example_com
✓ Moved: 'Meeting Request' from contact@another-domain.org → another_domain_org
...

============================================================
Summary:
  Processed: 25 messages
  Moved: 25 messages
  Errors: 0
  Folders used: 8
============================================================
```

## How It Works

1. **Authentication**: Uses Microsoft Authentication Library (MSAL) with client credentials flow to authenticate with Microsoft Graph API.

2. **Email Retrieval**: Fetches emails from the inbox using the Microsoft Graph API `/users/{id}/mailFolders/inbox/messages` endpoint.

3. **Folder Organization**: 
   - Extracts the domain from each sender's email address
   - Creates a folder named after the domain (e.g., `example.com` → `example_com`)
   - If the folder already exists, uses the existing folder

4. **Email Movement**: Moves each email to its corresponding folder using the Graph API's move endpoint.

## Folder Naming Convention

Folders are named based on the sender's email domain:
- `sender@example.com` → folder: `example_com`
- `user@company.co.za` → folder: `company_co_za`

Special characters are replaced with underscores to ensure valid folder names.

## Troubleshooting

### Error: "Failed to acquire token"
- Verify your `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID` are correct
- Ensure the client secret hasn't expired
- Check that admin consent has been granted for the permissions

### Error: "Permission denied" or 403 Forbidden
- Ensure you've granted admin consent for all required permissions
- Verify the permissions are "Application permissions" (not Delegated)
- Check that the mailbox email address is correct

### Error: "User not found" or 404
- Verify the mailbox email address is correct
- Ensure the application has permission to access that mailbox
- For shared mailboxes, you may need additional permissions

### No emails are being retrieved
- Check that the inbox actually contains emails
- Verify the mailbox email address is accessible
- Try increasing the limit in the `get_inbox_messages()` function

## Security Notes

- **Never commit your `.env` file** to version control (it's already in `.gitignore`)
- **Keep your client secret secure** - treat it like a password
- **Rotate secrets regularly** - set expiration dates and update your `.env` file accordingly
- **Use least privilege** - only grant the minimum permissions needed

## Limitations

- Currently processes up to 100 emails per run (configurable in code)
- Uses client credentials flow, which requires application permissions
- Folder names are based on domains, not individual email addresses
- Rate limiting is implemented (0.5 second delay between operations)

## Future Enhancements

Potential improvements:
- Support for custom folder naming rules
- Filter emails by date range
- Support for multiple mailboxes
- Email filtering rules (e.g., only move unread emails)
- Webhook support for real-time sorting
- GUI interface

## Support

For issues related to:
- **Azure AD setup**: Consult [Microsoft Graph API documentation](https://docs.microsoft.com/en-us/graph/)
- **API permissions**: Check [Microsoft Graph permissions reference](https://docs.microsoft.com/en-us/graph/permissions-reference)
- **Script errors**: Review the error messages and check your configuration

## License

This project is provided as-is for internal use.
