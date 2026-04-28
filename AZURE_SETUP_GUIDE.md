# Azure AD Setup Guide - Step by Step

Follow these steps to register your application in Azure AD and get the required credentials.

## Step 1: Access Azure Portal

1. Open your web browser
2. Go to: **https://portal.azure.com**
3. Sign in with your Microsoft account (the one that has access to the mailbox)

## Step 2: Navigate to App Registrations

1. In the Azure Portal, click on **"Azure Active Directory"** in the left sidebar
   - If you don't see it, search for "Azure Active Directory" in the top search bar
2. In the left menu, click on **"App registrations"**
3. You should see a list of existing applications (if any)

## Step 3: Register a New Application

1. Click the **"+ New registration"** button at the top
2. Fill in the registration form:
   - **Name**: `Outlook Email Sorter` (or any name you prefer)
   - **Supported account types**: 
     - Select **"Accounts in this organizational directory only"** (Single tenant)
     - OR **"Accounts in any organizational directory"** (Multi-tenant) if you need broader access
   - **Redirect URI**: Leave this section empty (we're using client credentials flow)
3. Click **"Register"** at the bottom

## Step 4: Copy Application Details

After registration, you'll be on the application's **Overview** page. Copy these values:

1. **Application (client) ID** - This is your `CLIENT_ID`
   - Copy this value and save it somewhere safe
2. **Directory (tenant) ID** - This is your `TENANT_ID`
   - Copy this value and save it somewhere safe

**⚠️ Keep these values handy - you'll need them for the .env file**

## Step 5: Create a Client Secret

1. In the left menu of your app registration, click on **"Certificates & secrets"**
2. Under the **"Client secrets"** section, click **"+ New client secret"**
3. Fill in:
   - **Description**: `Email Sorter Secret` (or any description)
   - **Expires**: Choose an expiration period (e.g., 24 months)
4. Click **"Add"**
5. **⚠️ IMPORTANT**: Copy the **Value** of the secret immediately
   - This is your `CLIENT_SECRET`
   - **You won't be able to see this value again after you leave this page!**
   - Save it in a secure location

## Step 6: Configure API Permissions

1. In the left menu, click on **"API permissions"**
2. You'll see a list of permissions (might be empty or have default permissions)
3. Click **"+ Add a permission"**
4. Select **"Microsoft Graph"**
5. Select **"Application permissions"** (NOT "Delegated permissions")
   - This is important! Application permissions allow the app to access mailboxes without user interaction
6. In the search box, type "Mail" and find:
   - **Mail.Read** - Check this box
   - **Mail.ReadWrite** - Check this box
7. Scroll down and click **"Add permissions"**

## Step 7: Grant Admin Consent

**This is a critical step!** Without admin consent, the permissions won't work.

1. Still on the **"API permissions"** page
2. Look for a yellow banner that says **"Grant admin consent for [Your Organization]"**
3. Click the **"Grant admin consent for [Your Organization]"** button
4. Confirm by clicking **"Yes"** if prompted
5. Wait a few seconds for the status to update
6. You should see green checkmarks (✓) next to the permissions with status "Granted for [Your Organization]"

**If you don't have admin rights:**
- You'll need to ask your Azure AD administrator to grant consent
- Or use a different account that has admin privileges

## Step 8: Verify Permissions

Double-check that:
- ✅ Both `Mail.Read` and `Mail.ReadWrite` show as **"Granted for [Your Organization]"**
- ✅ The status column shows green checkmarks
- ✅ Type is **"Application"** (not Delegated)

## Step 9: Create Your .env File

Now you have all three values:
- `CLIENT_ID` - From Step 4
- `CLIENT_SECRET` - From Step 5
- `TENANT_ID` - From Step 4

Create a file named `.env` in your project folder with this content:

```env
CLIENT_ID=paste_your_client_id_here
CLIENT_SECRET=paste_your_client_secret_here
TENANT_ID=paste_your_tenant_id_here
MAILBOX_EMAIL=proofofflightZA@publicis.co.za
GRAPH_API_ENDPOINT=https://graph.microsoft.com/v1.0
```

Replace the placeholder values with your actual values.

## Step 10: Test Your Setup

Run the setup helper script to verify everything is configured correctly:

```bash
python setup_helper.py
```

This will:
- Check if your .env file is configured
- Test the connection to Microsoft Graph API
- Verify that authentication works

## Troubleshooting

### "Failed to acquire token"
- Double-check your CLIENT_ID, CLIENT_SECRET, and TENANT_ID
- Make sure there are no extra spaces or quotes in the .env file
- Verify the client secret hasn't expired

### "Permission denied" or 403 Forbidden
- Ensure admin consent was granted (Step 7)
- Verify permissions are "Application permissions" (not Delegated)
- Check that both Mail.Read and Mail.ReadWrite are granted

### "User not found"
- Verify the mailbox email address is correct
- Ensure the application has permission to access that specific mailbox
- For shared mailboxes, you may need additional configuration

## Next Steps

Once setup is complete:
1. Test with dry-run: `python outlook_email_sorter.py --dry-run`
2. Run the sorter: `python outlook_email_sorter.py`

## Need Help?

If you encounter issues:
1. Check the error message carefully
2. Verify each step was completed correctly
3. Ensure admin consent was granted
4. Check that the client secret hasn't expired
