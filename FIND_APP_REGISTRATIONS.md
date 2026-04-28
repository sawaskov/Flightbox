# How to Find App Registrations in Azure Portal

The Azure Portal interface can vary. Here are multiple ways to find App Registrations:

## Method 1: Using the Search Bar (Easiest)

1. At the **top of the Azure Portal**, you'll see a search bar that says "Search resources, services, and docs"
2. Type: **"App registrations"** (without quotes)
3. Click on **"App registrations"** from the search results
4. You should see a page with a list of applications and a **"+ New registration"** button at the top

## Method 2: Through Microsoft Entra ID (New Name)

Microsoft has rebranded Azure AD to "Microsoft Entra ID":

1. In the Azure Portal, look for **"Microsoft Entra ID"** in the left sidebar
   - It might be under "All services" if not visible
2. Click on **"Microsoft Entra ID"**
3. In the left menu, look for **"App registrations"**
4. Click on it

## Method 3: Direct URL

You can go directly to the App Registrations page:

**https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade**

Just paste this URL in your browser while logged into Azure Portal.

## Method 4: Through All Services

1. Click on **"All services"** in the left sidebar (or the hamburger menu ☰)
2. Search for **"App registrations"** in the services list
3. Click on it

## Method 5: If You See "Identity" Instead

Some organizations see:
1. **"Identity"** in the left sidebar
2. Click on it
3. Look for **"Applications"** or **"App registrations"**

## What You Should See

Once you're on the App Registrations page, you should see:
- A list of existing applications (might be empty)
- A **"+ New registration"** button at the top left
- Or a button that says **"New registration"**

## If You Still Can't Find It

**Possible reasons:**
1. **Permissions**: You might not have permission to view App Registrations
   - You need to be a Global Administrator, Application Administrator, or have custom permissions
   - Contact your IT administrator

2. **Different Tenant**: You might be in a different Azure tenant
   - Check the account you're logged in with
   - Make sure you're in the correct tenant for your organization

3. **UI Update**: Microsoft may have updated the interface
   - Try using the search bar (Method 1) - this usually works regardless of UI changes

## Quick Test

Try this:
1. Click in the search bar at the top
2. Type: **"app reg"**
3. See if "App registrations" appears in the results

If it doesn't appear, you might need to:
- Ask your IT administrator to grant you permissions
- Or have them create the app registration for you
