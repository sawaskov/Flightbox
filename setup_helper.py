"""
Setup Helper Script
This script helps verify your configuration and test the connection.
"""

import os
import sys
from pathlib import Path

def check_env_file():
    """Check if .env file exists and has required variables."""
    env_path = Path(".env")
    
    if not env_path.exists():
        print("❌ .env file not found!")
        print("\nPlease create a .env file with the following variables:")
        print("  CLIENT_ID=your_client_id")
        print("  CLIENT_SECRET=your_client_secret")
        print("  TENANT_ID=your_tenant_id")
        print("  MAILBOX_EMAIL=proofofflightZA@publicis.co.za")
        return False
    
    print("✓ .env file found")
    
    # Load and check variables
    from dotenv import load_dotenv
    load_dotenv()
    
    required_vars = ['CLIENT_ID', 'CLIENT_SECRET', 'TENANT_ID']
    missing = []
    
    for var in required_vars:
        value = os.getenv(var)
        if not value or value.startswith('your_'):
            missing.append(var)
            print(f"❌ {var} is missing or not configured")
        else:
            print(f"✓ {var} is configured")
    
    mailbox = os.getenv('MAILBOX_EMAIL', 'proofofflightZA@publicis.co.za')
    print(f"✓ Mailbox: {mailbox}")
    
    if missing:
        print(f"\n⚠️  Please configure the missing variables in your .env file")
        return False
    
    return True

def test_connection():
    """Test the connection to Microsoft Graph API."""
    print("\n" + "="*60)
    print("Testing Microsoft Graph API Connection...")
    print("="*60)
    
    try:
        from outlook_email_sorter import OutlookEmailSorter
        
        sorter = OutlookEmailSorter()
        print("✓ OutlookEmailSorter initialized")
        
        # Try to get access token
        token = sorter.get_access_token()
        if token:
            print("✓ Successfully obtained access token")
            print(f"  Token preview: {token[:20]}...")
        else:
            print("❌ Failed to obtain access token")
            return False
        
        # Try to get mailbox info
        try:
            mailbox_id = sorter.get_mailbox_id()
            print(f"✓ Mailbox ID retrieved: {mailbox_id}")
        except Exception as e:
            print(f"⚠️  Could not retrieve mailbox ID: {e}")
            print("   This might be okay if using 'me' endpoint")
        
        print("\n✅ Connection test successful!")
        print("   You can now run: python outlook_email_sorter.py --dry-run")
        return True
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("   Please install dependencies: pip install -r requirements.txt")
        return False
    except Exception as e:
        print(f"❌ Connection test failed: {e}")
        print("\nCommon issues:")
        print("  1. Incorrect CLIENT_ID, CLIENT_SECRET, or TENANT_ID")
        print("  2. Client secret has expired")
        print("  3. Admin consent not granted for permissions")
        print("  4. App permissions not configured correctly")
        return False

def main():
    """Main setup verification."""
    print("="*60)
    print("Outlook Email Sorter - Setup Helper")
    print("="*60)
    print()
    
    # Check if dependencies are installed
    try:
        import msal
        import requests
        from dotenv import load_dotenv
        print("✓ Required packages are installed")
    except ImportError as e:
        print(f"❌ Missing package: {e}")
        print("\nPlease install dependencies:")
        print("  pip install -r requirements.txt")
        return
    
    print()
    
    # Check environment file
    if not check_env_file():
        print("\n" + "="*60)
        print("SETUP INSTRUCTIONS:")
        print("="*60)
        print("\n1. Go to Azure Portal: https://portal.azure.com")
        print("2. Navigate to: Azure Active Directory > App registrations")
        print("3. Click 'New registration'")
        print("4. Fill in the details and register")
        print("5. Go to 'Certificates & secrets' and create a client secret")
        print("6. Go to 'API permissions' and add:")
        print("   - Mail.Read (Application permission)")
        print("   - Mail.ReadWrite (Application permission)")
        print("7. Click 'Grant admin consent'")
        print("8. Copy CLIENT_ID, CLIENT_SECRET, and TENANT_ID to .env file")
        print("\nSee README.md for detailed instructions.")
        return
    
    print()
    
    # Test connection
    if test_connection():
        print("\n🎉 Setup complete! You're ready to sort emails.")
    else:
        print("\n⚠️  Setup incomplete. Please check the errors above.")

if __name__ == "__main__":
    main()
