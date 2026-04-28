"""
Microsoft Outlook Email Sorter
This script connects to a Microsoft Outlook mailbox via Microsoft Graph API,
reads emails from the inbox, and sorts them into subfolders based on sender email address.
"""

import os
import json
import time
from typing import List, Dict, Optional
from msal import ConfidentialClientApplication
from dotenv import load_dotenv
import requests

# Load environment variables
load_dotenv()


class OutlookEmailSorter:
    """Class to handle Outlook mailbox access and email sorting."""
    
    def __init__(self):
        """Initialize the email sorter with configuration from environment variables."""
        self.client_id = os.getenv('CLIENT_ID')
        self.client_secret = os.getenv('CLIENT_SECRET')
        self.tenant_id = os.getenv('TENANT_ID')
        self.mailbox_email = os.getenv('MAILBOX_EMAIL', 'proofofflightZA@publicis.co.za')
        self.graph_endpoint = os.getenv('GRAPH_API_ENDPOINT', 'https://graph.microsoft.com/v1.0')
        
        # Validate configuration
        if not all([self.client_id, self.client_secret, self.tenant_id]):
            raise ValueError(
                "Missing required configuration. Please set CLIENT_ID, CLIENT_SECRET, and TENANT_ID in your .env file"
            )
        
        # Initialize MSAL app
        self.app = ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}"
        )
        
        self.access_token = None
    
    def get_access_token(self) -> str:
        """
        Get an access token using client credentials flow.
        This requires the app to have appropriate permissions in Azure AD.
        """
        if self.access_token:
            return self.access_token
        
        # Request token for Microsoft Graph API
        result = self.app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        
        if "access_token" in result:
            self.access_token = result["access_token"]
            return self.access_token
        else:
            error_msg = result.get("error_description", result.get("error", "Unknown error"))
            raise Exception(f"Failed to acquire token: {error_msg}")
    
    def make_graph_request(self, endpoint: str, method: str = "GET", data: Optional[Dict] = None) -> Dict:
        """
        Make a request to Microsoft Graph API.
        
        Args:
            endpoint: The Graph API endpoint (e.g., '/me/mailFolders/inbox/messages')
            method: HTTP method (GET, POST, PATCH, etc.)
            data: Optional data for POST/PATCH requests
        
        Returns:
            JSON response from the API
        """
        token = self.get_access_token()
        url = f"{self.graph_endpoint}{endpoint}"
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        if method == "GET":
            response = requests.get(url, headers=headers)
        elif method == "POST":
            response = requests.post(url, headers=headers, json=data)
        elif method == "PATCH":
            response = requests.patch(url, headers=headers, json=data)
        elif method == "MOVE":
            response = requests.post(url, headers=headers, json=data)
        else:
            raise ValueError(f"Unsupported HTTP method: {method}")
        
        response.raise_for_status()
        return response.json()
    
    def get_mailbox_id(self) -> str:
        """
        Get the mailbox ID for the specified email address.
        Uses the user's mailbox or shared mailbox.
        """
        # Try to get the mailbox using the email address
        try:
            # For shared mailboxes or other users, use /users/{email}
            endpoint = f"/users/{self.mailbox_email}"
            user_info = self.make_graph_request(endpoint)
            return user_info.get("id", self.mailbox_email)
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                # If user not found, try using 'me' (requires delegated permissions)
                print(f"Warning: Could not find user {self.mailbox_email}. Trying 'me' endpoint...")
                return "me"
            raise
    
    def get_inbox_messages(self, limit: int = 100) -> List[Dict]:
        """
        Get messages from the inbox.
        
        Args:
            limit: Maximum number of messages to retrieve
        
        Returns:
            List of message objects
        """
        mailbox_id = self.get_mailbox_id()
        endpoint = f"/users/{mailbox_id}/mailFolders/inbox/messages?$top={limit}&$orderby=receivedDateTime desc"
        
        try:
            response = self.make_graph_request(endpoint)
            messages = response.get("value", [])
            print(f"Retrieved {len(messages)} messages from inbox")
            return messages
        except requests.exceptions.HTTPError as e:
            print(f"Error retrieving messages: {e}")
            if e.response.status_code == 403:
                print("Permission denied. Ensure your app has 'Mail.Read' and 'Mail.ReadWrite' permissions.")
            raise
    
    def get_or_create_folder(self, folder_name: str, parent_folder_id: str = "inbox") -> str:
        """
        Get or create a folder in the mailbox.
        
        Args:
            folder_name: Name of the folder to create/get
            parent_folder_id: ID of the parent folder (default: "inbox")
        
        Returns:
            Folder ID
        """
        mailbox_id = self.get_mailbox_id()
        
        # First, try to find existing folder
        endpoint = f"/users/{mailbox_id}/mailFolders/{parent_folder_id}/childFolders?$filter=displayName eq '{folder_name}'"
        
        try:
            response = self.make_graph_request(endpoint)
            folders = response.get("value", [])
            
            if folders:
                folder_id = folders[0]["id"]
                print(f"Found existing folder: {folder_name} (ID: {folder_id})")
                return folder_id
        except requests.exceptions.HTTPError:
            pass
        
        # Folder doesn't exist, create it
        endpoint = f"/users/{mailbox_id}/mailFolders/{parent_folder_id}/childFolders"
        folder_data = {
            "displayName": folder_name,
            "isHidden": False
        }
        
        try:
            response = self.make_graph_request(endpoint, method="POST", data=folder_data)
            folder_id = response["id"]
            print(f"Created new folder: {folder_name} (ID: {folder_id})")
            return folder_id
        except requests.exceptions.HTTPError as e:
            print(f"Error creating folder {folder_name}: {e}")
            if e.response.status_code == 403:
                print("Permission denied. Ensure your app has 'Mail.ReadWrite' permission.")
            raise
    
    def move_message_to_folder(self, message_id: str, destination_folder_id: str):
        """
        Move a message to a different folder.
        
        Args:
            message_id: ID of the message to move
            destination_folder_id: ID of the destination folder
        """
        mailbox_id = self.get_mailbox_id()
        endpoint = f"/users/{mailbox_id}/messages/{message_id}/move"
        
        data = {
            "destinationId": destination_folder_id
        }
        
        try:
            self.make_graph_request(endpoint, method="POST", data=data)
            return True
        except requests.exceptions.HTTPError as e:
            print(f"Error moving message {message_id}: {e}")
            return False
    
    def extract_sender_domain(self, sender_email: str) -> str:
        """
        Extract a clean folder name from sender email.
        Uses the domain name or a sanitized version of the email.
        
        Args:
            sender_email: Email address of the sender
        
        Returns:
            Clean folder name
        """
        if not sender_email:
            return "Unknown"
        
        # Remove angle brackets if present
        sender_email = sender_email.strip("<>")
        
        # Extract domain
        if "@" in sender_email:
            domain = sender_email.split("@")[1].lower()
            # Sanitize domain name for folder (remove invalid characters)
            folder_name = domain.replace(".", "_").replace("-", "_")
            return folder_name
        
        return "Unknown"
    
    def sort_emails_by_sender(self, dry_run: bool = False):
        """
        Main function to sort emails from inbox into subfolders based on sender.
        
        Args:
            dry_run: If True, only show what would be done without actually moving emails
        """
        print(f"\n{'='*60}")
        print(f"Outlook Email Sorter")
        print(f"Mailbox: {self.mailbox_email}")
        print(f"Mode: {'DRY RUN' if dry_run else 'LIVE'}")
        print(f"{'='*60}\n")
        
        # Get messages from inbox
        messages = self.get_inbox_messages(limit=100)
        
        if not messages:
            print("No messages found in inbox.")
            return
        
        # Statistics
        stats = {
            "processed": 0,
            "moved": 0,
            "errors": 0,
            "folders_created": set()
        }
        
        # Process each message
        for message in messages:
            try:
                message_id = message["id"]
                subject = message.get("subject", "(No Subject)")
                
                # Get sender email
                sender = message.get("from", {})
                sender_email = sender.get("emailAddress", {}).get("address", "")
                sender_name = sender.get("emailAddress", {}).get("name", "")
                
                if not sender_email:
                    print(f"⚠️  Skipping message '{subject}' - no sender email found")
                    stats["errors"] += 1
                    continue
                
                # Extract folder name from sender
                folder_name = self.extract_sender_domain(sender_email)
                
                # Get or create folder
                folder_id = self.get_or_create_folder(folder_name)
                stats["folders_created"].add(folder_name)
                
                # Move message
                if not dry_run:
                    success = self.move_message_to_folder(message_id, folder_id)
                    if success:
                        print(f"✓ Moved: '{subject}' from {sender_email} → {folder_name}")
                        stats["moved"] += 1
                    else:
                        stats["errors"] += 1
                else:
                    print(f"[DRY RUN] Would move: '{subject}' from {sender_email} → {folder_name}")
                    stats["moved"] += 1
                
                stats["processed"] += 1
                
                # Rate limiting - be nice to the API
                time.sleep(0.5)
                
            except Exception as e:
                print(f"❌ Error processing message: {e}")
                stats["errors"] += 1
        
        # Print summary
        print(f"\n{'='*60}")
        print("Summary:")
        print(f"  Processed: {stats['processed']} messages")
        print(f"  Moved: {stats['moved']} messages")
        print(f"  Errors: {stats['errors']}")
        print(f"  Folders used: {len(stats['folders_created'])}")
        print(f"{'='*60}\n")


def main():
    """Main entry point."""
    import sys
    
    # Check for dry-run flag
    dry_run = "--dry-run" in sys.argv or "-d" in sys.argv
    
    try:
        sorter = OutlookEmailSorter()
        sorter.sort_emails_by_sender(dry_run=dry_run)
    except Exception as e:
        print(f"❌ Fatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
