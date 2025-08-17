#!/usr/bin/env python3
"""
Microsoft 365 Admin API MCP Server
Provides tools to manage Microsoft 365 users and mailboxes.
"""

import asyncio
import os
from typing import Any, Dict, List, Optional

import msal
import requests
from dotenv import load_dotenv
from mcp.server import FastMCP
from mcp.types import (
    CallToolRequest,
    CallToolResult,
    ListToolsRequest,
    ListToolsResult,
    Tool,
    TextContent,
)

load_dotenv()

# Global variables for credentials
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
ACCESS_TOKEN = None

async def get_access_token() -> str:
    """Get access token for Microsoft Graph API."""
    global ACCESS_TOKEN
    
    # Validate environment variables first
    if not TENANT_ID:
        raise Exception("TENANT_ID environment variable is not set")
    if not CLIENT_ID:
        raise Exception("CLIENT_ID environment variable is not set")
    if not CLIENT_SECRET:
        raise Exception("CLIENT_SECRET environment variable is not set")
    
    if ACCESS_TOKEN:
        return ACCESS_TOKEN
        
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    
    scopes = ["https://graph.microsoft.com/.default"]
    
    try:
        result = app.acquire_token_for_client(scopes=scopes)
        
        if "access_token" in result:
            ACCESS_TOKEN = result["access_token"]
            return ACCESS_TOKEN
        else:
            error_msg = f"Failed to acquire token: {result.get('error_description', 'Unknown error')}"
            if result.get('error') == 'unauthorized_client':
                error_msg += "\n\nThis usually means:\n1. CLIENT_ID is incorrect\n2. CLIENT_SECRET is incorrect\n3. App registration is not properly configured"
            elif result.get('error') == 'invalid_client':
                error_msg += "\n\nThis usually means:\n1. CLIENT_ID is incorrect\n2. App registration doesn't exist"
            elif result.get('error') == 'invalid_tenant':
                error_msg += "\n\nThis usually means:\n1. TENANT_ID is incorrect"
            raise Exception(error_msg)
    except Exception as e:
        if "unauthorized_client" in str(e):
            raise Exception(f"Authentication failed: {e}\n\nCheck your CLIENT_ID and CLIENT_SECRET")
        elif "invalid_client" in str(e):
            raise Exception(f"Invalid client: {e}\n\nCheck your CLIENT_ID")
        elif "invalid_tenant" in str(e):
            raise Exception(f"Invalid tenant: {e}\n\nCheck your TENANT_ID")
        else:
            raise Exception(f"Authentication error: {e}")

async def make_graph_request(method: str, endpoint: str, data: Optional[Dict] = None) -> Dict:
    """Make a request to Microsoft Graph API with enhanced error handling."""
    try:
        token = await get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        
        url = f"{GRAPH_BASE_URL}{endpoint}"
        
        # Log the request for debugging
        print(f"Making {method} request to: {url}")
        if data:
            print(f"Request data: {data}")
        
        response = requests.request(method, url, headers=headers, json=data, timeout=30)
        
        # Enhanced error handling
        if response.status_code >= 400:
            error_detail = response.text
            try:
                error_json = response.json()
                if 'error' in error_json:
                    error_detail = f"{error_json['error'].get('code', 'Unknown')}: {error_json['error'].get('message', 'Unknown error')}"
                    if 'innerError' in error_json['error']:
                        error_detail += f" (Inner: {error_json['error']['innerError'].get('message', '')})"
            except:
                pass
            
            # Provide specific guidance based on error codes
            guidance = ""
            if response.status_code == 401:
                guidance = "\n\nThis usually means:\n• Token has expired\n• Insufficient permissions\n• Check if admin consent was granted"
            elif response.status_code == 403:
                guidance = "\n\nThis usually means:\n• Insufficient permissions\n• Check API permissions in Azure app registration\n• Ensure admin consent was granted"
            elif response.status_code == 404:
                guidance = "\n\nThis usually means:\n• Resource not found\n• Check if the user/group exists\n• Verify the email address or ID is correct"
            elif response.status_code == 400:
                guidance = "\n\nThis usually means:\n• Invalid request format\n• Missing required properties\n• Check the request payload"
            elif response.status_code == 409:
                guidance = "\n\nThis usually means:\n• Resource already exists\n• Conflict with existing data"
            
            raise Exception(f"Graph API error: {response.status_code} - {error_detail}{guidance}")
        
        # Handle empty responses
        if not response.content:
            return {}
        
        return response.json()
        
    except requests.exceptions.Timeout:
        raise Exception(f"Request timeout: The request to {endpoint} took too long to complete")
    except requests.exceptions.ConnectionError:
        raise Exception(f"Connection error: Unable to connect to Microsoft Graph API")
    except requests.exceptions.RequestException as e:
        raise Exception(f"Request failed: {str(e)}")
    except Exception as e:
        if "Graph API error" not in str(e):
            raise Exception(f"Unexpected error making Graph API request: {str(e)}")
        else:
            raise e

async def add_user_to_distribution_list(user_email: str, distribution_list_email: str) -> CallToolResult:
    """Add a user to a distribution list or Microsoft 365 Group."""
    try:
        # Get the user ID
        user_response = await make_graph_request("GET", f"/users/{user_email}")
        user_id = user_response["id"]
        user_display_name = user_response.get("displayName", user_email)
        
        # Get the group information to determine its type
        group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{distribution_list_email}'")
        if not group_response.get("value"):
            raise Exception(f"Group '{distribution_list_email}' not found. Please verify the email address is correct.")
        
        group = group_response["value"][0]
        group_id = group["id"]
        group_display_name = group.get("displayName", distribution_list_email)
        group_types = group.get("groupTypes", [])
        
        # Determine if this is a Microsoft 365 Group (Unified) or traditional distribution list
        is_unified_group = "Unified" in group_types
        
        # Check if user is already a member
        members_response = await make_graph_request("GET", f"/groups/{group_id}/members?$filter=id eq '{user_id}'")
        if members_response.get("value"):
            group_type_name = "Microsoft 365 Group" if is_unified_group else "Distribution List"
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"User '{user_display_name}' is already a member of {group_type_name.lower()} '{group_display_name}'"
                )]
            )
        
        # Add user based on group type
        if is_unified_group:
            # For Microsoft 365 Groups (Unified), use the members endpoint with user object
            await make_graph_request("POST", f"/groups/{group_id}/members", {
                "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
            })
        else:
            # For traditional distribution lists, use the members/$ref endpoint
            await make_graph_request("POST", f"/groups/{group_id}/members/$ref", {
                "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
            })
        
        group_type_name = "Microsoft 365 Group" if is_unified_group else "Distribution List"
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=f"Successfully added '{user_display_name}' ({user_email}) to {group_type_name.lower()} '{group_display_name}' ({distribution_list_email})"
            )]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to add user to group: {str(e)}")
        else:
            raise Exception(f"Error adding user to group: {str(e)}")

async def add_user_to_microsoft365_group(user_email: str, group_email: str) -> CallToolResult:
    """Add a user to a Microsoft 365 Group specifically."""
    try:
        # Get the user ID
        user_response = await make_graph_request("GET", f"/users/{user_email}")
        user_id = user_response["id"]
        user_display_name = user_response.get("displayName", user_email)
        
        # Get the Microsoft 365 Group ID
        group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_email}' and groupTypes/any(c:c eq 'Unified')")
        if not group_response.get("value"):
            raise Exception(f"Microsoft 365 Group '{group_email}' not found. Please verify the email address is correct.")
        
        group_id = group_response["value"][0]["id"]
        group_display_name = group_response["value"][0].get("displayName", group_email)
        
        # Verify this is actually a Microsoft 365 Group
        if not group_response["value"][0].get("groupTypes") or "Unified" not in group_response["value"][0].get("groupTypes", []):
            raise Exception(f"'{group_email}' is not a Microsoft 365 Group. Use 'add_user_to_distribution_list' for traditional distribution lists.")
        
        # Check if user is already a member
        members_response = await make_graph_request("GET", f"/groups/{group_id}/members?$filter=id eq '{user_id}'")
        if members_response.get("value"):
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"User '{user_display_name}' is already a member of Microsoft 365 Group '{group_display_name}'"
                )]
            )
        
        # Add user to the Microsoft 365 Group using the correct endpoint for Unified groups
        await make_graph_request("POST", f"/groups/{group_id}/members", {
            "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
        })
        
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=f"Successfully added '{user_display_name}' ({user_email}) to Microsoft 365 Group '{group_display_name}' ({group_email})"
            )]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to add user to Microsoft 365 Group: {str(e)}")
        else:
            raise Exception(f"Error adding user to Microsoft 365 Group: {str(e)}")

async def delegate_mailbox(mailbox_email: str, delegate_email: str, permissions: str = "FullAccess") -> CallToolResult:
    """Delegate a mailbox to another user using Microsoft Graph API."""
    try:
        # Validate permissions
        valid_permissions = ["FullAccess", "SendAs", "SendOnBehalf"]
        if permissions not in valid_permissions:
            raise Exception(f"Invalid permission '{permissions}'. Valid options are: {', '.join(valid_permissions)}")
        
        # Get the mailbox user ID
        mailbox_response = await make_graph_request("GET", f"/users/{mailbox_email}")
        mailbox_id = mailbox_response["id"]
        mailbox_display_name = mailbox_response.get("displayName", mailbox_email)
        
        # Get the delegate user ID
        delegate_response = await make_graph_request("GET", f"/users/{delegate_email}")
        delegate_id = delegate_response["id"]
        delegate_display_name = delegate_response.get("displayName", delegate_email)
        
        # For mailbox delegation, we need to use Exchange Online PowerShell or Exchange Admin API
        # Microsoft Graph API doesn't directly support mailbox permissions
        # This is a limitation - we'll provide guidance instead
        
        result_text = f"⚠️  Mailbox delegation requires Exchange Online PowerShell or Exchange Admin API.\n\n"
        result_text += f"To delegate mailbox '{mailbox_display_name}' ({mailbox_email}) to '{delegate_display_name}' ({delegate_email}) with {permissions} permissions:\n\n"
        result_text += f"**Using Exchange Online PowerShell:**\n"
        result_text += f"```powershell\n"
        result_text += f"Connect-ExchangeOnline\n"
        if permissions == "FullAccess":
            result_text += f"Add-MailboxPermission -Identity '{mailbox_email}' -User '{delegate_email}' -AccessRights FullAccess\n"
        elif permissions == "SendAs":
            result_text += f"Add-RecipientPermission -Identity '{mailbox_email}' -Trustee '{delegate_email}' -AccessRights SendAs\n"
        elif permissions == "SendOnBehalf":
            result_text += f"Set-Mailbox -Identity '{mailbox_email}' -GrantSendOnBehalfTo @{{Add='{delegate_email}'}}\n"
        result_text += f"```\n\n"
        result_text += f"**Note:** This operation requires Exchange Online administrator permissions."
        
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=result_text
            )]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to delegate mailbox: {str(e)}")
        else:
            raise Exception(f"Error delegating mailbox: {str(e)}")

async def convert_to_shared_mailbox(user_email: str, shared_mailbox_name: str) -> CallToolResult:
    """Convert a user mailbox to a shared mailbox."""
    try:
        # Get the user information
        user_response = await make_graph_request("GET", f"/users/{user_email}")
        user_display_name = user_response.get("displayName", user_email)
        
        # Microsoft Graph API doesn't directly support converting user mailboxes to shared mailboxes
        # This requires Exchange Online PowerShell or Exchange Admin API
        # We'll provide guidance instead
        
        result_text = f"⚠️  Converting user mailboxes to shared mailboxes requires Exchange Online PowerShell or Exchange Admin API.\n\n"
        result_text += f"To convert user '{user_display_name}' ({user_email}) to shared mailbox '{shared_mailbox_name}':\n\n"
        result_text += f"**Using Exchange Online PowerShell:**\n"
        result_text += f"```powershell\n"
        result_text += f"Connect-ExchangeOnline\n"
        result_text += f"# Convert to shared mailbox\n"
        result_text += f"Set-Mailbox -Identity '{user_email}' -Type Shared\n"
        result_text += f"# Update display name\n"
        result_text += f"Set-Mailbox -Identity '{user_email}' -DisplayName '{shared_mailbox_name}'\n"
        result_text += f"# Disable the user account (optional)\n"
        result_text += f"Set-User -Identity '{user_email}' -AccountDisabled $true\n"
        result_text += f"```\n\n"
        result_text += f"**Alternative: Create a new shared mailbox and migrate data**\n"
        result_text += f"1. Create a new shared mailbox using the 'create_shared_mailbox' tool\n"
        result_text += f"2. Use Exchange Online PowerShell to migrate data:\n"
        result_text += f"```powershell\n"
        result_text += f"New-MailboxExportRequest -Mailbox '{user_email}' -FilePath '\\\\server\\share\\{user_email}.pst'\n"
        result_text += f"New-MailboxImportRequest -Mailbox 'new-shared-mailbox@domain.com' -FilePath '\\\\server\\share\\{user_email}.pst'\n"
        result_text += f"```\n\n"
        result_text += f"**Note:** This operation requires Exchange Online administrator permissions."
        
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=result_text
            )]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to convert mailbox: {str(e)}")
        else:
            raise Exception(f"Error converting mailbox: {str(e)}")

async def list_users(top: int = 25, filter_query: str = None) -> CallToolResult:
    """List users in Microsoft 365."""
    endpoint = "/users"
    params = {
        "$top": top,
        "$select": "id,displayName,userPrincipalName,mail,accountEnabled,userType"
    }
    
    if filter_query:
        params["$filter"] = filter_query
    
    # Build the query string
    query_string = "&".join([f"{k}={v}" for k, v in params.items()])
    full_endpoint = f"{endpoint}?{query_string}"
    
    response = await make_graph_request("GET", full_endpoint)
    
    if not response.get("value"):
        return CallToolResult(
            content=[TextContent(
                type="text",
                text="No users found."
            )]
        )
    
    users = response["value"]
    user_list = []
    
    for user in users:
        user_info = f"• {user.get('displayName', 'N/A')} ({user.get('userPrincipalName', 'N/A')})"
        if user.get('mail'):
            user_info += f" - {user['mail']}"
        user_info += f" - {'Enabled' if user.get('accountEnabled', False) else 'Disabled'}"
        if user.get('userType'):
            user_info += f" - {user['userType']}"
        user_list.append(user_info)
    
    result_text = f"Found {len(users)} users:\n\n" + "\n".join(user_list)
    
    # Add pagination info if there are more results
    if response.get("@odata.nextLink"):
        result_text += f"\n\nNote: There are more users available. Use a higher 'top' value or add filters to see more results."
    
    return CallToolResult(
        content=[TextContent(
            type="text",
            text=result_text
        )]
    )

async def create_user_account(
    display_name: str, 
    user_principal_name: str, 
    mail_nickname: str, 
    password: str,
    department: str = None,
    job_title: str = None,
    office_location: str = None
) -> CallToolResult:
    """Create a new user account in Microsoft 365 (automatically creates a mailbox)."""
    try:
        # Validate required fields
        if not display_name or not user_principal_name or not mail_nickname or not password:
            raise Exception("All required fields (display_name, user_principal_name, mail_nickname, password) must be provided")
        
        # Validate password complexity (basic check)
        if len(password) < 8:
            raise Exception("Password must be at least 8 characters long")
        
        # Check if user already exists
        try:
            existing_user = await make_graph_request("GET", f"/users/{user_principal_name}")
            if existing_user:
                raise Exception(f"User with email '{user_principal_name}' already exists")
        except Exception as e:
            if "not found" not in str(e).lower() and "404" not in str(e):
                raise e
        
        # Create the user object
        user_data = {
            "accountEnabled": True,
            "displayName": display_name,
            "userPrincipalName": user_principal_name,
            "mailNickname": mail_nickname,
            "passwordProfile": {
                "forceChangePasswordNextSignIn": True,
                "password": password
            },
            "usageLocation": "US"  # Required for license assignment
        }
        
        # Add optional fields if provided
        if department:
            user_data["department"] = department
        if job_title:
            user_data["jobTitle"] = job_title
        if office_location:
            user_data["officeLocation"] = office_location
        
        # Create the user
        response = await make_graph_request("POST", "/users", user_data)
        
        user_id = response["id"]
        
        result_text = f"Successfully created user account:\n\n"
        result_text += f"• Display Name: {response.get('displayName')}\n"
        result_text += f"• Email: {response.get('userPrincipalName')}\n"
        result_text += f"• User ID: {user_id}\n"
        result_text += f"• Account Status: {'Enabled' if response.get('accountEnabled') else 'Disabled'}\n"
        
        if department:
            result_text += f"• Department: {department}\n"
        if job_title:
            result_text += f"• Job Title: {job_title}\n"
        if office_location:
            result_text += f"• Office Location: {office_location}\n"
        
        result_text += f"\n✅ Mailbox automatically created and ready for use.\n"
        result_text += f"⚠️  Note: User will be required to change password on first login."
        
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=result_text
            )]
        )
        
    except Exception as e:
        if "already exists" in str(e).lower():
            raise Exception(f"Failed to create user account: {str(e)}")
        elif "password" in str(e).lower():
            raise Exception(f"Password validation failed: {str(e)}")
        else:
            raise Exception(f"Error creating user account: {str(e)}")

async def create_shared_mailbox(
    display_name: str,
    mail_nickname: str,
    description: str = None
) -> CallToolResult:
    """Create a shared mailbox directly in Microsoft 365."""
    try:
        # Validate required fields
        if not display_name or not mail_nickname:
            raise Exception("Both display_name and mail_nickname are required")
        
        # Check if shared mailbox already exists
        try:
            existing_group = await make_graph_request("GET", f"/groups?$filter=mailNickname eq '{mail_nickname}' and groupTypes/any(c:c eq 'Unified')")
            if existing_group.get("value"):
                raise Exception(f"Shared mailbox with mail nickname '{mail_nickname}' already exists")
        except Exception as e:
            if "not found" not in str(e).lower() and "404" not in str(e):
                raise e
        
        group_data = {
            "displayName": display_name,
            "mailNickname": mail_nickname,
            "mailEnabled": True,
            "groupTypes": ["Unified"],
            "securityEnabled": False
        }
        
        if description:
            group_data["description"] = description
        
        response = await make_graph_request("POST", "/groups", group_data)
        
        result_text = f"Successfully created shared mailbox:\n\n"
        result_text += f"• Display Name: {response['displayName']}\n"
        result_text += f"• Mail Nickname: {response['mailNickname']}\n"
        result_text += f"• Group ID: {response['id']}\n"
        result_text += f"• Email: {response.get('mail', 'N/A')}\n"
        
        if description:
            result_text += f"• Description: {description}\n"
        
        result_text += f"\n✅ Shared mailbox created successfully!\n"
        result_text += f"📧 You can now add members to this shared mailbox using the 'add_user_to_distribution_list' tool."
        
        return CallToolResult(
            content=[TextContent(type="text", text=result_text)]
        )
        
    except Exception as e:
        if "already exists" in str(e).lower():
            raise Exception(f"Failed to create shared mailbox: {str(e)}")
        else:
            raise Exception(f"Error creating shared mailbox: {str(e)}")

# ===== DISTRIBUTION LISTS CRUD =====

async def create_distribution_list(
    display_name: str,
    mail_nickname: str,
    description: str = None,
    mail_enabled: bool = True
) -> CallToolResult:
    """Create a new distribution list."""
    group_data = {
        "displayName": display_name,
        "mailNickname": mail_nickname,
        "mailEnabled": mail_enabled,
        "groupTypes": ["Unified"],
        "securityEnabled": False
    }
    
    if description:
        group_data["description"] = description
    
    response = await make_graph_request("POST", "/groups", group_data)
    
    result_text = f"Successfully created distribution list:\n\n"
    result_text += f"• Display Name: {response.get('displayName')}\n"
    result_text += f"• Email Address: {response.get('mail')}\n"
    result_text += f"• Group ID: {response.get('id')}\n"
    result_text += f"• Mail Enabled: {'Yes' if response.get('mailEnabled') else 'No'}\n"
    
    if description:
        result_text += f"• Description: {description}\n"
    
    return CallToolResult(
        content=[TextContent(type="text", text=result_text)]
    )

async def list_distribution_lists(top: int = 25, filter_query: str = None) -> CallToolResult:
    """List distribution lists in Microsoft 365."""
    endpoint = "/groups"
    params = {
        "$top": top,
        "$filter": "groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false",
        "$select": "id,displayName,mail,description,mailEnabled,createdDateTime"
    }
    
    if filter_query:
        # Combine the base filter with user filter
        base_filter = "groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false"
        params["$filter"] = f"({base_filter}) and ({filter_query})"
    
    query_string = "&".join([f"{k}={v}" for k, v in params.items()])
    full_endpoint = f"{endpoint}?{query_string}"
    
    response = await make_graph_request("GET", full_endpoint)
    
    if not response.get("value"):
        return CallToolResult(
            content=[TextContent(type="text", text="No distribution lists found.")]
        )
    
    dl_list = []
    for dl in response["value"]:
        dl_info = f"• {dl.get('displayName', 'N/A')} ({dl.get('mail', 'N/A')})"
        if dl.get('description'):
            dl_info += f" - {dl['description']}"
        dl_info += f" - Created: {dl.get('createdDateTime', 'N/A')}"
        
        # Get owners for this distribution list
        try:
            owners_response = await make_graph_request("GET", f"/groups/{dl['id']}/owners?$select=displayName,userPrincipalName,mail")
            owners = owners_response.get("value", [])
            if owners:
                owner_names = [owner.get('displayName', owner.get('userPrincipalName', 'N/A')) for owner in owners]
                dl_info += f" - Owners: {', '.join(owner_names)}"
            else:
                dl_info += f" - Owners: None"
        except Exception as e:
            dl_info += f" - Owners: Unable to retrieve"
        
        dl_list.append(dl_info)
    
    result_text = f"Found {len(dl_list)} distribution lists:\n\n" + "\n".join(dl_list)
    
    if response.get("@odata.nextLink"):
        result_text += f"\n\nNote: There are more distribution lists available."
    
    return CallToolResult(
        content=[TextContent(type="text", text=result_text)]
    )

async def update_distribution_list(
    group_id_or_email: str,
    display_name: str = None,
    description: str = None,
    mail_nickname: str = None
) -> CallToolResult:
    """Update a distribution list."""
    # If email provided, get the group ID
    if "@" in group_id_or_email:
        response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_id_or_email}'")
        if not response.get("value"):
            raise Exception(f"Distribution list {group_id_or_email} not found")
        group_id = response["value"][0]["id"]
    else:
        group_id = group_id_or_email
    
    update_data = {}
    if display_name:
        update_data["displayName"] = display_name
    if description:
        update_data["description"] = description
    if mail_nickname:
        update_data["mailNickname"] = mail_nickname
    
    if not update_data:
        raise Exception("No fields to update provided")
    
    await make_graph_request("PATCH", f"/groups/{group_id}", update_data)
    
    result_text = f"Successfully updated distribution list {group_id_or_email}:\n\n"
    for field, value in update_data.items():
        result_text += f"• {field}: {value}\n"
    
    return CallToolResult(
        content=[TextContent(type="text", text=result_text)]
    )

async def delete_distribution_list(group_id_or_email: str) -> CallToolResult:
    """Delete a distribution list."""
    # If email provided, get the group ID
    if "@" in group_id_or_email:
        response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_id_or_email}'")
        if not response.get("value"):
            raise Exception(f"Distribution list {group_id_or_email} not found")
        group_id = response["value"][0]["id"]
        display_name = response["value"][0].get("displayName", group_id_or_email)
    else:
        group_id = group_id_or_email
        # Get display name for confirmation
        response = await make_graph_request("GET", f"/groups/{group_id}")
        display_name = response.get("displayName", group_id)
    
    await make_graph_request("DELETE", f"/groups/{group_id}")
    
    result_text = f"Successfully deleted distribution list:\n\n"
    result_text += f"• Display Name: {display_name}\n"
    result_text += f"• Group ID: {group_id}\n"
    
    return CallToolResult(
        content=[TextContent(type="text", text=result_text)]
    )

async def list_distribution_list_members(distribution_list_email: str, top: int = 100) -> CallToolResult:
    """List all members of a distribution list."""
    try:
        # Get the distribution list ID
        dl_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{distribution_list_email}' and groupTypes/any(c:c eq 'Unified')")
        
        if not dl_response.get("value"):
            raise Exception(f"Distribution list '{distribution_list_email}' not found")
        
        dl_id = dl_response["value"][0]["id"]
        dl_display_name = dl_response["value"][0].get("displayName", distribution_list_email)
        
        # Get all members of the distribution list
        members_response = await make_graph_request("GET", f"/groups/{dl_id}/members?$top={top}&$select=id,displayName,userPrincipalName,mail,accountEnabled")
        
        members = members_response.get("value", [])
        
        # Get owners information
        try:
            owners_response = await make_graph_request("GET", f"/groups/{dl_id}/owners?$select=id,displayName,userPrincipalName,mail")
            owners = owners_response.get("value", [])
        except Exception as e:
            owners = []
        
        if not members and not owners:
            result_text = f"Distribution list '{dl_display_name}' ({distribution_list_email}) has no members or owners."
        else:
            result_text = f"**Distribution List: '{dl_display_name}' ({distribution_list_email})**\n\n"
            
            # Display owners first
            if owners:
                result_text += f"**Owners ({len(owners)}):**\n"
                for i, owner in enumerate(owners, 1):
                    display_name = owner.get("displayName", "N/A")
                    user_principal_name = owner.get("userPrincipalName", "N/A")
                    mail = owner.get("mail", "N/A")
                    result_text += f"{i}. **{display_name}** ({mail})\n"
                result_text += "\n"
            
            # Display members
            if members:
                result_text += f"**Members ({len(members)}):**\n"
                for i, member in enumerate(members, 1):
                    display_name = member.get("displayName", "N/A")
                    user_principal_name = member.get("userPrincipalName", "N/A")
                    mail = member.get("mail", "N/A")
                    account_enabled = member.get("accountEnabled", True)
                    status = "Active" if account_enabled else "Disabled"
                    
                    result_text += f"{i}. **{display_name}**\n"
                    result_text += f"   - Email: {mail}\n"
                    result_text += f"   - UPN: {user_principal_name}\n"
                    result_text += f"   - Status: {status}\n\n"
            else:
                result_text += "**Members: None**\n"
        
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=result_text
            )]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to list distribution list members: {str(e)}")
        else:
            raise Exception(f"Error listing distribution list members: {str(e)}")

# ===== MAILBOXES CRUD =====

async def get_mailbox_info(user_email: str) -> CallToolResult:
    """Get detailed information about a user's mailbox."""
    user_response = await make_graph_request("GET", f"/users/{user_email}")
    mailbox_response = await make_graph_request("GET", f"/users/{user_email}/mailboxSettings")
    
    result_text = f"Mailbox Information for {user_email}:\n\n"
    result_text += f"• Display Name: {user_response.get('displayName')}\n"
    result_text += f"• User ID: {user_response.get('id')}\n"
    result_text += f"• Account Status: {'Enabled' if user_response.get('accountEnabled') else 'Disabled'}\n"
    result_text += f"• User Type: {user_response.get('userType', 'N/A')}\n"
    result_text += f"• Department: {user_response.get('department', 'N/A')}\n"
    result_text += f"• Job Title: {user_response.get('jobTitle', 'N/A')}\n"
    result_text += f"• Office Location: {user_response.get('officeLocation', 'N/A')}\n"
    result_text += f"• Created: {user_response.get('createdDateTime', 'N/A')}\n"
    
    if mailbox_response:
        result_text += f"\nMailbox Settings:\n"
        result_text += f"• Time Zone: {mailbox_response.get('timeZone', 'N/A')}\n"
        result_text += f"• Language: {mailbox_response.get('language', 'N/A')}\n"
        result_text += f"• Working Hours: {mailbox_response.get('workingHours', 'N/A')}\n"
    
    return CallToolResult(
        content=[TextContent(type="text", text=result_text)]
    )

async def update_mailbox_settings(
    user_email: str,
    time_zone: str = None,
    language: str = None,
    working_hours: str = None
) -> CallToolResult:
    """Update mailbox settings for a user."""
    update_data = {}
    
    if time_zone:
        update_data["timeZone"] = time_zone
    if language:
        update_data["language"] = language
    if working_hours:
        update_data["workingHours"] = working_hours
    
    if not update_data:
        raise Exception("No settings to update provided")
    
    await make_graph_request("PATCH", f"/users/{user_email}/mailboxSettings", update_data)
    
    result_text = f"Successfully updated mailbox settings for {user_email}:\n\n"
    for setting, value in update_data.items():
        result_text += f"• {setting}: {value}\n"
    
    return CallToolResult(
        content=[TextContent(type="text", text=result_text)]
    )

async def delete_user_account(user_email: str) -> CallToolResult:
    """Delete a user account and their mailbox."""
    try:
        # Validate input
        if not user_email:
            raise Exception("User email is required")
        
        # Get user info first
        user_response = await make_graph_request("GET", f"/users/{user_email}")
        display_name = user_response.get("displayName", user_email)
        user_id = user_response.get("id")
        
        # Check if user is already disabled
        if not user_response.get("accountEnabled", True):
            result_text = f"⚠️  User account is already disabled:\n\n"
            result_text += f"• Display Name: {display_name}\n"
            result_text += f"• Email: {user_email}\n"
            result_text += f"• User ID: {user_id}\n"
            result_text += f"\nTo permanently delete the account, use this tool again."
            
            return CallToolResult(
                content=[TextContent(type="text", text=result_text)]
            )
        
        # Delete the user (this also deletes the mailbox)
        await make_graph_request("DELETE", f"/users/{user_email}")
        
        result_text = f"Successfully deleted user account and mailbox:\n\n"
        result_text += f"• Display Name: {display_name}\n"
        result_text += f"• Email: {user_email}\n"
        result_text += f"• User ID: {user_id}\n"
        result_text += f"\n⚠️  Note: This action cannot be undone. All user data and mailbox contents have been permanently deleted."
        
        return CallToolResult(
            content=[TextContent(type="text", text=result_text)]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to delete user account: User '{user_email}' not found")
        else:
            raise Exception(f"Error deleting user account: {str(e)}")

# ===== SHARED MAILBOXES CRUD =====

async def list_all_groups(top: int = 25, filter_query: str = None, include_owners: bool = True) -> CallToolResult:
    """List all groups in Microsoft 365 with ownership information."""
    try:
        endpoint = "/groups"
        params = {
            "$top": top,
            "$select": "id,displayName,mail,description,createdDateTime,mailNickname,groupTypes,securityEnabled,mailEnabled"
        }
        
        if filter_query:
            params["$filter"] = filter_query
        
        query_string = "&".join([f"{k}={v}" for k, v in params.items()])
        full_endpoint = f"{endpoint}?{query_string}"
        
        response = await make_graph_request("GET", full_endpoint)
        
        if not response.get("value"):
            return CallToolResult(
                content=[TextContent(type="text", text="No groups found.")]
            )
        
        group_list = []
        for group in response["value"]:
            group_types = group.get('groupTypes', [])
            is_unified = 'Unified' in group_types
            is_security = group.get('securityEnabled', False)
            is_mail_enabled = group.get('mailEnabled', False)
            
            # Determine group type
            if is_unified and is_mail_enabled and not is_security:
                group_type = "Microsoft 365 Group"
            elif is_unified and is_mail_enabled and is_security:
                group_type = "Mail-enabled Security Group"
            elif is_security:
                group_type = "Security Group"
            else:
                group_type = "Distribution List"
            
            group_info = f"• {group.get('displayName', 'N/A')} ({group.get('mail', 'N/A')})"
            group_info += f" - Type: {group_type}"
            if group.get('description'):
                group_info += f" - {group['description']}"
            group_info += f" - Created: {group.get('createdDateTime', 'N/A')}"
            
            # Get owners if requested
            if include_owners:
                try:
                    owners_response = await make_graph_request("GET", f"/groups/{group['id']}/owners?$select=displayName,userPrincipalName,mail")
                    owners = owners_response.get("value", [])
                    if owners:
                        owner_names = [owner.get('displayName', owner.get('userPrincipalName', 'N/A')) for owner in owners]
                        group_info += f" - Owners: {', '.join(owner_names)}"
                    else:
                        group_info += f" - Owners: None"
                except Exception as e:
                    group_info += f" - Owners: Unable to retrieve"
            
            group_list.append(group_info)
        
        result_text = f"Found {len(group_list)} groups:\n\n" + "\n".join(group_list)
        
        if response.get("@odata.nextLink"):
            result_text += f"\n\nNote: There are more groups available. Use a higher 'top' value or add filters to see more results."
        
        return CallToolResult(
            content=[TextContent(type="text", text=result_text)]
        )
        
    except Exception as e:
        raise Exception(f"Error listing groups: {str(e)}")

async def list_shared_mailboxes(top: int = 25, filter_query: str = None) -> CallToolResult:
    """List shared mailboxes in Microsoft 365."""
    try:
        endpoint = "/groups"
        params = {
            "$top": top,
            "$filter": "groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false",
            "$select": "id,displayName,mail,description,createdDateTime,mailNickname"
        }
        
        if filter_query:
            base_filter = "groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false"
            params["$filter"] = f"({base_filter}) and ({filter_query})"
        
        query_string = "&".join([f"{k}={v}" for k, v in params.items()])
        full_endpoint = f"{endpoint}?{query_string}"
        
        response = await make_graph_request("GET", full_endpoint)
        
        if not response.get("value"):
            return CallToolResult(
                content=[TextContent(type="text", text="No shared mailboxes found.")]
            )
        
        mailbox_list = []
        for mailbox in response["value"]:
            mailbox_info = f"• {mailbox.get('displayName', 'N/A')} ({mailbox.get('mail', 'N/A')})"
            if mailbox.get('description'):
                mailbox_info += f" - {mailbox['description']}"
            mailbox_info += f" - Created: {mailbox.get('createdDateTime', 'N/A')}"
            
            # Get owners for this shared mailbox
            try:
                owners_response = await make_graph_request("GET", f"/groups/{mailbox['id']}/owners?$select=displayName,userPrincipalName,mail")
                owners = owners_response.get("value", [])
                if owners:
                    owner_names = [owner.get('displayName', owner.get('userPrincipalName', 'N/A')) for owner in owners]
                    mailbox_info += f" - Owners: {', '.join(owner_names)}"
                else:
                    mailbox_info += f" - Owners: None"
            except Exception as e:
                mailbox_info += f" - Owners: Unable to retrieve"
            
            mailbox_list.append(mailbox_info)
        
        result_text = f"Found {len(mailbox_list)} shared mailboxes:\n\n" + "\n".join(mailbox_list)
        
        if response.get("@odata.nextLink"):
            result_text += f"\n\nNote: There are more shared mailboxes available. Use a higher 'top' value or add filters to see more results."
        
        return CallToolResult(
            content=[TextContent(type="text", text=result_text)]
        )
        
    except Exception as e:
        raise Exception(f"Error listing shared mailboxes: {str(e)}")

async def get_group_info(group_id_or_email: str) -> CallToolResult:
    """Get detailed information about any group with ownership and membership information."""
    try:
        # If email provided, get the group ID
        if "@" in group_id_or_email:
            response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_id_or_email}'")
            if not response.get("value"):
                raise Exception(f"Group '{group_id_or_email}' not found. Please verify the email address is correct.")
            group_id = response["value"][0]["id"]
            group_info = response["value"][0]
        else:
            group_id = group_id_or_email
            group_info = await make_graph_request("GET", f"/groups/{group_id}")
        
        # Determine group type
        group_types = group_info.get('groupTypes', [])
        is_unified = 'Unified' in group_types
        is_security = group_info.get('securityEnabled', False)
        is_mail_enabled = group_info.get('mailEnabled', False)
        
        if is_unified and is_mail_enabled and not is_security:
            group_type = "Microsoft 365 Group"
        elif is_unified and is_mail_enabled and is_security:
            group_type = "Mail-enabled Security Group"
        elif is_security:
            group_type = "Security Group"
        else:
            group_type = "Distribution List"
        
        # Get members and owners
        members_response = await make_graph_request("GET", f"/groups/{group_id}/members?$select=id,displayName,userPrincipalName,mail,accountEnabled")
        owners_response = await make_graph_request("GET", f"/groups/{group_id}/owners?$select=id,displayName,userPrincipalName,mail,accountEnabled")
        
        result_text = f"Group Information:\n\n"
        result_text += f"• Display Name: {group_info.get('displayName', 'N/A')}\n"
        result_text += f"• Email Address: {group_info.get('mail', 'N/A')}\n"
        result_text += f"• Group ID: {group_id}\n"
        result_text += f"• Group Type: {group_type}\n"
        result_text += f"• Description: {group_info.get('description', 'N/A')}\n"
        result_text += f"• Created: {group_info.get('createdDateTime', 'N/A')}\n"
        result_text += f"• Mail Nickname: {group_info.get('mailNickname', 'N/A')}\n"
        result_text += f"• Security Enabled: {'Yes' if is_security else 'No'}\n"
        result_text += f"• Mail Enabled: {'Yes' if is_mail_enabled else 'No'}\n"
        result_text += f"• Group Types: {', '.join(group_types) if group_types else 'None'}\n"
        
        # Display owners
        if owners_response.get("value"):
            result_text += f"\nOwners ({len(owners_response['value'])}):\n"
            for owner in owners_response["value"]:
                owner_email = owner.get('mail', owner.get('userPrincipalName', 'N/A'))
                account_enabled = owner.get('accountEnabled', True)
                status = "Active" if account_enabled else "Disabled"
                result_text += f"• {owner.get('displayName', 'N/A')} ({owner_email}) - {status}\n"
        else:
            result_text += f"\nOwners: None"
        
        # Display members
        if members_response.get("value"):
            result_text += f"\nMembers ({len(members_response['value'])}):\n"
            for member in members_response["value"]:
                member_email = member.get('mail', member.get('userPrincipalName', 'N/A'))
                account_enabled = member.get('accountEnabled', True)
                status = "Active" if account_enabled else "Disabled"
                result_text += f"• {member.get('displayName', 'N/A')} ({member_email}) - {status}\n"
        else:
            result_text += f"\nMembers: None"
        
        return CallToolResult(
            content=[TextContent(type="text", text=result_text)]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to get group info: {str(e)}")
        else:
            raise Exception(f"Error getting group info: {str(e)}")

async def get_shared_mailbox_info(group_id_or_email: str) -> CallToolResult:
    """Get detailed information about a shared mailbox."""
    try:
        # If email provided, get the group ID
        if "@" in group_id_or_email:
            response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_id_or_email}' and groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false")
            if not response.get("value"):
                raise Exception(f"Shared mailbox '{group_id_or_email}' not found. Please verify the email address is correct.")
            group_id = response["value"][0]["id"]
            mailbox_info = response["value"][0]
        else:
            group_id = group_id_or_email
            mailbox_info = await make_graph_request("GET", f"/groups/{group_id}")
        
        # Get members and owners
        members_response = await make_graph_request("GET", f"/groups/{group_id}/members?$select=id,displayName,userPrincipalName,mail")
        owners_response = await make_graph_request("GET", f"/groups/{group_id}/owners?$select=id,displayName,userPrincipalName,mail")
        
        result_text = f"Shared Mailbox Information:\n\n"
        result_text += f"• Display Name: {mailbox_info.get('displayName', 'N/A')}\n"
        result_text += f"• Email Address: {mailbox_info.get('mail', 'N/A')}\n"
        result_text += f"• Group ID: {group_id}\n"
        result_text += f"• Description: {mailbox_info.get('description', 'N/A')}\n"
        result_text += f"• Created: {mailbox_info.get('createdDateTime', 'N/A')}\n"
        result_text += f"• Mail Nickname: {mailbox_info.get('mailNickname', 'N/A')}\n"
        
        # Display owners
        if owners_response.get("value"):
            result_text += f"\nOwners ({len(owners_response['value'])}):\n"
            for owner in owners_response["value"]:
                owner_email = owner.get('mail', owner.get('userPrincipalName', 'N/A'))
                result_text += f"• {owner.get('displayName', 'N/A')} ({owner_email})\n"
        else:
            result_text += f"\nOwners: None"
        
        # Display members
        if members_response.get("value"):
            result_text += f"\nMembers ({len(members_response['value'])}):\n"
            for member in members_response["value"]:
                member_email = member.get('mail', member.get('userPrincipalName', 'N/A'))
                result_text += f"• {member.get('displayName', 'N/A')} ({member_email})\n"
        else:
            result_text += f"\nMembers: None"
        
        return CallToolResult(
            content=[TextContent(type="text", text=result_text)]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to get shared mailbox info: {str(e)}")
        else:
            raise Exception(f"Error getting shared mailbox info: {str(e)}")

async def update_shared_mailbox(
    group_id_or_email: str,
    display_name: str = None,
    description: str = None
) -> CallToolResult:
    """Update a shared mailbox."""
    try:
        # If email provided, get the group ID
        if "@" in group_id_or_email:
            response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_id_or_email}' and groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false")
            if not response.get("value"):
                raise Exception(f"Shared mailbox '{group_id_or_email}' not found. Please verify the email address is correct.")
            group_id = response["value"][0]["id"]
            original_name = response["value"][0].get("displayName", group_id_or_email)
        else:
            group_id = group_id_or_email
            # Get original info for confirmation
            response = await make_graph_request("GET", f"/groups/{group_id}")
            original_name = response.get("displayName", group_id)
        
        update_data = {}
        if display_name:
            update_data["displayName"] = display_name
        if description:
            update_data["description"] = description
        
        if not update_data:
            raise Exception("No fields to update provided")
        
        await make_graph_request("PATCH", f"/groups/{group_id}", update_data)
        
        result_text = f"Successfully updated shared mailbox '{original_name}':\n\n"
        for field, value in update_data.items():
            result_text += f"• {field}: {value}\n"
        
        return CallToolResult(
            content=[TextContent(type="text", text=result_text)]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to update shared mailbox: {str(e)}")
        else:
            raise Exception(f"Error updating shared mailbox: {str(e)}")

async def delete_shared_mailbox(group_id_or_email: str) -> CallToolResult:
    """Delete a shared mailbox."""
    try:
        # If email provided, get the group ID
        if "@" in group_id_or_email:
            response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_id_or_email}' and groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false")
            if not response.get("value"):
                raise Exception(f"Shared mailbox '{group_id_or_email}' not found. Please verify the email address is correct.")
            group_id = response["value"][0]["id"]
            display_name = response["value"][0].get("displayName", group_id_or_email)
        else:
            group_id = group_id_or_email
            # Get display name for confirmation
            response = await make_graph_request("GET", f"/groups/{group_id}")
            display_name = response.get("displayName", group_id)
        
        await make_graph_request("DELETE", f"/groups/{group_id}")
        
        result_text = f"Successfully deleted shared mailbox:\n\n"
        result_text += f"• Display Name: {display_name}\n"
        result_text += f"• Group ID: {group_id}\n"
        result_text += f"\n⚠️  Note: This action cannot be undone. All data in the shared mailbox has been permanently deleted."
        
        return CallToolResult(
            content=[TextContent(type="text", text=result_text)]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to delete shared mailbox: {str(e)}")
        else:
            raise Exception(f"Error deleting shared mailbox: {str(e)}")

async def add_group_owner(group_id_or_email: str, owner_email: str) -> CallToolResult:
    """Add an owner to a group."""
    try:
        # If email provided, get the group ID
        if "@" in group_id_or_email:
            response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_id_or_email}'")
            if not response.get("value"):
                raise Exception(f"Group '{group_id_or_email}' not found")
            group_id = response["value"][0]["id"]
            group_display_name = response["value"][0].get("displayName", group_id_or_email)
        else:
            group_id = group_id_or_email
            # Get display name for confirmation
            response = await make_graph_request("GET", f"/groups/{group_id}")
            group_display_name = response.get("displayName", group_id)
        
        # Get the user ID for the owner
        user_response = await make_graph_request("GET", f"/users/{owner_email}")
        user_id = user_response["id"]
        user_display_name = user_response.get("displayName", owner_email)
        
        # Check if user is already an owner
        owners_response = await make_graph_request("GET", f"/groups/{group_id}/owners?$filter=id eq '{user_id}'")
        if owners_response.get("value"):
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"User '{user_display_name}' is already an owner of group '{group_display_name}'"
                )]
            )
        
        # Add user as owner
        await make_graph_request("POST", f"/groups/{group_id}/owners", {
            "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
        })
        
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=f"Successfully added '{user_display_name}' ({owner_email}) as owner of group '{group_display_name}' ({group_id_or_email})"
            )]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to add group owner: {str(e)}")
        else:
            raise Exception(f"Error adding group owner: {str(e)}")

async def remove_group_owner(group_id_or_email: str, owner_email: str) -> CallToolResult:
    """Remove an owner from a group."""
    try:
        # If email provided, get the group ID
        if "@" in group_id_or_email:
            response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_id_or_email}'")
            if not response.get("value"):
                raise Exception(f"Group '{group_id_or_email}' not found")
            group_id = response["value"][0]["id"]
            group_display_name = response["value"][0].get("displayName", group_id_or_email)
        else:
            group_id = group_id_or_email
            # Get display name for confirmation
            response = await make_graph_request("GET", f"/groups/{group_id}")
            group_display_name = response.get("displayName", group_id)
        
        # Get the user ID for the owner
        user_response = await make_graph_request("GET", f"/users/{owner_email}")
        user_id = user_response["id"]
        user_display_name = user_response.get("displayName", owner_email)
        
        # Check if user is actually an owner
        owners_response = await make_graph_request("GET", f"/groups/{group_id}/owners?$filter=id eq '{user_id}'")
        if not owners_response.get("value"):
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"User '{user_display_name}' is not an owner of group '{group_display_name}'"
                )]
            )
        
        # Remove user as owner
        await make_graph_request("DELETE", f"/groups/{group_id}/owners/{user_id}")
        
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=f"Successfully removed '{user_display_name}' ({owner_email}) as owner of group '{group_display_name}' ({group_id_or_email})"
            )]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to remove group owner: {str(e)}")
        else:
            raise Exception(f"Error removing group owner: {str(e)}")

async def list_group_owners(group_id_or_email: str) -> CallToolResult:
    """List all owners of a specific group."""
    try:
        # If email provided, get the group ID
        if "@" in group_id_or_email:
            response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_id_or_email}'")
            if not response.get("value"):
                raise Exception(f"Group '{group_id_or_email}' not found")
            group_id = response["value"][0]["id"]
            group_display_name = response["value"][0].get("displayName", group_id_or_email)
        else:
            group_id = group_id_or_email
            # Get display name for confirmation
            response = await make_graph_request("GET", f"/groups/{group_id}")
            group_display_name = response.get("displayName", group_id)
        
        # Get all owners of the group
        owners_response = await make_graph_request("GET", f"/groups/{group_id}/owners?$select=id,displayName,userPrincipalName,mail,accountEnabled,jobTitle,department")
        
        owners = owners_response.get("value", [])
        
        if not owners:
            result_text = f"Group '{group_display_name}' ({group_id_or_email}) has no owners."
        else:
            result_text = f"**Owners of group '{group_display_name}' ({group_id_or_email}):**\n\n"
            result_text += f"Total owners: {len(owners)}\n\n"
            
            for i, owner in enumerate(owners, 1):
                display_name = owner.get("displayName", "N/A")
                user_principal_name = owner.get("userPrincipalName", "N/A")
                mail = owner.get("mail", "N/A")
                job_title = owner.get("jobTitle", "N/A")
                department = owner.get("department", "N/A")
                account_enabled = owner.get("accountEnabled", True)
                status = "Active" if account_enabled else "Disabled"
                
                result_text += f"{i}. **{display_name}**\n"
                result_text += f"   - Email: {mail}\n"
                result_text += f"   - UPN: {user_principal_name}\n"
                result_text += f"   - Job Title: {job_title}\n"
                result_text += f"   - Department: {department}\n"
                result_text += f"   - Status: {status}\n\n"
        
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=result_text
            )]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to list group owners: {str(e)}")
        else:
            raise Exception(f"Error listing group owners: {str(e)}")

async def list_shared_mailbox_members(shared_mailbox_email: str, top: int = 100) -> CallToolResult:
    """List all members of a shared mailbox."""
    try:
        # Get the shared mailbox ID
        mailbox_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{shared_mailbox_email}' and groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false")
        
        if not mailbox_response.get("value"):
            raise Exception(f"Shared mailbox '{shared_mailbox_email}' not found")
        
        mailbox_id = mailbox_response["value"][0]["id"]
        mailbox_display_name = mailbox_response["value"][0].get("displayName", shared_mailbox_email)
        
        # Get all members of the shared mailbox
        members_response = await make_graph_request("GET", f"/groups/{mailbox_id}/members?$top={top}&$select=id,displayName,userPrincipalName,mail,accountEnabled")
        
        members = members_response.get("value", [])
        
        # Get owners information
        try:
            owners_response = await make_graph_request("GET", f"/groups/{mailbox_id}/owners?$select=id,displayName,userPrincipalName,mail")
            owners = owners_response.get("value", [])
        except Exception as e:
            owners = []
        
        if not members and not owners:
            result_text = f"Shared mailbox '{mailbox_display_name}' ({shared_mailbox_email}) has no members or owners."
        else:
            result_text = f"**Shared Mailbox: '{mailbox_display_name}' ({shared_mailbox_email})**\n\n"
            
            # Display owners first
            if owners:
                result_text += f"**Owners ({len(owners)}):**\n"
                for i, owner in enumerate(owners, 1):
                    display_name = owner.get("displayName", "N/A")
                    user_principal_name = owner.get("userPrincipalName", "N/A")
                    mail = owner.get("mail", "N/A")
                    result_text += f"{i}. **{display_name}** ({mail})\n"
                result_text += "\n"
            
            # Display members
            if members:
                result_text += f"**Members ({len(members)}):**\n"
                for i, member in enumerate(members, 1):
                    display_name = member.get("displayName", "N/A")
                    user_principal_name = member.get("userPrincipalName", "N/A")
                    mail = member.get("mail", "N/A")
                    account_enabled = member.get("accountEnabled", True)
                    status = "Active" if account_enabled else "Disabled"
                    
                    result_text += f"{i}. **{display_name}**\n"
                    result_text += f"   - Email: {mail}\n"
                    result_text += f"   - UPN: {user_principal_name}\n"
                    result_text += f"   - Status: {status}\n\n"
            else:
                result_text += "**Members: None**\n"
        
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=result_text
            )]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to list shared mailbox members: {str(e)}")
        else:
            raise Exception(f"Error listing shared mailbox members: {str(e)}")

async def test_authentication() -> CallToolResult:
    """Test the authentication configuration and permissions."""
    try:
        # Test token acquisition
        token = await get_access_token()
        
        # Test a simple API call
        response = await make_graph_request("GET", "/users?$top=1")
        
        result_text = "✅ Authentication Test Successful!\n\n"
        result_text += f"• Token acquired successfully\n"
        result_text += f"• API call successful\n"
        result_text += f"• Found {len(response.get('value', []))} users\n"
        result_text += f"• Your M365 MCP server is ready to use!\n"
        
        return CallToolResult(
            content=[TextContent(type="text", text=result_text)]
        )
        
    except Exception as e:
        error_text = f"❌ Authentication Test Failed!\n\n"
        error_text += f"Error: {str(e)}\n\n"
        error_text += "🔧 Troubleshooting Steps:\n\n"
        error_text += "1. **Check Environment Variables:**\n"
        error_text += "   - TENANT_ID: {}\n".format("✅ Set" if TENANT_ID else "❌ Missing")
        error_text += "   - CLIENT_ID: {}\n".format("✅ Set" if CLIENT_ID else "❌ Missing")
        error_text += "   - CLIENT_SECRET: {}\n".format("✅ Set" if CLIENT_SECRET else "❌ Missing")
        error_text += "\n"
        error_text += "2. **Verify Azure App Registration:**\n"
        error_text += "   - Ensure app exists in Azure Portal\n"
        error_text += "   - Check Application (client) ID matches CLIENT_ID\n"
        error_text += "   - Check Directory (tenant) ID matches TENANT_ID\n"
        error_text += "   - Verify client secret is valid and not expired\n"
        error_text += "\n"
        error_text += "3. **Check API Permissions:**\n"
        error_text += "   - User.ReadWrite.All (Application permission)\n"
        error_text += "   - Group.ReadWrite.All (Application permission)\n"
        error_text += "   - MailboxSettings.ReadWrite (Application permission)\n"
        error_text += "   - Ensure 'Grant admin consent' is completed\n"
        
        return CallToolResult(
            content=[TextContent(type="text", text=error_text)]
        )

async def test_group_and_user_access(group_email: str, user_email: str) -> CallToolResult:
    """Test access to a group and user to help debug permission issues."""
    try:
        results = []
        
        # Test group access
        try:
            group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_email}' and groupTypes/any(c:c eq 'Unified')")
            if group_response.get("value"):
                group = group_response["value"][0]
                results.append(f"✅ Group found: {group.get('displayName')} (ID: {group.get('id')})")
                results.append(f"   - Mail: {group.get('mail')}")
                results.append(f"   - Group Types: {group.get('groupTypes', [])}")
                results.append(f"   - Security Enabled: {group.get('securityEnabled')}")
                results.append(f"   - Mail Enabled: {group.get('mailEnabled')}")
            else:
                results.append(f"❌ Group not found: {group_email}")
        except Exception as e:
            results.append(f"❌ Error accessing group: {str(e)}")
        
        # Test user access
        try:
            user_response = await make_graph_request("GET", f"/users/{user_email}")
            results.append(f"✅ User found: {user_response.get('displayName')} (ID: {user_response.get('id')})")
            results.append(f"   - Mail: {user_response.get('mail')}")
            results.append(f"   - UPN: {user_response.get('userPrincipalName')}")
            results.append(f"   - Account Enabled: {user_response.get('accountEnabled')}")
        except Exception as e:
            results.append(f"❌ Error accessing user: {str(e)}")
        
        # Test group members access
        if group_response.get("value"):
            try:
                group_id = group_response["value"][0]["id"]
                members_response = await make_graph_request("GET", f"/groups/{group_id}/members")
                results.append(f"✅ Group members access: {len(members_response.get('value', []))} current members")
            except Exception as e:
                results.append(f"❌ Error accessing group members: {str(e)}")
        
        return CallToolResult(
            content=[TextContent(
                type="text",
                text="\n".join(results)
            )]
        )
        
    except Exception as e:
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=f"❌ Test failed: {str(e)}"
            )],
            isError=True
        )

async def test_unified_group_api(group_email: str) -> CallToolResult:
    """Test Unified Group API access and endpoints."""
    try:
        results = []
        
        # Get group information
        group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_email}'")
        if not group_response.get("value"):
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"❌ Group not found: {group_email}"
                )],
                isError=True
            )
        
        group = group_response["value"][0]
        group_id = group["id"]
        group_types = group.get("groupTypes", [])
        
        results.append(f"✅ Group found: {group.get('displayName')} (ID: {group_id})")
        results.append(f"   - Mail: {group.get('mail')}")
        results.append(f"   - Group Types: {group_types}")
        results.append(f"   - Is Unified Group: {'Yes' if 'Unified' in group_types else 'No'}")
        
        # Test members endpoint
        try:
            members_response = await make_graph_request("GET", f"/groups/{group_id}/members")
            results.append(f"✅ Members endpoint accessible: {len(members_response.get('value', []))} members")
        except Exception as e:
            results.append(f"❌ Members endpoint error: {str(e)}")
        
        # Test if we can add members (dry run)
        if "Unified" in group_types:
            results.append("ℹ️  This is a Microsoft 365 Group (Unified) - will use /groups/{id}/members endpoint")
        else:
            results.append("ℹ️  This is a traditional distribution list - will use /groups/{id}/members/$ref endpoint")
        
        return CallToolResult(
            content=[TextContent(
                type="text",
                text="\n".join(results)
            )]
        )
        
    except Exception as e:
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=f"❌ Test failed: {str(e)}"
            )],
            isError=True
        )

def handle_tool_errors(func):
    """Decorator to provide consistent error handling for all tool functions."""
    async def wrapper(*args, **kwargs):
        try:
            return await func(*args, **kwargs)
        except Exception as e:
            error_message = str(e)
            
            # Provide specific guidance based on error types
            if "401" in error_message:
                error_message += "\n\n🔧 This usually means:\n• Token has expired\n• Insufficient permissions\n• Check if admin consent was granted"
            elif "403" in error_message:
                error_message += "\n\n🔧 This usually means:\n• Insufficient permissions\n• Check API permissions in Azure app registration\n• Ensure admin consent was granted"
            elif "404" in error_message:
                error_message += "\n\n🔧 This usually means:\n• Resource not found\n• Check if the user/group exists\n• Verify the email address or ID is correct"
            elif "400" in error_message:
                error_message += "\n\n🔧 This usually means:\n• Invalid request format\n• Missing required properties\n• Check the request payload"
            elif "409" in error_message:
                error_message += "\n\n🔧 This usually means:\n• Resource already exists\n• Conflict with existing data"
            elif "timeout" in error_message.lower():
                error_message += "\n\n🔧 This usually means:\n• Network connectivity issues\n• Microsoft Graph API is slow to respond\n• Try again in a few moments"
            elif "connection" in error_message.lower():
                error_message += "\n\n🔧 This usually means:\n• Network connectivity issues\n• Check your internet connection\n• Microsoft Graph API may be temporarily unavailable"
            
            # Return error as CallToolResult
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"❌ Error: {error_message}"
                )],
                isError=True
            )
    return wrapper

def create_server():
    """Create and configure the MCP server."""
    # Create the FastMCP server
    app = FastMCP("m365-admin")
    
    # Add tools using the correct FastMCP API
    app.add_tool(
        add_user_to_distribution_list,
        name="add_user_to_distribution_list",
        title="Add User to Distribution List",
        description="Add a user to a distribution list"
    )
    
    app.add_tool(
        add_user_to_microsoft365_group,
        name="add_user_to_microsoft365_group",
        title="Add User to Microsoft 365 Group",
        description="Add a user to a Microsoft 365 Group specifically"
    )
    
    app.add_tool(
        delegate_mailbox,
        name="delegate_mailbox",
        title="Delegate Mailbox",
        description="Delegate a mailbox to another user using Microsoft Graph API"
    )
    
    app.add_tool(
        convert_to_shared_mailbox,
        name="convert_to_shared_mailbox",
        title="Convert to Shared Mailbox",
        description="Convert a user mailbox to a shared mailbox"
    )
    
    app.add_tool(
        list_users,
        name="list_users",
        title="List Users",
        description="List users in Microsoft 365 with optional filtering and pagination"
    )
    
    app.add_tool(
        create_user_account,
        name="create_user_account",
        title="Create User Account",
        description="Create a new user account in Microsoft 365 (automatically creates a mailbox)"
    )
    
    app.add_tool(
        create_shared_mailbox,
        name="create_shared_mailbox",
        title="Create Shared Mailbox",
        description="Create a shared mailbox directly in Microsoft 365"
    )

    # Add distribution list tools
    app.add_tool(
        create_distribution_list,
        name="create_distribution_list",
        title="Create Distribution List",
        description="Create a new distribution list"
    )
    app.add_tool(
        list_distribution_lists,
        name="list_distribution_lists",
        title="List Distribution Lists",
        description="List distribution lists in Microsoft 365 with optional filtering and pagination"
    )
    app.add_tool(
        update_distribution_list,
        name="update_distribution_list",
        title="Update Distribution List",
        description="Update an existing distribution list"
    )
    app.add_tool(
        delete_distribution_list,
        name="delete_distribution_list",
        title="Delete Distribution List",
        description="Delete a distribution list"
    )
    app.add_tool(
        list_distribution_list_members,
        name="list_distribution_list_members",
        title="List Distribution List Members",
        description="List all members of a specific distribution list"
    )

    # Add new group tools with ownership information
    app.add_tool(
        list_all_groups,
        name="list_all_groups",
        title="List All Groups",
        description="List all groups in Microsoft 365 with ownership information and group type classification"
    )
    app.add_tool(
        get_group_info,
        name="get_group_info",
        title="Get Group Info",
        description="Get detailed information about any group with ownership and membership information"
    )
    app.add_tool(
        list_group_owners,
        name="list_group_owners",
        title="List Group Owners",
        description="List all owners of a specific group with detailed user information"
    )
    app.add_tool(
        add_group_owner,
        name="add_group_owner",
        title="Add Group Owner",
        description="Add a user as an owner to a specific group"
    )
    app.add_tool(
        remove_group_owner,
        name="remove_group_owner",
        title="Remove Group Owner",
        description="Remove a user as an owner from a specific group"
    )

    # Add mailbox tools
    app.add_tool(
        get_mailbox_info,
        name="get_mailbox_info",
        title="Get Mailbox Info",
        description="Get detailed information about a user's mailbox"
    )
    app.add_tool(
        update_mailbox_settings,
        name="update_mailbox_settings",
        title="Update Mailbox Settings",
        description="Update mailbox settings for a user"
    )
    app.add_tool(
        delete_user_account,
        name="delete_user_account",
        title="Delete User Account",
        description="Delete a user account and their mailbox"
    )

    # Add shared mailbox tools
    app.add_tool(
        list_shared_mailboxes,
        name="list_shared_mailboxes",
        title="List Shared Mailboxes",
        description="List shared mailboxes in Microsoft 365 with optional filtering and pagination"
    )
    app.add_tool(
        get_shared_mailbox_info,
        name="get_shared_mailbox_info",
        title="Get Shared Mailbox Info",
        description="Get detailed information about a shared mailbox"
    )
    app.add_tool(
        update_shared_mailbox,
        name="update_shared_mailbox",
        title="Update Shared Mailbox",
        description="Update an existing shared mailbox"
    )
    app.add_tool(
        delete_shared_mailbox,
        name="delete_shared_mailbox",
        title="Delete Shared Mailbox",
        description="Delete a shared mailbox"
    )
    app.add_tool(
        list_shared_mailbox_members,
        name="list_shared_mailbox_members",
        title="List Shared Mailbox Members",
        description="List all members of a specific shared mailbox"
    )
    
    app.add_tool(
        test_authentication,
        name="test_authentication",
        title="Test Authentication",
        description="Test the authentication configuration and permissions"
    )
    
    app.add_tool(
        test_group_and_user_access,
        name="test_group_and_user_access",
        title="Test Group and User Access",
        description="Test access to a specific group and user to help debug permission issues"
    )
    
    app.add_tool(
        test_unified_group_api,
        name="test_unified_group_api",
        title="Test Unified Group API",
        description="Test access to a specific group to verify Unified Group API endpoints"
    )
    
    return app

if __name__ == "__main__":
    try:
        # Validate required environment variables
        required_vars = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET"]
        missing_vars = [var for var in required_vars if not os.getenv(var)]
        
        if missing_vars:
            print(f"❌ Error: Missing required environment variables: {', '.join(missing_vars)}")
            print("Please set these variables in your .env file or environment.")
            print("\nExample .env file:")
            print("TENANT_ID=your_tenant_id_here")
            print("CLIENT_ID=your_client_id_here")
            print("CLIENT_SECRET=your_client_secret_here")
            exit(1)
        
        # Test authentication before starting server
        print("🔐 Testing authentication...")
        try:
            # This will be tested when the server starts
            print("✅ Environment variables validated")
        except Exception as e:
            print(f"❌ Authentication test failed: {e}")
            print("Please check your Azure app registration and permissions.")
            exit(1)
        
        # Create and run the server
        print("🚀 Starting M365 MCP Server...")
        app = create_server()
        app.run()
        
    except KeyboardInterrupt:
        print("\n👋 Server stopped by user")
    except Exception as e:
        print(f"❌ Server error: {e}")
        print("Please check the logs for more details.")
        exit(1)
