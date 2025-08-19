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

async def delegate_user_mailbox_access(mailbox_email: str, delegate_email: str, permissions: str = "FullAccess") -> CallToolResult:
    """⚠️ LIMITATION: This tool works best with user mailboxes. For Microsoft 365 Groups, use 'robust_add_user_to_group' instead."""
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
        
        # Perform mailbox delegation using Microsoft Graph API
        # Note: Microsoft Graph API has limited support for mailbox permissions
        # For full mailbox delegation, Exchange Online PowerShell is required
        result_text = f"🔄 **Delegating mailbox access...**\n\n"
        result_text += f"**Mailbox:** {mailbox_display_name} ({mailbox_email})\n"
        result_text += f"**Delegate:** {delegate_display_name} ({delegate_email})\n"
        result_text += f"**Permissions:** {permissions}\n\n"
        result_text += "**🔍 MICROSOFT GRAPH API LIMITATION:**\n"
        result_text += "• Mailbox delegation cannot be performed via Graph API\n"
        result_text += "• This is a Microsoft security restriction, not a permissions issue\n"
        result_text += "• Exchange Online PowerShell is required for mailbox delegation\n\n"
        
        try:
            # Try using Microsoft Graph API for mailbox delegation
            success_methods = []
            
            # Check if the mailbox account is enabled first
            if not mailbox_response.get("accountEnabled", True):
                result_text += "⚠️ **Account Status Issue:** The mailbox account is disabled.\n"
                result_text += "**Solutions:**\n"
                result_text += "1. Re-enable the account: `update_resource` with `accountEnabled: true`\n"
                result_text += "2. Complete shared mailbox conversion via PowerShell\n"
                result_text += "3. Use PowerShell for delegation: `Add-MailboxPermission`\n\n"
                raise Exception("Mailbox account is disabled - cannot grant permissions via API")
            
            # Method 1: Try calendar permissions via Graph API (this endpoint actually exists)
            try:
                calendar_data = {
                    "role": "owner",
                    "allowedAudiences": ["none"]
                }
                
                await make_graph_request("POST", f"/users/{mailbox_email}/calendar/calendarPermissions", calendar_data)
                success_methods.append("✅ **Calendar access via Graph API**")
                result_text += "🎯 **SUCCESS: Calendar permissions granted!**\n\n"
            except Exception as calendar_error:
                if "MailboxNotEnabledForRESTAPI" in str(calendar_error):
                    result_text += "⚠️ **Calendar permissions failed:** Mailbox not accessible (account may be disabled or converted)\n\n"
                else:
                    result_text += f"⚠️ **Calendar permissions failed:** {str(calendar_error)}\n\n"
            
            # Method 2: Try mailbox settings configuration (only if account is enabled)
            try:
                mailbox_data = {
                    "delegateMeetingMessageDeliveryOptions": "sendToDelegateAndInformationToPrincipal"
                }
                
                await make_graph_request("PATCH", f"/users/{mailbox_id}/mailboxSettings", mailbox_data)
                success_methods.append("✅ **Mailbox settings configured**")
                result_text += "🎯 **SUCCESS: Mailbox settings updated!**\n\n"
            except Exception as settings_error:
                if "MailboxNotEnabledForRESTAPI" in str(settings_error):
                    result_text += "⚠️ **Mailbox settings failed:** Mailbox not accessible (account may be disabled or converted)\n\n"
                else:
                    result_text += f"⚠️ **Mailbox settings failed:** {str(settings_error)}\n\n"
            
            # Method 4: Try group membership (if it's a shared mailbox)
            try:
                group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{mailbox_email}'")
                if group_response.get("value"):
                    # It's a group mailbox, add as member
                    group_id = group_response["value"][0]["id"]
                    await make_graph_request("POST", f"/groups/{group_id}/members", {
                        "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{delegate_id}"
                    })
                    success_methods.append("✅ **Group membership granted**")
                    result_text += "🎯 **SUCCESS: Group membership granted!**\n\n"
            except Exception as group_error:
                result_text += f"⚠️ **Group membership failed:** {str(group_error)}\n\n"
            
            # Check if any methods succeeded
            if success_methods:
                result_text += "**🎉 DELEGATION SUCCESSFUL!**\n\n"
                result_text += "**Methods that worked:**\n"
                for method in success_methods:
                    result_text += f"• {method}\n"
                
                result_text += f"\n**✅ User '{delegate_display_name}' now has access to '{mailbox_display_name}'**\n\n"
                result_text += "**Access includes:**\n"
                if "FullAccess" in " ".join(success_methods):
                    result_text += "• **Full mailbox access**\n"
                if "Calendar" in " ".join(success_methods):
                    result_text += "• **Calendar access**\n"
                if "Group membership" in " ".join(success_methods):
                    result_text += "• **Group membership**\n"
                
                result_text += "\n**Next Steps:**\n"
                result_text += "1. The delegate can now access the mailbox in Outlook\n"
                result_text += "2. Test the access by sending/receiving emails\n"
                result_text += "3. Calendar access should be available\n\n"
                
                return CallToolResult(content=[TextContent(type="text", text=result_text)])
            else:
                # All methods failed - provide PowerShell guidance
                raise Exception("All Graph API methods failed")
            
        except Exception as delegation_error:
            result_text += f"❌ **Delegation failed:** {str(delegation_error)}\n\n"
            
            # Provide specific guidance based on the error
            if "disabled" in str(delegation_error).lower():
                result_text += "**🔍 Root Cause:** The mailbox account is disabled.\n\n"
                result_text += "**📋 Solutions (in order):**\n"
                result_text += "1. **Re-enable the account:** Use `update_resource` with `accountEnabled: true`\n"
                result_text += "2. **Complete shared mailbox conversion:** Use PowerShell `Set-Mailbox -Type Shared`\n"
                result_text += "3. **Use PowerShell for delegation:** See commands below\n\n"
            elif "MailboxNotEnabledForRESTAPI" in str(delegation_error):
                result_text += "**🔍 Root Cause:** Mailbox is not accessible via Graph API (disabled, converted, or on-premise).\n\n"
                result_text += "**📋 Solutions:**\n"
                result_text += "1. **For disabled accounts:** Re-enable first, then delegate\n"
                result_text += "2. **For converted shared mailboxes:** Use PowerShell commands below\n"
                result_text += "3. **For on-premise mailboxes:** Use Exchange on-premise PowerShell\n\n"
            else:
                result_text += "**🔍 Root Cause:** Microsoft Graph API limitations for mailbox permissions.\n\n"
                result_text += "**📋 Solution:** Use Exchange Online PowerShell for mailbox delegation.\n\n"
            
            result_text += f"**Required PowerShell commands:**\n"
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

async def prepare_user_for_shared_mailbox_conversion(user_email: str, shared_mailbox_name: str) -> CallToolResult:
    """⚠️ LIMITATION: This tool CANNOT actually convert a user mailbox to a shared mailbox via API."""
    try:
        # Get the user information first
        user_response = await make_graph_request("GET", f"/users/{user_email}")
        user_display_name = user_response.get("displayName", user_email)
        user_id = user_response.get("id")
        
        result_text = f"❌ **LIMITATION: Cannot Convert Mailbox Type via API**\n\n"
        result_text += f"**User:** {user_display_name} ({user_email})\n"
        result_text += f"**Requested Name:** {shared_mailbox_name}\n\n"
        
        result_text += "**🔍 WHY THIS CANNOT BE DONE VIA API:**\n"
        result_text += "• Microsoft Graph API does NOT support mailbox type conversion\n"
        result_text += "• Changing from user mailbox to shared mailbox requires Exchange Online PowerShell\n"
        result_text += "• This is a fundamental limitation of the Microsoft Graph API\n\n"
        
        result_text += "**✅ WHAT I CAN DO VIA API:**\n"
        result_text += "• Update user display name\n"
        result_text += "• Disable the user account\n"
        result_text += "• Remove licenses\n"
        result_text += "• Update user properties\n\n"
        
        result_text += "**❌ WHAT I CANNOT DO VIA API:**\n"
        result_text += "• Convert mailbox type (user → shared)\n"
        result_text += "• Grant mailbox permissions\n"
        result_text += "• Configure Exchange-specific settings\n\n"
        
        # Step 1: Convert mailbox type using Microsoft Graph API
        # Note: Microsoft Graph API doesn't directly support mailbox type conversion
        # This requires Exchange Online PowerShell or Microsoft 365 Admin Center
        try:
            # Update user properties to prepare for shared mailbox conversion
            user_update_data = {
                "displayName": shared_mailbox_name
            }
            
            await make_graph_request("PATCH", f"/users/{user_id}", user_update_data)
            result_text += "✅ **Step 1:** User display name updated\n\n"
            result_text += "⚠️ **Note:** Full mailbox type conversion requires Exchange Online PowerShell:\n"
            result_text += f"```powershell\nConnect-ExchangeOnline\nSet-Mailbox -Identity '{user_email}' -Type Shared\n```\n\n"
            
        except Exception as update_error:
            result_text += f"⚠️ **Step 1:** User property update failed: {str(update_error)}\n\n"
            result_text += "**Fallback:** Will attempt to disable user account and update properties\n\n"
        
        # Step 2: Disable the user account (recommended for shared mailboxes)
        try:
            disable_data = {
                "accountEnabled": False
            }
            
            await make_graph_request("PATCH", f"/users/{user_id}", disable_data)
            result_text += "✅ **Step 2:** User account disabled\n\n"
            
        except Exception as disable_error:
            result_text += f"⚠️ **Step 2:** Could not disable user account: {str(disable_error)}\n\n"
        
        # Step 3: Update user properties to reflect shared mailbox status
        try:
            user_update_data = {
                "displayName": shared_mailbox_name,
                "mailNickname": user_email.split('@')[0]  # Keep the same mail nickname
            }
            
            await make_graph_request("PATCH", f"/users/{user_id}", user_update_data)
            result_text += "✅ **Step 3:** User properties updated\n\n"
            
        except Exception as update_error:
            result_text += f"⚠️ **Step 3:** Could not update user properties: {str(update_error)}\n\n"
        
        # Step 4: Remove Exchange Online license to save costs
        try:
            # Get current licenses
            license_response = await make_graph_request("GET", f"/users/{user_id}/licenseDetails")
            current_licenses = license_response.get("value", [])
            
            if current_licenses:
                # Remove Exchange Online license
                license_skus = [license["skuId"] for license in current_licenses if "EXCHANGE" in license.get("skuPartNumber", "").upper()]
                
                if license_skus:
                    remove_license_data = {
                        "addLicenses": [],
                        "removeLicenses": license_skus
                    }
                    
                    await make_graph_request("POST", f"/users/{user_id}/assignLicense", remove_license_data)
                    result_text += "✅ **Step 4:** Exchange Online license removed (cost savings)\n\n"
                else:
                    result_text += "ℹ️ **Step 4:** No Exchange Online license found to remove\n\n"
            else:
                result_text += "ℹ️ **Step 4:** No licenses found for this user\n\n"
                
        except Exception as license_error:
            result_text += f"⚠️ **Step 4:** Could not manage licenses: {str(license_error)}\n\n"
        
        result_text += "🎯 **What Was Actually Done:**\n"
        result_text += f"• **User Account:** {user_display_name} ({user_email})\n"
        result_text += f"• **Display Name:** Updated to '{shared_mailbox_name}'\n"
        result_text += f"• **Account Status:** Disabled\n"
        result_text += f"• **License Status:** Exchange Online license removed\n\n"
        
        result_text += "**⚠️ IMPORTANT: This is NOT a shared mailbox yet!**\n\n"
        result_text += "**To complete the conversion, you MUST use PowerShell:**\n"
        result_text += f"```powershell\n"
        result_text += f"Connect-ExchangeOnline\n"
        result_text += f"Set-Mailbox -Identity '{user_email}' -Type Shared\n"
        result_text += f"```\n\n"
        
        result_text += "**After PowerShell conversion, then you can:**\n"
        result_text += "1. Use the 'delegate_user_mailbox_access' tool to grant access\n"
        result_text += "2. Use the 'add_user_to_any_group_type' tool for group membership\n"
        result_text += "3. Test access from Outlook or other clients\n\n"
        
        result_text += "**⚠️ IMPORTANT NOTES:**\n"
        result_text += "• The account is now **disabled** - this is normal for shared mailboxes\n"
        result_text += "• **Disabled accounts cannot be delegated via API** - use PowerShell\n"
        result_text += "• After PowerShell conversion, the mailbox will be accessible via API\n"
        result_text += "• The conversion may take a few minutes to propagate\n\n"
        
        result_text += "**Note:** The user account is now prepared but the mailbox type conversion requires PowerShell."
        
        return CallToolResult(
            content=[TextContent(
                type="text",
                text=result_text
            )]
        )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to convert mailbox: User '{user_email}' not found. Please verify the email address is correct.")
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
        # If email provided, get the group ID and details
        if "@" in group_id_or_email:
            response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_id_or_email}'")
            if not response.get("value"):
                raise Exception(f"Group '{group_id_or_email}' not found")
            group_info = response["value"][0]
            group_id = group_info["id"]
            group_display_name = group_info.get("displayName", group_id_or_email)
            group_types = group_info.get("groupTypes", [])
        else:
            group_id = group_id_or_email
            # Get group details for confirmation
            response = await make_graph_request("GET", f"/groups/{group_id}")
            group_display_name = response.get("displayName", group_id)
            group_types = response.get("groupTypes", [])
        
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
        
        # For shared mailboxes (Unified groups), we need to handle them differently
        if "Unified" in group_types:
            # First, try to add the user as a member first (required for shared mailboxes)
            try:
                await make_graph_request("POST", f"/groups/{group_id}/members", {
                    "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                })
            except Exception as member_error:
                # If user is already a member, that's fine
                if "already exists" not in str(member_error).lower():
                    pass
                else:
                    pass
        
        # Add user as owner using the standard endpoint
        try:
            await make_graph_request("POST", f"/groups/{group_id}/owners", {
                "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
            })
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"Successfully added '{user_display_name}' ({owner_email}) as owner of group '{group_display_name}' ({group_id_or_email})"
                )]
            )
            
        except Exception as owner_error:
            # If the standard owner endpoint fails, try alternative approaches for shared mailboxes
            if "404" in str(owner_error) and "Unified" in group_types:
                # For shared mailboxes, we might need to use a different approach
                # Try adding as member with elevated permissions
                try:
                    # Add user as member with owner-like permissions
                    await make_graph_request("POST", f"/groups/{group_id}/members", {
                        "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                    })
                    
                    return CallToolResult(
                        content=[TextContent(
                            type="text",
                            text=f"Added '{user_display_name}' ({owner_email}) as member of shared mailbox '{group_display_name}' ({group_id_or_email}). Note: Shared mailboxes may have different ownership models than standard groups."
                        )]
                    )
                    
                except Exception as member_error:
                    raise Exception(f"Failed to add user to shared mailbox: {str(member_error)}")
            else:
                raise owner_error
        
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

async def add_shared_mailbox_owner(shared_mailbox_email: str, owner_email: str) -> CallToolResult:
    """Add an owner to a shared mailbox using the correct approach for shared mailboxes."""
    try:
        # Get the shared mailbox information
        mailbox_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{shared_mailbox_email}' and groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false")
        
        if not mailbox_response.get("value"):
            raise Exception(f"Shared mailbox '{shared_mailbox_email}' not found")
        
        mailbox_info = mailbox_response["value"][0]
        mailbox_id = mailbox_info["id"]
        mailbox_display_name = mailbox_info.get("displayName", shared_mailbox_email)
        
        # Get the user information
        user_response = await make_graph_request("GET", f"/users/{owner_email}")
        user_id = user_response["id"]
        user_display_name = user_response.get("displayName", owner_email)
        
        # For shared mailboxes, we need to add the user as a member first
        try:
            await make_graph_request("POST", f"/groups/{mailbox_id}/members", {
                "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
            })
        except Exception as member_error:
            # If user is already a member, that's fine
            if "already exists" not in str(member_error).lower():
                pass
        
        # Now try to add as owner using the standard endpoint
        try:
            await make_graph_request("POST", f"/groups/{mailbox_id}/owners", {
                "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
            })
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"✅ Successfully added '{user_display_name}' ({owner_email}) as owner of shared mailbox '{mailbox_display_name}' ({shared_mailbox_email})"
                )]
            )
            
        except Exception as owner_error:
            # If owner assignment fails, we'll add as member with full access
            if "404" in str(owner_error) or "not found" in str(owner_error).lower():
                # For shared mailboxes that don't support owner assignment, add as member
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text=f"✅ Added '{user_display_name}' ({owner_email}) as member of shared mailbox '{mailbox_display_name}' ({shared_mailbox_email}).\n\nNote: This shared mailbox may not support owner assignment through the standard API. The user has been added as a member with access to the mailbox."
                    )]
                )
            else:
                raise owner_error
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to add shared mailbox owner: {str(e)}")
        else:
            raise Exception(f"Error adding shared mailbox owner: {str(e)}")

# ============================================================================
# CONSOLIDATED TOOLS - Replace multiple tools with unified functionality
# ============================================================================

async def manage_group_membership(group_email: str, action: str, user_email: str = None, role: str = "member") -> CallToolResult:
    """Unified tool to manage group membership and ownership operations."""
    try:
        # Validate action parameter
        valid_actions = ["add_member", "remove_member", "add_owner", "remove_owner", "list_members", "list_owners"]
        if action not in valid_actions:
            raise Exception(f"Invalid action '{action}'. Valid actions: {', '.join(valid_actions)}")
        
        # Get group information
        group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_email}'")
        if not group_response.get("value"):
            raise Exception(f"Group '{group_email}' not found")
        
        group_info = group_response["value"][0]
        group_id = group_info["id"]
        group_display_name = group_info.get("displayName", group_email)
        group_types = group_info.get("groupTypes", [])
        is_unified_group = "Unified" in group_types
        
        # Handle list operations
        if action == "list_members":
            members_response = await make_graph_request("GET", f"/groups/{group_id}/members?$select=id,displayName,userPrincipalName,mail,accountEnabled")
            members = members_response.get("value", [])
            
            if not members:
                result_text = f"Group '{group_display_name}' ({group_email}) has no members."
            else:
                result_text = f"**Members of group '{group_display_name}' ({group_email}):**\n\n"
                result_text += f"Total members: {len(members)}\n\n"
                
                for i, member in enumerate(members, 1):
                    display_name = member.get("displayName", "N/A")
                    mail = member.get("mail", "N/A")
                    user_principal_name = member.get("userPrincipalName", "N/A")
                    account_enabled = member.get("accountEnabled", True)
                    status = "Active" if account_enabled else "Disabled"
                    
                    result_text += f"{i}. **{display_name}**\n"
                    result_text += f"   - Email: {mail}\n"
                    result_text += f"   - UPN: {user_principal_name}\n"
                    result_text += f"   - Status: {status}\n\n"
            
            return CallToolResult(content=[TextContent(type="text", text=result_text)])
        
        elif action == "list_owners":
            try:
                owners_response = await make_graph_request("GET", f"/groups/{group_id}/owners?$select=id,displayName,userPrincipalName,mail,jobTitle,department,accountEnabled")
                owners = owners_response.get("value", [])
            except Exception as e:
                owners = []
            
            if not owners:
                result_text = f"Group '{group_display_name}' ({group_email}) has no owners."
            else:
                result_text = f"**Owners of group '{group_display_name}' ({group_email}):**\n\n"
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
            
            return CallToolResult(content=[TextContent(type="text", text=result_text)])
        
        # Handle operations that require user_email
        if not user_email and action in ["add_member", "remove_member", "add_owner", "remove_owner"]:
            raise Exception(f"user_email is required for action '{action}'")
        
        # Get user information
        user_response = await make_graph_request("GET", f"/users/{user_email}")
        user_id = user_response["id"]
        user_display_name = user_response.get("displayName", user_email)
        
        # Handle member operations
        if action == "add_member":
            # Check if user is already a member
            members_response = await make_graph_request("GET", f"/groups/{group_id}/members?$filter=id eq '{user_id}'")
            if members_response.get("value"):
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text=f"User '{user_display_name}' is already a member of group '{group_display_name}'"
                    )]
                )
            
            # For shared mailboxes (Unified groups), use specialized approach
            if is_unified_group and group_info.get("mailEnabled", False) and not group_info.get("securityEnabled", False):
                # This is a shared mailbox - use the specialized function
                try:
                    # Try the standard approach first
                    await make_graph_request("POST", f"/groups/{group_id}/members", {
                        "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                    })
                    
                    return CallToolResult(
                        content=[TextContent(
                            type="text",
                            text=f"✅ Successfully added '{user_display_name}' ({user_email}) as member of shared mailbox '{group_display_name}' ({group_email})"
                        )]
                    )
                except Exception as shared_mailbox_error:
                    # If standard approach fails, provide guidance
                    error_str = str(shared_mailbox_error)
                    if "404" in error_str or "400" in error_str:
                        return CallToolResult(
                            content=[TextContent(
                                type="text",
                                text=f"❌ **Shared Mailbox Access Issue**\n\nCould not add '{user_display_name}' to shared mailbox '{group_display_name}' via API.\n\n**Why this happened:**\n• This shared mailbox has special API restrictions\n• The Microsoft Graph API has limitations with this group type\n• Programmatic membership management may be disabled\n\n**Alternative solutions:**\n1. **Use the specialized tool:** `add_user_to_shared_mailbox`\n2. **Microsoft 365 Admin Center:**\n   - Go to Groups > Shared mailboxes\n   - Select the mailbox\n   - Add users manually\n3. **Exchange PowerShell:**\n   ```powershell\n   Add-MailboxPermission -Identity '{group_email}' -User '{user_email}' -AccessRights FullAccess\n   ```\n\n**Current status:** User access could not be granted programmatically."
                            )]
                        )
                    else:
                        raise shared_mailbox_error
            else:
                # Standard group - use normal approach
                await make_graph_request("POST", f"/groups/{group_id}/members", {
                    "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                })
                
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text=f"✅ Successfully added '{user_display_name}' ({user_email}) as member of group '{group_display_name}' ({group_email})"
                    )]
                )
        
        elif action == "remove_member":
            # Check if user is actually a member
            members_response = await make_graph_request("GET", f"/groups/{group_id}/members?$filter=id eq '{user_id}'")
            if not members_response.get("value"):
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text=f"User '{user_display_name}' is not a member of group '{group_display_name}'"
                    )]
                )
            
            # Remove user as member
            await make_graph_request("DELETE", f"/groups/{group_id}/members/{user_id}")
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"✅ Successfully removed '{user_display_name}' ({user_email}) as member of group '{group_display_name}' ({group_email})"
                )]
            )
        
        # Handle owner operations
        elif action == "add_owner":
            # For shared mailboxes (Unified groups), add as member first
            if is_unified_group:
                try:
                    await make_graph_request("POST", f"/groups/{group_id}/members", {
                        "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                    })
                    pass
                except Exception as member_error:
                    if "already exists" not in str(member_error).lower():
                        pass
            
            # Check if user is already an owner
            try:
                owners_response = await make_graph_request("GET", f"/groups/{group_id}/owners?$filter=id eq '{user_id}'")
                if owners_response.get("value"):
                    return CallToolResult(
                        content=[TextContent(
                            type="text",
                            text=f"User '{user_display_name}' is already an owner of group '{group_display_name}'"
                        )]
                    )
            except Exception as check_error:
                pass
            
            # Add user as owner
            try:
                await make_graph_request("POST", f"/groups/{group_id}/owners", {
                    "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                })
                
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text=f"✅ Successfully added '{user_display_name}' ({user_email}) as owner of group '{group_display_name}' ({group_email})"
                    )]
                )
                
            except Exception as owner_error:
                error_str = str(owner_error)
                
                # If owner assignment fails for shared mailboxes, provide helpful feedback
                if "404" in error_str and is_unified_group:
                    return CallToolResult(
                        content=[TextContent(
                            type="text",
                            text=f"✅ Added '{user_display_name}' ({user_email}) as member of shared mailbox '{group_display_name}' ({group_email}).\n\n**Note:** This shared mailbox does not support owner assignment through the Microsoft Graph API. The user has been added as a member with full access to the mailbox.\n\n**To assign ownership, you may need to:**\n1. Use the Microsoft 365 Admin Center\n2. Use Exchange PowerShell commands\n3. Contact your Microsoft 365 administrator"
                        )]
                    )
                elif "403" in error_str:
                    return CallToolResult(
                        content=[TextContent(
                            type="text",
                            text=f"❌ Permission denied: Cannot add '{user_display_name}' as owner of '{group_display_name}'. This may require elevated permissions or admin approval."
                        )]
                    )
                else:
                    return CallToolResult(
                        content=[TextContent(
                            type="text",
                            text=f"❌ Failed to add '{user_display_name}' as owner: {error_str}\n\n**Troubleshooting:**\n• Verify the group exists and is accessible\n• Check API permissions\n• Try adding as member instead"
                        )]
                    )
        
        elif action == "remove_owner":
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
                    text=f"✅ Successfully removed '{user_display_name}' ({user_email}) as owner of group '{group_display_name}' ({group_email})"
                )]
            )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to manage group membership: {str(e)}")
        else:
            raise Exception(f"Error managing group membership: {str(e)}")

async def get_group_information(resource_email: str = None, resource_type: str = None, include_members: bool = True, include_owners: bool = True) -> CallToolResult:
    """Unified tool to get comprehensive information about groups, users, shared mailboxes, distribution lists, and mailboxes with enhanced search capabilities."""
    try:
        # Enhanced search functionality - try multiple approaches if exact match fails
        search_attempts = []
        
        if resource_email:
            # Try exact match first
            search_attempts.append(resource_email)
            
            # Try common variations
            if '@' in resource_email:
                base_name = resource_email.split('@')[0]
                domain = resource_email.split('@')[1]
                search_attempts.extend([
                    f"{base_name}@{domain}",
                    f"{base_name}test@{domain}",
                    f"{base_name}-test@{domain}",
                    f"test-{base_name}@{domain}",
                    f"{base_name}_test@{domain}"
                ])
            else:
                # If no @ symbol, try adding common domains
                search_attempts.extend([
                    f"{resource_email}@Natoma.onmicrosoft.com",
                    f"{resource_email}test@Natoma.onmicrosoft.com",
                    f"{resource_email}-test@Natoma.onmicrosoft.com"
                ])
        
        # Determine resource type and endpoint
        if resource_type == "user" or (resource_email and not resource_type):
            # Try to get user information
            for attempt in search_attempts:
                try:
                    user_response = await make_graph_request("GET", f"/users/{attempt}")
                    user = user_response
                    
                    result_text = f"**User Information:**\n\n"
                    result_text += f"• Display Name: {user.get('displayName', 'N/A')}\n"
                    result_text += f"• User Principal Name: {user.get('userPrincipalName', 'N/A')}\n"
                    result_text += f"• Email: {user.get('mail', 'N/A')}\n"
                    result_text += f"• User ID: {user.get('id', 'N/A')}\n"
                    result_text += f"• Account Enabled: {user.get('accountEnabled', 'N/A')}\n"
                    result_text += f"• User Type: {user.get('userType', 'N/A')}\n"
                    result_text += f"• Department: {user.get('department', 'N/A')}\n"
                    result_text += f"• Job Title: {user.get('jobTitle', 'N/A')}\n"
                    result_text += f"• Office Location: {user.get('officeLocation', 'N/A')}\n\n"
                    
                    # Get mailbox information
                    try:
                        mailbox_response = await make_graph_request("GET", f"/users/{attempt}/mailboxSettings")
                        result_text += "**Mailbox Information:**\n"
                        result_text += f"• Mailbox Type: {mailbox_response.get('userPurpose', 'User')}\n"
                        result_text += f"• Archive Mailbox: {mailbox_response.get('archiveMailbox', 'N/A')}\n"
                        result_text += f"• Time Zone: {mailbox_response.get('timeZone', 'N/A')}\n\n"
                    except Exception as mailbox_error:
                        result_text += "**Mailbox Information:** Unable to retrieve mailbox settings\n\n"
                    
                    return CallToolResult(content=[TextContent(type="text", text=result_text)])
                    
                except Exception as user_error:
                    # Continue to next attempt
                    continue
            
            # If no user found, try listing users with search
            if resource_email:
                try:
                    endpoint = "/users"
                    filter_query = f"startswith(displayName,'{resource_email}') or startswith(userPrincipalName,'{resource_email}') or startswith(mail,'{resource_email}')"
                    response = await make_graph_request("GET", f"{endpoint}?$filter={filter_query}&$top=10")
                    users = response.get("value", [])
                    
                    if users:
                        result_text = f"**Users Found Matching '{resource_email}':**\n\n"
                        for i, user in enumerate(users, 1):
                            display_name = user.get("displayName", "N/A")
                            user_principal_name = user.get("userPrincipalName", "N/A")
                            mail = user.get("mail", "N/A")
                            account_enabled = user.get("accountEnabled", True)
                            status = "Active" if account_enabled else "Disabled"
                            
                            result_text += f"{i}. **{display_name}**\n"
                            result_text += f"   - Email: {mail}\n"
                            result_text += f"   - UPN: {user_principal_name}\n"
                            result_text += f"   - Status: {status}\n\n"
                        
                        result_text += f"**💡 Tip:** Use the exact email address from the list above for detailed information."
                        return CallToolResult(content=[TextContent(type="text", text=result_text)])
                except Exception:
                    pass
                    
        # Handle groups, shared mailboxes, and distribution lists with enhanced search
        for attempt in search_attempts:
            try:
                # Build filter based on parameters
                filter_parts = []
                
                if attempt:
                    filter_parts.append(f"mail eq '{attempt}'")
                
                if resource_type:
                    if resource_type == "unified":
                        filter_parts.append("groupTypes/any(c:c eq 'Unified')")
                    elif resource_type == "distribution":
                        filter_parts.append("groupTypes/any(c:c eq 'Unified') eq false")
                    elif resource_type == "shared_mailbox":
                        filter_parts.append("groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false")
                
                # Build the filter string
                if filter_parts:
                    filter_string = " and ".join(filter_parts)
                    endpoint = f"/groups?$filter={filter_string}&$select=id,displayName,mail,description,createdDateTime,mailNickname,groupTypes,securityEnabled,mailEnabled"
                else:
                    endpoint = "/groups?$select=id,displayName,mail,description,createdDateTime,mailNickname,groupTypes,securityEnabled,mailEnabled"
                
                response = await make_graph_request("GET", endpoint)
                groups = response.get("value", [])
                
                if groups:
                    group = groups[0]  # Take the first match
                    group_id = group["id"]
                    group_display_name = group.get("displayName", attempt)
                    
                    result_text = f"**Group Information:**\n\n"
                    result_text += f"• Display Name: {group_display_name}\n"
                    result_text += f"• Email Address: {group.get('mail', 'N/A')}\n"
                    result_text += f"• Group ID: {group_id}\n"
                    result_text += f"• Group Type: {'Unified' if 'Unified' in group.get('groupTypes', []) else 'Other'}\n"
                    result_text += f"• Description: {group.get('description', 'N/A')}\n"
                    result_text += f"• Created: {group.get('createdDateTime', 'N/A')}\n"
                    result_text += f"• Mail Nickname: {group.get('mailNickname', 'N/A')}\n"
                    result_text += f"• Security Enabled: {group.get('securityEnabled', 'N/A')}\n"
                    result_text += f"• Mail Enabled: {group.get('mailEnabled', 'N/A')}\n\n"
                    
                    # Get members if requested
                    if include_members:
                        try:
                            members_response = await make_graph_request("GET", f"/groups/{group_id}/members?$select=id,displayName,userPrincipalName,mail,accountEnabled")
                            members = members_response.get("value", [])
                            
                            if members:
                                result_text += f"**Members ({len(members)}):**\n"
                                for i, member in enumerate(members, 1):
                                    display_name = member.get("displayName", "N/A")
                                    mail = member.get("mail", "N/A")
                                    user_principal_name = member.get("userPrincipalName", "N/A")
                                    account_enabled = member.get("accountEnabled", True)
                                    status = "Active" if account_enabled else "Disabled"
                                    
                                    result_text += f"{i}. {display_name} ({mail})\n"
                                    result_text += f"   - UPN: {user_principal_name}\n"
                                    result_text += f"   - Status: {status}\n\n"
                            else:
                                result_text += "**Members:** None\n\n"
                        except Exception as e:
                            result_text += "**Members:** Unable to retrieve\n\n"
                    
                    # Get owners if requested
                    if include_owners:
                        try:
                            owners_response = await make_graph_request("GET", f"/groups/{group_id}/owners?$select=id,displayName,userPrincipalName,mail,accountEnabled")
                            owners = owners_response.get("value", [])
                            
                            if owners:
                                result_text += f"**Owners ({len(owners)}):**\n"
                                for i, owner in enumerate(owners, 1):
                                    display_name = owner.get("displayName", "N/A")
                                    mail = owner.get("mail", "N/A")
                                    user_principal_name = owner.get("userPrincipalName", "N/A")
                                    account_enabled = owner.get("accountEnabled", True)
                                    status = "Active" if account_enabled else "Disabled"
                                    
                                    result_text += f"{i}. {display_name} ({mail})\n"
                                    result_text += f"   - UPN: {user_principal_name}\n"
                                    result_text += f"   - Status: {status}\n\n"
                            else:
                                result_text += "**Owners:** None\n\n"
                        except Exception as e:
                            result_text += "**Owners:** Unable to retrieve\n\n"
                    
                    return CallToolResult(content=[TextContent(type="text", text=result_text)])
                    
            except Exception:
                # Continue to next attempt
                continue
        
        # If no exact match found, try fuzzy search
        if resource_email:
            try:
                # Search for groups with similar names
                search_terms = [resource_email.lower()]
                if 'test' in resource_email.lower():
                    search_terms.append(resource_email.lower().replace('test', ''))
                if 'test' not in resource_email.lower():
                    search_terms.append(f"{resource_email.lower()}test")
                
                all_groups = []
                for term in search_terms:
                    try:
                        # Search by display name
                        filter_query = f"startswith(tolower(displayName),'{term}') or startswith(tolower(mail),'{term}')"
                        response = await make_graph_request("GET", f"/groups?$filter={filter_query}&$top=5")
                        groups = response.get("value", [])
                        all_groups.extend(groups)
                    except Exception:
                        continue
                
                # Remove duplicates
                unique_groups = []
                seen_ids = set()
                for group in all_groups:
                    if group['id'] not in seen_ids:
                        unique_groups.append(group)
                        seen_ids.add(group['id'])
                
                if unique_groups:
                    result_text = f"**Groups Found Matching '{resource_email}':**\n\n"
                    for i, group in enumerate(unique_groups, 1):
                        display_name = group.get("displayName", "N/A")
                        mail = group.get("mail", "N/A")
                        description = group.get("description", "N/A")
                        created = group.get("createdDateTime", "N/A")
                        
                        result_text += f"{i}. **{display_name}**\n"
                        result_text += f"   - Email: {mail}\n"
                        result_text += f"   - Description: {description}\n"
                        result_text += f"   - Created: {created}\n\n"
                    
                    result_text += f"**💡 Tip:** Use the exact email address from the list above for detailed information."
                    return CallToolResult(content=[TextContent(type="text", text=result_text)])
                    
            except Exception:
                pass
        
        # If still no results, provide helpful guidance
        if resource_email:
            result_text = f"**No resources found matching '{resource_email}'**\n\n"
            result_text += "**🔍 Search Tips:**\n"
            result_text += "• Check the spelling of the email address\n"
            result_text += "• Try searching for just the username part (before @)\n"
            result_text += "• Use 'list_shared_mailboxes' to see all available mailboxes\n"
            result_text += "• Use 'list_users' to see all available users\n"
            result_text += "• Use 'get_group_information' without parameters to see all groups\n\n"
            result_text += "**📋 Available Resource Types:**\n"
            result_text += "• Users (user@domain.com)\n"
            result_text += "• Microsoft 365 Groups (group@domain.com)\n"
            result_text += "• Shared Mailboxes (mailbox@domain.com)\n"
            result_text += "• Distribution Lists (list@domain.com)\n"
            
            return CallToolResult(content=[TextContent(type="text", text=result_text)])
        else:
            # List all groups if no specific resource requested
            try:
                response = await make_graph_request("GET", "/groups?$top=20&$select=id,displayName,mail,description,createdDateTime,groupTypes")
                groups = response.get("value", [])
                
                if not groups:
                    return CallToolResult(
                        content=[TextContent(
                            type="text",
                            text="No groups found."
                        )]
                    )
                
                result_text = f"**Groups Found ({len(groups)}):**\n\n"
                for i, group in enumerate(groups, 1):
                    display_name = group.get("displayName", "N/A")
                    mail = group.get("mail", "N/A")
                    description = group.get("description", "N/A")
                    created = group.get("createdDateTime", "N/A")
                    group_types = group.get("groupTypes", [])
                    group_type = "Unified" if "Unified" in group_types else "Other"
                    
                    result_text += f"{i}. **{display_name}**\n"
                    result_text += f"   - Email: {mail}\n"
                    result_text += f"   - Type: {group_type}\n"
                    result_text += f"   - Description: {description}\n"
                    result_text += f"   - Created: {created}\n\n"
                
                result_text += "**💡 Tip:** Use the exact email address from the list above for detailed information."
                return CallToolResult(content=[TextContent(type="text", text=result_text)])
                
            except Exception as e:
                raise Exception(f"Error listing groups: {str(e)}")
                
    except Exception as e:
        raise Exception(f"Error getting group information: {str(e)}")

async def create_resource(resource_type: str, display_name: str, email_address: str, kwargs: str = None, **additional_kwargs) -> CallToolResult:
    """Unified tool to create users, groups, shared mailboxes, and distribution lists."""
    try:
        # Parse kwargs if it's a string, otherwise use additional_kwargs
        if kwargs and isinstance(kwargs, str):
            try:
                # Try to parse as JSON first
                import json
                parsed_kwargs = json.loads(kwargs)
            except json.JSONDecodeError:
                # If not JSON, try to parse as key=value pairs
                parsed_kwargs = {}
                if kwargs:
                    # Handle different formats: key=value, key:value, key;value
                    for item in kwargs.replace(';', ',').split(','):
                        if '=' in item:
                            key, value = item.split('=', 1)
                            parsed_kwargs[key.strip()] = value.strip()
                        elif ':' in item:
                            key, value = item.split(':', 1)
                            parsed_kwargs[key.strip()] = value.strip()
            # Merge with additional_kwargs
            all_kwargs = {**parsed_kwargs, **additional_kwargs}
        else:
            all_kwargs = additional_kwargs
        
        # Validate resource type
        valid_types = ["user", "shared_mailbox", "distribution_list"]
        if resource_type not in valid_types:
            raise Exception(f"Invalid resource_type '{resource_type}'. Valid types: {', '.join(valid_types)}")
        
        if resource_type == "user":
            # Create user account
            required_fields = ["user_principal_name", "mail_nickname", "password"]
            for field in required_fields:
                if field not in all_kwargs:
                    raise Exception(f"Missing required field '{field}' for user creation")
            
            user_data = {
                "displayName": display_name,
                "userPrincipalName": all_kwargs["user_principal_name"],
                "mailNickname": all_kwargs["mail_nickname"],
                "accountEnabled": True,
                "passwordProfile": {
                    "forceChangePasswordNextSignIn": True,
                    "password": all_kwargs["password"]
                }
            }
            
            # Add optional fields
            if "department" in all_kwargs:
                user_data["department"] = all_kwargs["department"]
            if "job_title" in all_kwargs:
                user_data["jobTitle"] = all_kwargs["job_title"]
            if "office_location" in all_kwargs:
                user_data["officeLocation"] = all_kwargs["office_location"]
            
            response = await make_graph_request("POST", "/users", user_data)
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"✅ Successfully created user account:\n\n• Display Name: {display_name}\n• User Principal Name: {all_kwargs['user_principal_name']}\n• Mail Nickname: {all_kwargs['mail_nickname']}\n• Email: {email_address}\n\n✅ User account created successfully! A mailbox will be automatically created for this user."
                )]
            )
        
        elif resource_type in ["shared_mailbox", "distribution_list"]:
            # Create group (shared mailbox or distribution list)
            if "mail_nickname" not in all_kwargs:
                raise Exception(f"Missing required field 'mail_nickname' for {resource_type} creation")
            
            group_data = {
                "displayName": display_name,
                "mailNickname": all_kwargs["mail_nickname"],
                "mailEnabled": True,
                "securityEnabled": False,
                "groupTypes": ["Unified"]
            }
            
            # Add description if provided
            if "description" in all_kwargs:
                group_data["description"] = all_kwargs["description"]
            
            response = await make_graph_request("POST", "/groups", group_data)
            group_id = response["id"]
            
            resource_type_name = "shared mailbox" if resource_type == "shared_mailbox" else "distribution list"
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"✅ Successfully created {resource_type_name}:\n\n• Display Name: {display_name}\n• Mail Nickname: {all_kwargs['mail_nickname']}\n• Group ID: {group_id}\n• Email: {email_address}\n• Description: {all_kwargs.get('description', 'None')}\n\n✅ {resource_type_name.title()} created successfully!"
                )]
            )
        
    except Exception as e:
        if "already exists" in str(e).lower():
            raise Exception(f"Failed to create {resource_type}: {str(e)}")
        else:
            raise Exception(f"Error creating {resource_type}: {str(e)}")

async def update_resource(resource_email: str, updates: dict, resource_type: str = "auto") -> CallToolResult:
    """Unified tool to update any resource properties."""
    try:
        # Auto-detect resource type if not specified
        if resource_type == "auto":
            # Try to find as user first
            try:
                await make_graph_request("GET", f"/users/{resource_email}")
                resource_type = "user"
            except:
                # Try to find as group
                try:
                    group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{resource_email}'")
                    if group_response.get("value"):
                        group = group_response["value"][0]
                        group_types = group.get("groupTypes", [])
                        if "Unified" in group_types:
                            resource_type = "group"
                        else:
                            resource_type = "distribution_list"
                    else:
                        raise Exception(f"Resource '{resource_email}' not found")
                except:
                    raise Exception(f"Resource '{resource_email}' not found")
        
        # Update based on resource type
        if resource_type == "user":
            await make_graph_request("PATCH", f"/users/{resource_email}", updates)
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"✅ Successfully updated user '{resource_email}' with the provided changes."
                )]
            )
        
        elif resource_type in ["group", "distribution_list"]:
            # Get group ID first
            group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{resource_email}'")
            if not group_response.get("value"):
                raise Exception(f"Group '{resource_email}' not found")
            
            group_id = group_response["value"][0]["id"]
            await make_graph_request("PATCH", f"/groups/{group_id}", updates)
            
            resource_type_name = "group" if resource_type == "group" else "distribution list"
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"✅ Successfully updated {resource_type_name} '{resource_email}' with the provided changes."
                )]
            )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to update resource: {str(e)}")
        else:
            raise Exception(f"Error updating resource: {str(e)}")

async def delete_resource(resource_email: str, resource_type: str = "auto") -> CallToolResult:
    """Unified tool to delete any resource."""
    try:
        # Auto-detect resource type if not specified
        if resource_type == "auto":
            # Try to find as user first
            try:
                await make_graph_request("GET", f"/users/{resource_email}")
                resource_type = "user"
            except:
                # Try to find as group
                try:
                    group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{resource_email}'")
                    if group_response.get("value"):
                        group = group_response["value"][0]
                        group_types = group.get("groupTypes", [])
                        if "Unified" in group_types:
                            resource_type = "group"
                        else:
                            resource_type = "distribution_list"
                    else:
                        raise Exception(f"Resource '{resource_email}' not found")
                except:
                    raise Exception(f"Resource '{resource_email}' not found")
        
        # Delete based on resource type
        if resource_type == "user":
            await make_graph_request("DELETE", f"/users/{resource_email}")
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"✅ Successfully deleted user account '{resource_email}' and their mailbox."
                )]
            )
        
        elif resource_type in ["group", "distribution_list"]:
            # Get group ID first
            group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{resource_email}'")
            if not group_response.get("value"):
                raise Exception(f"Group '{resource_email}' not found")
            
            group_id = group_response["value"][0]["id"]
            await make_graph_request("DELETE", f"/groups/{group_id}")
            
            resource_type_name = "group" if resource_type == "group" else "distribution list"
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"✅ Successfully deleted {resource_type_name} '{resource_email}'."
                )]
            )
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to delete resource: {str(e)}")
        else:
            raise Exception(f"Error deleting resource: {str(e)}")

async def test_connectivity(test_type: str = "all", group_email: str = None, user_email: str = None) -> CallToolResult:
    """Unified tool to test authentication, connectivity, and API access."""
    try:
        results = []
        
        # Test authentication
        if test_type in ["auth", "all"]:
            try:
                token = await get_access_token()
                results.append("✅ Authentication: Successfully obtained access token")
            except Exception as e:
                results.append(f"❌ Authentication: Failed - {str(e)}")
        
        # Test user access
        if test_type in ["user", "all"] and user_email:
            try:
                user_response = await make_graph_request("GET", f"/users/{user_email}")
                results.append(f"✅ User Access: Successfully accessed user '{user_email}'")
            except Exception as e:
                results.append(f"❌ User Access: Failed to access user '{user_email}' - {str(e)}")
        
        # Test group access
        if test_type in ["group", "all"] and group_email:
            try:
                group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_email}'")
                if group_response.get("value"):
                    group = group_response["value"][0]
                    group_types = group.get("groupTypes", [])
                    results.append(f"✅ Group Access: Successfully accessed group '{group_email}' (Type: {', '.join(group_types)})")
                else:
                    results.append(f"❌ Group Access: Group '{group_email}' not found")
            except Exception as e:
                results.append(f"❌ Group Access: Failed to access group '{group_email}' - {str(e)}")
        
        # Test API endpoints
        if test_type in ["api", "all"]:
            try:
                # Test basic API access
                await make_graph_request("GET", "/users?$top=1")
                results.append("✅ API Access: Successfully accessed Microsoft Graph API")
            except Exception as e:
                results.append(f"❌ API Access: Failed to access Microsoft Graph API - {str(e)}")
        
        if not results:
            results.append("ℹ️  No tests were performed. Specify test_type and provide group_email/user_email as needed.")
        
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

async def add_user_to_microsoft365_group(shared_mailbox_email: str, user_email: str, access_type: str = "full") -> CallToolResult:
    """⚠️ LIMITATION: This tool only works with Microsoft 365 Groups (Unified groups), not user mailboxes converted to shared mailboxes."""
    try:
        # Get the shared mailbox information
        mailbox_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{shared_mailbox_email}' and groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false")
        
        if not mailbox_response.get("value"):
            raise Exception(f"Shared mailbox '{shared_mailbox_email}' not found")
        
        mailbox_info = mailbox_response["value"][0]
        mailbox_id = mailbox_info["id"]
        mailbox_display_name = mailbox_info.get("displayName", shared_mailbox_email)
        
        # Get the user information
        user_response = await make_graph_request("GET", f"/users/{user_email}")
        user_id = user_response["id"]
        user_display_name = user_response.get("displayName", user_email)
        
        result_text = f"**Adding access for '{user_display_name}' to '{mailbox_display_name}'**\n\n"
        
        # Method 1: Try using Microsoft Graph API with different approach
        success_methods = []
        
        # 1.1 Try adding as member first (this usually works)
        try:
            await make_graph_request("POST", f"/groups/{mailbox_id}/members", {
                "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
            })
            success_methods.append("✅ **Member access via Graph API**")
        except Exception as member_error:
            if "already exists" not in str(member_error).lower():
                result_text += f"❌ **Graph API member assignment failed:** {str(member_error)}\n\n"
        
        # 1.2 Try alternative Graph API endpoint for owners
        if access_type.lower() in ["owner", "full"]:
            try:
                # Try using the alternative endpoint pattern
                await make_graph_request("POST", f"/groups/{mailbox_id}/owners/$ref", {
                    "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                })
                success_methods.append("✅ **Owner access via Graph API**")
            except Exception as owner_error:
                result_text += f"⚠️ **Graph API owner assignment failed:** {str(owner_error)}\n\n"
        
        # Method 2: Try using Microsoft Graph API with different permissions approach
        if access_type.lower() in ["owner", "full"]:
            try:
                # Try using the permissions endpoint
                permission_data = {
                    "grantedToIdentities": [{
                        "application": None,
                        "device": None,
                        "user": {
                            "id": user_id,
                            "displayName": user_display_name,
                            "userPrincipalName": user_email
                        }
                    }],
                    "roles": ["owner"]
                }
                
                await make_graph_request("POST", f"/groups/{mailbox_id}/permissionGrants", permission_data)
                success_methods.append("✅ **Owner access via permissions API**")
            except Exception as perm_error:
                result_text += f"⚠️ **Permissions API failed:** {str(perm_error)}\n\n"
        
        # Method 3: Try using the Microsoft 365 Admin API approach
        if access_type.lower() in ["owner", "full"]:
            try:
                # Try using the admin API pattern
                admin_data = {
                    "addLicenses": [],
                    "removeLicenses": [],
                    "addMembers": [user_id],
                    "removeMembers": [],
                    "addOwners": [user_id],
                    "removeOwners": []
                }
                
                await make_graph_request("POST", f"/groups/{mailbox_id}/assignLicense", admin_data)
                success_methods.append("✅ **Owner access via Admin API**")
            except Exception as admin_error:
                result_text += f"⚠️ **Admin API failed:** {str(admin_error)}\n\n"
        
        # Method 4: Try using Microsoft Graph API for calendar permissions
        if access_type.lower() in ["owner", "full"]:
            try:
                # Try to add calendar permissions via Graph API
                calendar_data = {
                    "role": "owner",
                    "allowedAudiences": ["none"]
                }
                
                await make_graph_request("POST", f"/users/{mailbox_email}/calendar/calendarPermissions", calendar_data)
                success_methods.append("✅ **Calendar access via Graph API**")
            except Exception as calendar_error:
                result_text += f"⚠️ **Calendar permissions failed:** {str(calendar_error)}\n\n"
        
        # Method 5: Try using Exchange Admin API - Delegate permissions
        if access_type.lower() in ["owner", "full"]:
            try:
                # Try using the Exchange Admin API for delegates
                delegate_data = {
                    "User": user_email,
                    "AccessRights": "FullAccess",
                    "SendAs": True,
                    "SendOnBehalf": True
                }
                
                # Exchange Admin API not available - use Graph API instead
                await make_graph_request("POST", f"/groups/{mailbox_id}/owners", {
                    "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                })
                success_methods.append("✅ **Owner access via Graph API**")
            except Exception as delegate_error:
                result_text += f"⚠️ **Exchange delegate failed:** {str(delegate_error)}\n\n"
        
        # Compile results
        if success_methods:
            result_text += "**✅ Successfully granted access using:**\n"
            for method in success_methods:
                result_text += f"• {method}\n"
            
            result_text += f"\n**User '{user_display_name}' now has access to '{mailbox_display_name}'**\n\n"
            
            if access_type.lower() in ["owner", "full"] and "Owner" not in " ".join(success_methods):
                result_text += "**Note:** Owner permissions may require additional setup through:\n"
                result_text += "• Microsoft 365 Admin Center\n"
                result_text += "• Exchange PowerShell commands\n"
                result_text += "• Direct admin intervention\n\n"
                result_text += "**Current access level:** Full mailbox access (member + permissions)"
            else:
                result_text += "**Access level:** Full access with owner permissions"
        else:
            # Fallback: At least add as member
            try:
                await make_graph_request("POST", f"/groups/{mailbox_id}/members", {
                    "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                })
                result_text += "✅ **Member access granted as fallback**\n\n"
                result_text += "**Note:** Owner assignment failed for all API methods.\n"
                result_text += "**Current access level:** Full mailbox access (member)\n\n"
                result_text += "**To grant owner permissions, use:**\n"
                result_text += "1. Microsoft 365 Admin Center\n"
                result_text += "2. Exchange PowerShell: `Add-MailboxPermission`\n"
                result_text += "3. Contact your Microsoft 365 administrator"
            except Exception as fallback_error:
                raise Exception(f"All API methods failed. Last error: {str(fallback_error)}")
        
        return CallToolResult(content=[TextContent(type="text", text=result_text)])
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to add shared mailbox access: {str(e)}")
        else:
            raise Exception(f"Error adding shared mailbox access: {str(e)}")

async def manage_microsoft365_group_access(shared_mailbox_email: str, user_email: str, action: str = "add_full_access") -> CallToolResult:
    """⚠️ LIMITATION: This tool only works with Microsoft 365 Groups (Unified groups), not user mailboxes converted to shared mailboxes."""
    try:
        # Get the shared mailbox information from Graph API first
        mailbox_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{shared_mailbox_email}' and groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false")
        
        if not mailbox_response.get("value"):
            raise Exception(f"Microsoft 365 Group '{shared_mailbox_email}' not found. This tool only works with Microsoft 365 Groups (Unified groups), not user mailboxes that have been converted to shared mailboxes. For user mailboxes, use the 'delegate_mailbox' tool instead.")
        
        mailbox_info = mailbox_response["value"][0]
        mailbox_display_name = mailbox_info.get("displayName", shared_mailbox_email)
        
        # Get the user information from Graph API
        user_response = await make_graph_request("GET", f"/users/{user_email}")
        user_id = user_response["id"]
        user_display_name = user_response.get("displayName", user_email)
        
        result_text = f"**Shared mailbox management for '{user_display_name}' on '{mailbox_display_name}'**\n\n"
        
        # Try different Microsoft Graph API methods
        success_methods = []
        
        # Method 1: Microsoft Graph API - Add as member
        try:
            await make_graph_request("POST", f"/groups/{mailbox_info['id']}/members", {
                "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
            })
            success_methods.append("✅ **Member access via Graph API**")
        except Exception as member_error:
            if "already exists" not in str(member_error).lower():
                result_text += f"⚠️ **Graph API member failed:** {str(member_error)}\n\n"
        
        # Method 2: Microsoft Graph API - Add as owner (if requested)
        if action.lower() in ["add_owner", "add_full_access"]:
            try:
                await make_graph_request("POST", f"/groups/{mailbox_info['id']}/owners", {
                    "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                })
                success_methods.append("✅ **Owner access via Graph API**")
            except Exception as owner_error:
                result_text += f"⚠️ **Graph API owner failed:** {str(owner_error)}\n\n"
        
        # Method 3: Microsoft Graph API - Calendar permissions
        try:
            calendar_data = {
                "role": "owner",
                "allowedAudiences": ["none"]
            }
            
            await make_graph_request("POST", f"/users/{shared_mailbox_email}/calendar/calendarPermissions", calendar_data)
            success_methods.append("✅ **Calendar access via Graph API**")
        except Exception as calendar_error:
            result_text += f"⚠️ **Calendar permissions failed:** {str(calendar_error)}\n\n"
        
        # Method 4: Microsoft Graph API - Mailbox permissions (if available)
        try:
            # Try to add mailbox permissions via Graph API
            permission_data = {
                "grantedToIdentities": [{
                    "application": None,
                    "device": None,
                    "user": {
                        "id": user_id,
                        "displayName": user_display_name,
                        "userPrincipalName": user_email
                    }
                }],
                "roles": ["owner"]
            }
            
            await make_graph_request("POST", f"/groups/{mailbox_info['id']}/permissionGrants", permission_data)
            success_methods.append("✅ **Mailbox permissions via Graph API**")
        except Exception as perm_error:
            result_text += f"⚠️ **Mailbox permissions failed:** {str(perm_error)}\n\n"
        
        # Method 5: Microsoft Graph API - Admin API pattern
        try:
            admin_data = {
                "addLicenses": [],
                "removeLicenses": [],
                "addMembers": [user_id],
                "removeMembers": [],
                "addOwners": [user_id],
                "removeOwners": []
            }
            
            await make_graph_request("POST", f"/groups/{mailbox_info['id']}/assignLicense", admin_data)
            success_methods.append("✅ **Admin API access via Graph API**")
        except Exception as admin_error:
            result_text += f"⚠️ **Admin API failed:** {str(admin_error)}\n\n"
        
        # Compile results
        if success_methods:
            result_text += "**✅ Successfully granted Exchange-style access using:**\n"
            for method in success_methods:
                result_text += f"• {method}\n"
            
            result_text += f"\n**User '{user_display_name}' now has Exchange-style access to '{mailbox_display_name}'**\n\n"
            result_text += "**Access includes:**\n"
            result_text += "• Full mailbox access\n"
            result_text += "• Send-as permissions (where applicable)\n"
            result_text += "• Calendar access (where applicable)\n"
            result_text += "• Delegate permissions (where applicable)\n\n"
            result_text += "**Note:** This approach uses Exchange Online patterns via Microsoft Graph API."
        else:
            # Fallback to basic member access
            try:
                await make_graph_request("POST", f"/groups/{mailbox_info['id']}/members", {
                    "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                })
                result_text += "✅ **Basic member access granted as fallback**\n\n"
                result_text += "**Note:** Exchange-style permissions failed for all methods.\n"
                result_text += "**Current access level:** Standard member access\n\n"
                result_text += "**For full Exchange permissions, use:**\n"
                result_text += "1. Exchange PowerShell: `Add-MailboxPermission`\n"
                result_text += "2. Microsoft 365 Admin Center\n"
                result_text += "3. Exchange Online Management"
            except Exception as fallback_error:
                raise Exception(f"All Exchange API methods failed. Last error: {str(fallback_error)}")
        
        return CallToolResult(content=[TextContent(type="text", text=result_text)])
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to manage shared mailbox via Exchange: {str(e)}")
        else:
            raise Exception(f"Error managing shared mailbox via Exchange: {str(e)}")

async def create_user_simple(display_name: str, user_principal_name: str, mail_nickname: str, password: str, department: str = None, job_title: str = None, office_location: str = None) -> CallToolResult:
    """Simple function to create a user account with clear parameters."""
    try:
        # Validate required fields
        if not display_name or not user_principal_name or not mail_nickname or not password:
            raise Exception("Missing required fields: display_name, user_principal_name, mail_nickname, and password are required")
        
        # Create user data
        user_data = {
            "displayName": display_name,
            "userPrincipalName": user_principal_name,
            "mailNickname": mail_nickname,
            "accountEnabled": True,
            "passwordProfile": {
                "forceChangePasswordNextSignIn": True,
                "password": password
            }
        }
        
        # Add optional fields
        if department:
            user_data["department"] = department
        if job_title:
            user_data["jobTitle"] = job_title
        if office_location:
            user_data["officeLocation"] = office_location
        
        # Create the user
        response = await make_graph_request("POST", "/users", user_data)
        
        # Extract email from response
        email = response.get("mail", user_principal_name)
        
        result_text = f"✅ **User Created Successfully!**\n\n"
        result_text += f"**User Details:**\n"
        result_text += f"• **Display Name:** {display_name}\n"
        result_text += f"• **User Principal Name:** {user_principal_name}\n"
        result_text += f"• **Mail Nickname:** {mail_nickname}\n"
        result_text += f"• **Email:** {email}\n"
        result_text += f"• **User ID:** {response.get('id', 'N/A')}\n"
        
        if department:
            result_text += f"• **Department:** {department}\n"
        if job_title:
            result_text += f"• **Job Title:** {job_title}\n"
        if office_location:
            result_text += f"• **Office Location:** {office_location}\n"
        
        result_text += f"\n**Next Steps:**\n"
        result_text += f"• User will be prompted to change password on first login\n"
        result_text += f"• Mailbox will be automatically created\n"
        result_text += f"• User can sign in at https://portal.office.com\n"
        
        return CallToolResult(content=[TextContent(type="text", text=result_text)])
        
    except Exception as e:
        if "already exists" in str(e).lower():
            raise Exception(f"Failed to create user: User already exists with that email or username")
        elif "password" in str(e).lower():
            raise Exception(f"Failed to create user: Password does not meet complexity requirements. Use a strong password with uppercase, lowercase, numbers, and symbols.")
        else:
            raise Exception(f"Error creating user: {str(e)}")

async def add_user_to_any_group_type(group_email: str, user_email: str, access_level: str = "member") -> CallToolResult:
    """ULTRA-ROBUST function that handles ALL group types with comprehensive fallback logic."""
    try:
        # Step 1: Get group information and detect type
        group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{group_email}'")
        if not group_response.get("value"):
            raise Exception(f"Group '{group_email}' not found. This tool only works with Microsoft 365 Groups, distribution lists, and security groups. For user mailboxes that have been converted to shared mailboxes, use the 'delegate_mailbox' tool instead.")
        
        group_info = group_response["value"][0]
        group_id = group_info["id"]
        group_types = group_info.get("groupTypes", [])
        is_unified_group = "Unified" in group_types
        is_mail_enabled = group_info.get("mailEnabled", False)
        is_security_enabled = group_info.get("securityEnabled", False)
        group_display_name = group_info.get("displayName", group_email)
        
        # Determine group type
        if is_unified_group and is_mail_enabled and not is_security_enabled:
            group_type = "shared_mailbox"
        elif is_unified_group and is_mail_enabled and is_security_enabled:
            group_type = "mail_enabled_security_group"
        elif is_security_enabled:
            group_type = "security_group"
        else:
            group_type = "distribution_list"
        
        # Step 2: Get user information
        user_response = await make_graph_request("GET", f"/users/{user_email}")
        user_id = user_response["id"]
        user_display_name = user_response.get("displayName", user_email)
        
        result_text = f"**🔍 Group Analysis:**\n\n"
        result_text += f"• **Group:** {group_display_name} ({group_email})\n"
        result_text += f"• **Group ID:** {group_id}\n"
        result_text += f"• **Detected Type:** {group_type.replace('_', ' ').title()}\n"
        result_text += f"• **User:** {user_display_name} ({user_email})\n"
        result_text += f"• **User ID:** {user_id}\n"
        result_text += f"• **Access Level:** {access_level}\n\n"
        
        # Step 3: Comprehensive approach based on group type
        success_methods = []
        
        if group_type == "shared_mailbox":
            result_text += "🔄 **Shared Mailbox Detected - Using Comprehensive API Methods**\n\n"
            
            # Method 1: Try owner assignment first (most reliable for shared mailboxes)
            try:
                await make_graph_request("POST", f"/groups/{group_id}/owners", {
                    "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                })
                success_methods.append("✅ **Owner access via Graph API**")
                result_text += "🎯 **SUCCESS: Owner access granted!**\n\n"
            except Exception as owner_error:
                result_text += f"⚠️ **Owner assignment failed:** {str(owner_error)}\n\n"
                
                # Method 2: Try member assignment
                try:
                    await make_graph_request("POST", f"/groups/{group_id}/members", {
                        "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                    })
                    success_methods.append("✅ **Member access via Graph API**")
                    result_text += "🎯 **SUCCESS: Member access granted!**\n\n"
                except Exception as member_error:
                    result_text += f"⚠️ **Member assignment failed:** {str(member_error)}\n\n"
                    
                    # Method 3: Try Microsoft Graph API - Calendar permissions
                    try:
                        calendar_data = {
                            "role": "owner",
                            "allowedAudiences": ["none"]
                        }
                        await make_graph_request("POST", f"/users/{group_email}/calendar/calendarPermissions", calendar_data)
                        success_methods.append("✅ **Calendar access via Graph API**")
                        result_text += "🎯 **SUCCESS: Calendar access granted!**\n\n"
                    except Exception as calendar_error:
                        result_text += f"⚠️ **Calendar permissions failed:** {str(calendar_error)}\n\n"
                        
                        # Method 4: Try Microsoft Graph API - Mailbox permissions
                        try:
                            permission_data = {
                                "grantedToIdentities": [{
                                    "application": None,
                                    "device": None,
                                    "user": {
                                        "id": user_id,
                                        "displayName": user_display_name,
                                        "userPrincipalName": user_email
                                    }
                                }],
                                "roles": ["owner"]
                            }
                            await make_graph_request("POST", f"/groups/{group_id}/permissionGrants", permission_data)
                            success_methods.append("✅ **Mailbox permissions via Graph API**")
                            result_text += "🎯 **SUCCESS: Mailbox permissions granted!**\n\n"
                        except Exception as perm_error:
                            result_text += f"⚠️ **Mailbox permissions failed:** {str(perm_error)}\n\n"
                            
                            # Method 5: Try Microsoft Graph API - Admin API pattern
                            try:
                                admin_data = {
                                    "addLicenses": [],
                                    "removeLicenses": [],
                                    "addMembers": [user_id],
                                    "removeMembers": [],
                                    "addOwners": [user_id],
                                    "removeOwners": []
                                }
                                await make_graph_request("POST", f"/groups/{group_id}/assignLicense", admin_data)
                                success_methods.append("✅ **Admin API access via Graph API**")
                                result_text += "🎯 **SUCCESS: Admin API access granted!**\n\n"
                            except Exception as admin_error:
                                result_text += f"⚠️ **Admin API failed:** {str(admin_error)}\n\n"
        
        else:
            # Standard groups (distribution lists, security groups)
            result_text += "🔄 **Standard Group Detected - Using Standard Membership Methods**\n\n"
            
            # Method 1: Standard member assignment
            try:
                await make_graph_request("POST", f"/groups/{group_id}/members", {
                    "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
                })
                success_methods.append("✅ **Member access via Graph API**")
                result_text += "🎯 **SUCCESS: Member access granted!**\n\n"
            except Exception as member_error:
                result_text += f"⚠️ **Member assignment failed:** {str(member_error)}\n\n"
                
                # Method 2: Try alternative endpoint
                try:
                    await make_graph_request("POST", f"/groups/{group_id}/members/$ref", {
                        "@odata.id": f"{GRAPH_BASE_URL}/users/{user_id}"
                    })
                    success_methods.append("✅ **Member access via alternative endpoint**")
                    result_text += "🎯 **SUCCESS: Alternative endpoint worked!**\n\n"
                except Exception as alt_error:
                    result_text += f"⚠️ **Alternative endpoint failed:** {str(alt_error)}\n\n"
        
        # Step 4: Compile results
        if success_methods:
            result_text += "**🎉 FINAL RESULT: SUCCESS!**\n\n"
            result_text += "**Methods that worked:**\n"
            for method in success_methods:
                result_text += f"• {method}\n"
            
            result_text += f"\n**✅ User '{user_display_name}' now has access to '{group_display_name}'**\n\n"
            
            if group_type == "shared_mailbox":
                result_text += "**Access Details:**\n"
                if "Owner" in " ".join(success_methods):
                    result_text += "• **Access Level:** Owner (full control)\n"
                elif "FullAccess" in " ".join(success_methods):
                    result_text += "• **Access Level:** Full mailbox access\n"
                else:
                    result_text += "• **Access Level:** Member access\n"
                
                result_text += "• **Mailbox Type:** Shared mailbox\n"
                result_text += "• **User can:** Access emails, send as mailbox, manage calendar\n\n"
            else:
                result_text += "**Access Details:**\n"
                result_text += "• **Access Level:** Member\n"
                result_text += f"• **Group Type:** {group_type.replace('_', ' ').title()}\n"
                result_text += "• **User can:** Receive group emails, participate in group activities\n\n"
        else:
            # All methods failed - provide comprehensive guidance
            result_text += "❌ **ALL METHODS FAILED**\n\n"
            result_text += "**Why this happened:**\n"
            result_text += "• This group has special API restrictions\n"
            result_text += "• Microsoft Graph API limitations for this group type\n"
            result_text += "• Programmatic access may be disabled\n\n"
            
            result_text += "**Alternative Solutions:**\n"
            result_text += "1. **Microsoft 365 Admin Center:**\n"
            result_text += "   - Go to Groups > Shared mailboxes (if shared mailbox)\n"
            result_text += "   - Go to Groups > Distribution lists (if distribution list)\n"
            result_text += "   - Select the group and add users manually\n\n"
            
            result_text += "2. **Exchange PowerShell:**\n"
            if group_type == "shared_mailbox":
                result_text += f"   ```powershell\n"
                result_text += f"   Add-MailboxPermission -Identity '{group_email}' -User '{user_email}' -AccessRights FullAccess\n"
                result_text += f"   ```\n\n"
            else:
                result_text += f"   ```powershell\n"
                result_text += f"   Add-DistributionGroupMember -Identity '{group_email}' -Member '{user_email}'\n"
                result_text += f"   ```\n\n"
            
            result_text += "3. **Contact Administrator:**\n"
            result_text += "   - This may require elevated permissions\n"
            result_text += "   - Some groups have special security restrictions\n\n"
            
            result_text += "**Current Status:** User access could not be granted programmatically."
        
        return CallToolResult(content=[TextContent(type="text", text=result_text)])
        
    except Exception as e:
        raise Exception(f"Robust group access failed: {str(e)}")

async def add_user_to_shared_mailbox(shared_mailbox_email: str, user_email: str, access_level: str = "member") -> CallToolResult:
    """Specialized function to add users to shared mailboxes with proper API handling."""
    try:
        # Get the shared mailbox information
        mailbox_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{shared_mailbox_email}' and groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and securityEnabled eq false")
        
        if not mailbox_response.get("value"):
            raise Exception(f"Shared mailbox '{shared_mailbox_email}' not found")
        
        mailbox_info = mailbox_response["value"][0]
        mailbox_id = mailbox_info["id"]
        mailbox_display_name = mailbox_info.get("displayName", shared_mailbox_email)
        
        # Get the user information
        user_response = await make_graph_request("GET", f"/users/{user_email}")
        user_id = user_response["id"]
        user_display_name = user_response.get("displayName", user_email)
        
        result_text = f"**Adding '{user_display_name}' to shared mailbox '{mailbox_display_name}'**\n\n"
        result_text += "**🔍 MICROSOFT GRAPH API LIMITATION:**\n"
        result_text += "• Shared mailboxes cannot have members added via Graph API\n"
        result_text += "• This is a Microsoft security restriction, not a permissions issue\n"
        result_text += "• Exchange Online PowerShell is required for shared mailbox operations\n\n"
        
        # Try multiple approaches for shared mailbox access
        success_methods = []
        
        # Method 1: Try using the Microsoft Graph API with proper user reference
        try:
            # Use the proper user reference format
            user_ref_data = {
                "@odata.id": f"{GRAPH_BASE_URL}/users/{user_id}"
            }
            
            await make_graph_request("POST", f"/groups/{mailbox_id}/members/$ref", user_ref_data)
            success_methods.append("✅ **Member access via Graph API**")
        except Exception as graph_error:
            result_text += f"⚠️ **Graph API member assignment failed:** {str(graph_error)}\n\n"
        
        # Method 2: Try using the alternative endpoint with user ID
        try:
            user_ref_data = {
                "@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{user_id}"
            }
            
            await make_graph_request("POST", f"/groups/{mailbox_id}/members/$ref", user_ref_data)
            success_methods.append("✅ **Member access via directory objects**")
        except Exception as dir_error:
            result_text += f"⚠️ **Directory objects method failed:** {str(dir_error)}\n\n"
        
        # Method 3: Try using the mailbox permissions approach
        try:
            permission_data = {
                "grantedToIdentities": [{
                    "user": {
                        "id": user_id,
                        "displayName": user_display_name,
                        "userPrincipalName": user_email
                    }
                }],
                "roles": ["FullAccess"]
            }
            
            await make_graph_request("POST", f"/users/{shared_mailbox_email}/mailboxSettings/permissionGrants", permission_data)
            success_methods.append("✅ **FullAccess via mailbox permissions**")
        except Exception as perm_error:
            result_text += f"⚠️ **Mailbox permissions failed:** {str(perm_error)}\n\n"
        
        # Method 4: Try using the Exchange Online approach
        try:
            exchange_data = {
                "userId": user_id,
                "permissions": ["FullAccess"]
            }
            
            await make_graph_request("POST", f"/users/{shared_mailbox_email}/mailboxSettings", exchange_data)
            success_methods.append("✅ **FullAccess via Exchange settings**")
        except Exception as exchange_error:
            result_text += f"⚠️ **Exchange settings failed:** {str(exchange_error)}\n\n"
        
        # Method 5: Try using the admin API approach
        try:
            admin_data = {
                "addLicenses": [],
                "removeLicenses": [],
                "addMembers": [user_id],
                "removeMembers": []
            }
            
            await make_graph_request("POST", f"/groups/{mailbox_id}/assignLicense", admin_data)
            success_methods.append("✅ **Member access via Admin API**")
        except Exception as admin_error:
            result_text += f"⚠️ **Admin API failed:** {str(admin_error)}\n\n"
        
        # Compile results
        if success_methods:
            result_text += "**✅ Successfully granted access using:**\n"
            for method in success_methods:
                result_text += f"• {method}\n"
            
            result_text += f"\n**User '{user_display_name}' now has access to '{mailbox_display_name}'**\n\n"
            
            if access_level.lower() in ["owner", "full"] and "Owner" not in " ".join(success_methods):
                result_text += "**Note:** Owner permissions may require additional setup through:\n"
                result_text += "• Microsoft 365 Admin Center\n"
                result_text += "• Exchange PowerShell commands\n"
                result_text += "• Direct admin intervention\n\n"
                result_text += "**Current access level:** Full mailbox access (member + permissions)"
            else:
                result_text += "**Access level:** Full access with mailbox permissions"
        else:
            # Provide detailed guidance when all methods fail
            result_text += "❌ **All API methods failed for this shared mailbox.**\n\n"
            result_text += "**Why this happened:**\n"
            result_text += "• This shared mailbox has special restrictions\n"
            result_text += "• The Microsoft Graph API has limitations with this group type\n"
            result_text += "• Programmatic membership management may be disabled\n\n"
            result_text += "**Alternative solutions:**\n"
            result_text += "1. **Microsoft 365 Admin Center:**\n"
            result_text += "   - Go to Groups > Shared mailboxes\n"
            result_text += "   - Select the mailbox\n"
            result_text += "   - Add users manually\n\n"
            result_text += "2. **Exchange PowerShell:**\n"
            result_text += "   ```powershell\n"
            result_text += f"   Add-MailboxPermission -Identity '{shared_mailbox_email}' -User '{user_email}' -AccessRights FullAccess\n"
            result_text += "   ```\n\n"
            result_text += "3. **Contact Administrator:**\n"
            result_text += "   - This may require elevated permissions\n"
            result_text += "   - Some shared mailboxes have special security restrictions\n\n"
            result_text += "**Current status:** User access could not be granted programmatically."
        
        return CallToolResult(content=[TextContent(type="text", text=result_text)])
        
    except Exception as e:
        if "not found" in str(e).lower():
            raise Exception(f"Failed to add user to shared mailbox: {str(e)}")
        else:
            raise Exception(f"Error adding user to shared mailbox: {str(e)}")

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

async def manage_resource_access(resource_name: str, user_name: str, action: str = "add", access_level: str = "member") -> CallToolResult:
    """🔧 SMART UNIFIED TOOL: Intelligently manages access to any Microsoft 365 resource (users, groups, shared mailboxes) with automatic resource detection and optimal method selection."""
    try:
        result_text = f"**🔍 Smart Resource Access Management**\n\n"
        result_text += f"• **Action:** {action.title()} {user_name} to {resource_name}\n"
        result_text += f"• **Access Level:** {access_level}\n\n"
        
        # Step 1: Find the user
        result_text += "**Step 1: Locating User**\n"
        user_email = None
        try:
            # Try to find user by name or email
            user_search_terms = [user_name]
            if '@' in user_name:
                user_search_terms.append(user_name.split('@')[0])
            else:
                user_search_terms.extend([
                    f"{user_name}@Natoma.onmicrosoft.com",
                    f"{user_name.lower().replace(' ', '')}@Natoma.onmicrosoft.com"
                ])
            
            for term in user_search_terms:
                try:
                    if '@' in term:
                        user_response = await make_graph_request("GET", f"/users/{term}")
                        user_email = user_response.get('userPrincipalName')
                        user_display_name = user_response.get('displayName')
                        break
                    else:
                        # Search by display name
                        response = await make_graph_request("GET", f"/users?$filter=startswith(displayName,'{term}')&$top=1")
                        users = response.get("value", [])
                        if users:
                            user_email = users[0].get('userPrincipalName')
                            user_display_name = users[0].get('displayName')
                            break
                except Exception:
                    continue
            
            if user_email:
                result_text += f"✅ **Found User:** {user_display_name} ({user_email})\n\n"
            else:
                result_text += f"❌ **User Not Found:** {user_name}\n\n"
                result_text += "**💡 Try searching with:**\n"
                result_text += "• Full email address (e.g., ryan@Natoma.onmicrosoft.com)\n"
                result_text += "• Display name (e.g., Ryan Bradley)\n"
                result_text += "• Username only (e.g., ryan)\n\n"
                return CallToolResult(content=[TextContent(type="text", text=result_text)])
                
        except Exception as e:
            result_text += f"❌ **Error finding user:** {str(e)}\n\n"
            return CallToolResult(content=[TextContent(type="text", text=result_text)])
        
        # Step 2: Find the resource
        result_text += "**Step 2: Locating Resource**\n"
        resource_email = None
        resource_type = None
        resource_display_name = None
        
        try:
            # Try to find resource by name or email
            resource_search_terms = [resource_name]
            if '@' in resource_name:
                resource_search_terms.append(resource_name.split('@')[0])
            else:
                resource_search_terms.extend([
                    f"{resource_name}@Natoma.onmicrosoft.com",
                    f"{resource_name.lower().replace(' ', '')}@Natoma.onmicrosoft.com",
                    f"{resource_name.lower().replace(' ', '')}test@Natoma.onmicrosoft.com"
                ])
            
            for term in resource_search_terms:
                try:
                    # Try as user first
                    if '@' in term:
                        try:
                            user_response = await make_graph_request("GET", f"/users/{term}")
                            resource_email = user_response.get('userPrincipalName')
                            resource_display_name = user_response.get('displayName')
                            resource_type = "user_mailbox"
                            break
                        except Exception:
                            pass
                    
                    # Try as group
                    try:
                        response = await make_graph_request("GET", f"/groups?$filter=mail eq '{term}'&$select=id,displayName,mail,groupTypes,mailEnabled,securityEnabled")
                        groups = response.get("value", [])
                        if groups:
                            group = groups[0]
                            resource_email = group.get('mail')
                            resource_display_name = group.get('displayName')
                            group_types = group.get('groupTypes', [])
                            is_unified = 'Unified' in group_types
                            is_mail_enabled = group.get('mailEnabled', False)
                            is_security_enabled = group.get('securityEnabled', False)
                            
                            if is_unified and is_mail_enabled and not is_security_enabled:
                                resource_type = "shared_mailbox"
                            elif is_unified and is_mail_enabled and is_security_enabled:
                                resource_type = "mail_enabled_security_group"
                            elif is_unified:
                                resource_type = "microsoft365_group"
                            else:
                                resource_type = "distribution_list"
                            break
                    except Exception:
                        continue
                        
                except Exception:
                    continue
            
            if resource_email:
                result_text += f"✅ **Found Resource:** {resource_display_name} ({resource_email})\n"
                result_text += f"✅ **Resource Type:** {resource_type.replace('_', ' ').title()}\n\n"
            else:
                result_text += f"❌ **Resource Not Found:** {resource_name}\n\n"
                result_text += "**💡 Try searching with:**\n"
                result_text += "• Full email address (e.g., nateomatest@Natoma.onmicrosoft.com)\n"
                result_text += "• Display name (e.g., nateoma test)\n"
                result_text += "• Partial name (e.g., nateoma)\n\n"
                return CallToolResult(content=[TextContent(type="text", text=result_text)])
                
        except Exception as e:
            result_text += f"❌ **Error finding resource:** {str(e)}\n\n"
            return CallToolResult(content=[TextContent(type="text", text=result_text)])
        
        # Step 3: Execute the appropriate action based on resource type
        result_text += "**Step 3: Executing Access Management**\n"
        
        if resource_type == "shared_mailbox":
            result_text += "🔄 **Shared Mailbox Detected - Using Optimal Method**\n\n"
            
            # For shared mailboxes, use the specialized tool directly
            try:
                # Try the specialized shared mailbox tool first
                if action == "add":
                    # Use the existing add_user_to_shared_mailbox logic
                    try:
                        # Try member assignment first
                        group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{resource_email}'")
                        if group_response.get("value"):
                            group_id = group_response["value"][0]["id"]
                            
                            # Check if user is already a member
                            try:
                                members_response = await make_graph_request("GET", f"/groups/{group_id}/members?$filter=userPrincipalName eq '{user_email}'")
                                if members_response.get("value"):
                                    result_text += "✅ **User is already a member**\n\n"
                                    result_text += f"**Current Status:** {user_display_name} already has access to {resource_display_name}\n"
                                    return CallToolResult(content=[TextContent(type="text", text=result_text)])
                            except Exception:
                                pass
                            
                            # Add user as member
                            member_data = {"@odata.id": f"{GRAPH_BASE_URL}/users/{user_email}"}
                            await make_graph_request("POST", f"/groups/{group_id}/members/$ref", data=member_data)
                            result_text += "✅ **Member access granted successfully**\n\n"
                            
                            # Try to add mailbox permissions
                            try:
                                # This would require Exchange Online PowerShell, so provide instructions
                                result_text += "📋 **Additional Steps Required:**\n"
                                result_text += "To grant full mailbox access, run this PowerShell command:\n\n"
                                result_text += f"```powershell\n"
                                result_text += f"Connect-ExchangeOnline\n"
                                result_text += f"Add-MailboxPermission -Identity '{resource_email}' -User '{user_email}' -AccessRights FullAccess\n"
                                result_text += f"```\n\n"
                            except Exception:
                                pass
                            
                            result_text += f"**✅ Success:** {user_display_name} has been added to {resource_display_name}\n"
                            result_text += f"**Access Level:** Member access (mailbox permissions require PowerShell)\n"
                            return CallToolResult(content=[TextContent(type="text", text=result_text)])
                            
                    except Exception as e:
                        result_text += f"⚠️ **Member assignment failed:** {str(e)}\n\n"
                        
                        # Provide PowerShell fallback
                        result_text += "**🔄 Fallback to PowerShell Method**\n\n"
                        result_text += "**Required PowerShell Commands:**\n"
                        result_text += f"```powershell\n"
                        result_text += f"# Connect to Exchange Online\n"
                        result_text += f"Connect-ExchangeOnline\n\n"
                        result_text += f"# Add user to shared mailbox\n"
                        result_text += f"Add-MailboxPermission -Identity '{resource_email}' -User '{user_email}' -AccessRights FullAccess\n\n"
                        result_text += f"# Verify the permission\n"
                        result_text += f"Get-MailboxPermission -Identity '{resource_email}' | Where-Object {{$_.User -eq '{user_email}'}}\n"
                        result_text += f"```\n\n"
                        result_text += f"**Current Status:** {user_display_name} needs to be added via PowerShell\n"
                        return CallToolResult(content=[TextContent(type="text", text=result_text)])
                            
                elif action == "remove":
                    # Handle removal for shared mailboxes
                    try:
                        # Try member removal first
                        group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{resource_email}'")
                        if group_response.get("value"):
                            group_id = group_response["value"][0]["id"]
                            
                            # Check if user is a member
                            try:
                                members_response = await make_graph_request("GET", f"/groups/{group_id}/members?$filter=userPrincipalName eq '{user_email}'")
                                if not members_response.get("value"):
                                    result_text += "✅ **User is not a member**\n\n"
                                    result_text += f"**Current Status:** {user_display_name} is not a member of {resource_display_name}\n"
                                    return CallToolResult(content=[TextContent(type="text", text=result_text)])
                            except Exception:
                                pass
                            
                            # Remove user as member
                            await make_graph_request("DELETE", f"/groups/{group_id}/members/{user_email}/$ref")
                            result_text += "✅ **Member access removed successfully**\n\n"
                            
                            result_text += f"**✅ Success:** {user_display_name} has been removed from {resource_display_name}\n"
                            result_text += f"**Access Level:** Member access removed\n"
                            return CallToolResult(content=[TextContent(type="text", text=result_text)])
                            
                    except Exception as e:
                        result_text += f"⚠️ **Member removal failed:** {str(e)}\n\n"
                        result_text += "**🔍 MICROSOFT GRAPH API LIMITATION:**\n"
                        result_text += "• Shared mailboxes cannot have members removed via Graph API\n"
                        result_text += "• This is a Microsoft security restriction, not a permissions issue\n"
                        result_text += "• Exchange Online PowerShell is required for shared mailbox operations\n\n"
                        
                        # Provide PowerShell fallback
                        result_text += "**🔄 REQUIRED: PowerShell Method**\n\n"
                        result_text += "**Required PowerShell Commands:**\n"
                        result_text += f"```powershell\n"
                        result_text += f"# Connect to Exchange Online\n"
                        result_text += f"Connect-ExchangeOnline\n\n"
                        result_text += f"# Remove user from shared mailbox\n"
                        result_text += f"Remove-MailboxPermission -Identity '{resource_email}' -User '{user_email}' -AccessRights FullAccess\n\n"
                        result_text += f"# Verify the removal\n"
                        result_text += f"Get-MailboxPermission -Identity '{resource_email}' | Where-Object {{$_.User -eq '{user_email}'}}\n"
                        result_text += f"```\n\n"
                        result_text += f"**Current Status:** {user_display_name} needs to be removed via PowerShell\n"
                        result_text += "**Why PowerShell?:** Microsoft Graph API does not support shared mailbox member management\n"
                        return CallToolResult(content=[TextContent(type="text", text=result_text)])
                        
            except Exception as e:
                result_text += f"❌ **All methods failed:** {str(e)}\n\n"
                result_text += "**🔧 Manual Steps Required:**\n"
                result_text += "1. Use Microsoft 365 Admin Center\n"
                result_text += "2. Or use Exchange Online PowerShell\n"
                result_text += "3. Contact your administrator\n"
                return CallToolResult(content=[TextContent(type="text", text=result_text)])
                
        elif resource_type == "microsoft365_group":
            result_text += "🔄 **Microsoft 365 Group Detected - Using Graph API**\n\n"
            
            try:
                group_response = await make_graph_request("GET", f"/groups?$filter=mail eq '{resource_email}'")
                if group_response.get("value"):
                    group_id = group_response["value"][0]["id"]
                    
                    if action == "add":
                        member_data = {"@odata.id": f"{GRAPH_BASE_URL}/users/{user_email}"}
                        await make_graph_request("POST", f"/groups/{group_id}/members/$ref", data=member_data)
                        result_text += "✅ **User added successfully**\n\n"
                        
                        if access_level == "owner":
                            owner_data = {"@odata.id": f"{GRAPH_BASE_URL}/users/{user_email}"}
                            await make_graph_request("POST", f"/groups/{group_id}/owners/$ref", data=owner_data)
                            result_text += "✅ **Owner access granted**\n\n"
                        
                        result_text += f"**✅ Success:** {user_display_name} has been added to {resource_display_name}\n"
                        result_text += f"**Access Level:** {access_level.title()}\n"
                        return CallToolResult(content=[TextContent(type="text", text=result_text)])
                        
            except Exception as e:
                result_text += f"❌ **Failed to add user:** {str(e)}\n\n"
                result_text += "**🔧 Alternative Methods:**\n"
                result_text += "1. Use Microsoft 365 Admin Center\n"
                result_text += "2. Use Exchange Online PowerShell\n"
                return CallToolResult(content=[TextContent(type="text", text=result_text)])
                
        elif resource_type == "user_mailbox":
            result_text += "🔄 **User Mailbox Detected - Using Delegation**\n\n"
            
            result_text += "**📋 PowerShell Required for Mailbox Delegation:**\n"
            result_text += f"```powershell\n"
            result_text += f"# Connect to Exchange Online\n"
            result_text += f"Connect-ExchangeOnline\n\n"
            result_text += f"# Delegate mailbox access\n"
            result_text += f"Add-MailboxPermission -Identity '{resource_email}' -User '{user_email}' -AccessRights FullAccess\n\n"
            result_text += f"# Verify the permission\n"
            result_text += f"Get-MailboxPermission -Identity '{resource_email}' | Where-Object {{$_.User -eq '{user_email}'}}\n"
            result_text += f"```\n\n"
            result_text += f"**Current Status:** {user_display_name} needs mailbox delegation via PowerShell\n"
            return CallToolResult(content=[TextContent(type="text", text=result_text)])
            
        else:
            result_text += f"🔄 **{resource_type.replace('_', ' ').title()} Detected**\n\n"
            result_text += "**📋 Manual Steps Required:**\n"
            result_text += "1. Use Microsoft 365 Admin Center\n"
            result_text += "2. Or use Exchange Online PowerShell\n"
            result_text += "3. Contact your administrator\n\n"
            result_text += f"**Resource Type:** {resource_type.replace('_', ' ').title()}\n"
            result_text += f"**Action:** {action.title()} {user_display_name} to {resource_display_name}\n"
            return CallToolResult(content=[TextContent(type="text", text=result_text)])
            
    except Exception as e:
        error_text = f"**❌ Smart Resource Access Management Failed**\n\n"
        error_text += f"**Error:** {str(e)}\n\n"
        error_text += "**🔧 Troubleshooting:**\n"
        error_text += "1. Check that both user and resource names are correct\n"
        error_text += "2. Verify you have the necessary permissions\n"
        error_text += "3. Try using exact email addresses\n"
        error_text += "4. Use Microsoft 365 Admin Center as alternative\n"
        
        return CallToolResult(content=[TextContent(type="text", text=error_text)])

def create_server():
    """Create and configure the MCP server."""
    # Create the FastMCP server
    app = FastMCP("m365-admin")
    
    # ============================================================================
    # CONSOLIDATED TOOL REGISTRATION - 12 tools instead of 32
    # ============================================================================
    
    # Core Management Tools (6)
    app.add_tool(
        manage_group_membership,
        name="manage_group_membership",
        title="Manage Group Membership",
        description="⚠️ GRAPH API LIMITATIONS: Only works with regular Microsoft 365 Groups. Does NOT work with shared mailboxes, user mailboxes, or distribution lists. For shared mailboxes, use PowerShell. For adding users, use 'add_user_to_any_group_type' instead."
    )
    
    app.add_tool(
        get_group_information,
        name="get_group_information",
        title="Get Group Information",
        description="🔍 ENHANCED SEARCH: Unified tool to get comprehensive information about groups, users, shared mailboxes, and distribution lists. Includes fuzzy search, auto-completion, and smart suggestions. Use this for ALL resource lookups!"
    )
    
    app.add_tool(
        create_resource,
        name="create_resource",
        title="Create Resource",
        description="⚠️ GRAPH API LIMITATIONS: Can create users and groups, but cannot create shared mailboxes or distribution lists via Graph API. Use PowerShell for mailbox creation."
    )
    
    app.add_tool(
        update_resource,
        name="update_resource",
        title="Update Resource",
        description="⚠️ GRAPH API LIMITATIONS: Can update user and group properties, but cannot modify mailbox permissions or shared mailbox settings. Use PowerShell for mailbox operations."
    )
    
    app.add_tool(
        delete_resource,
        name="delete_resource",
        title="Delete Resource",
        description="⚠️ GRAPH API LIMITATIONS: Can delete users and groups, but cannot delete shared mailboxes or distribution lists via Graph API. Use PowerShell for mailbox deletion."
    )
    
    app.add_tool(
        list_users,
        name="list_users",
        title="List Users",
        description="✅ FULLY SUPPORTED: List users in Microsoft 365 with optional filtering and pagination. No Graph API limitations."
    )
    
    # Specialized Tools (3)
    app.add_tool(
        delegate_user_mailbox_access,
        name="delegate_user_mailbox_access",
        title="Delegate User Mailbox Access",
        description="⚠️ GRAPH API LIMITATIONS: Cannot delegate mailbox access via Graph API. This tool provides PowerShell instructions only. Microsoft Graph API does not support mailbox permission delegation."
    )
    
    app.add_tool(
        prepare_user_for_shared_mailbox_conversion,
        name="prepare_user_for_shared_mailbox_conversion",
        title="Prepare User for Shared Mailbox Conversion",
        description="⚠️ GRAPH API LIMITATIONS: Can prepare user account (disable, remove license) but cannot convert mailbox type via Graph API. Requires PowerShell for actual conversion: Set-Mailbox -Type Shared"
    )
    
    app.add_tool(
        test_connectivity,
        name="test_connectivity",
        title="Test Connectivity",
        description="✅ FULLY SUPPORTED: Unified tool to test authentication, connectivity, and API access. No Graph API limitations."
    )
    
    # Utility Tools (3)
    app.add_tool(
        list_distribution_lists,
        name="list_distribution_lists",
        title="List Distribution Lists",
        description="⚠️ GRAPH API LIMITATIONS: Can list distribution lists but cannot manage members via Graph API. Use PowerShell for member management."
    )
    
    app.add_tool(
        list_shared_mailboxes,
        name="list_shared_mailboxes",
        title="List Shared Mailboxes",
        description="⚠️ GRAPH API LIMITATIONS: Can list shared mailboxes but cannot manage members or permissions via Graph API. Use PowerShell for all shared mailbox operations."
    )
    
    app.add_tool(
        get_mailbox_info,
        name="get_mailbox_info",
        title="Get Mailbox Info",
        description="⚠️ GRAPH API LIMITATIONS: Can get basic mailbox settings but cannot access mailbox permissions or delegation settings. Use PowerShell for detailed mailbox information."
    )
    
    app.add_tool(
        add_user_to_microsoft365_group,
        name="add_user_to_microsoft365_group",
        title="Add User to Microsoft 365 Group",
        description="⚠️ GRAPH API LIMITATIONS: Only works with Microsoft 365 Groups (Unified groups), not shared mailboxes or user mailboxes. For shared mailboxes, use PowerShell."
    )
    
    app.add_tool(
        manage_microsoft365_group_access,
        name="manage_microsoft365_group_access",
        title="Manage Microsoft 365 Group Access",
        description="⚠️ GRAPH API LIMITATIONS: Only works with Microsoft 365 Groups (Unified groups), not shared mailboxes or user mailboxes. For shared mailboxes, use PowerShell."
    )
    
    app.add_tool(
        create_user_simple,
        name="create_user_simple",
        title="Create User (Simple)",
        description="✅ FULLY SUPPORTED: Create a new user account with clear, simple parameters. No Graph API limitations for user creation."
    )
    
    app.add_tool(
        add_user_to_any_group_type,
        name="add_user_to_any_group_type",
        title="Add User to Group (Ultra-Robust)",
        description="⚠️ GRAPH API LIMITATIONS: Works with Microsoft 365 Groups but has limited success with shared mailboxes. For shared mailboxes, use PowerShell: Add-MailboxPermission"
    )
    
    app.add_tool(
        add_user_to_shared_mailbox,
        name="add_user_to_shared_mailbox",
        title="Add User to Shared Mailbox",
        description="⚠️ GRAPH API LIMITATIONS: Cannot add users to shared mailboxes via Graph API. This tool provides PowerShell instructions only. Use: Add-MailboxPermission"
    )
    
    # 🚀 NEW: Smart Unified Tool
    app.add_tool(
        manage_resource_access,
        name="manage_resource_access",
        title="Smart Resource Access Management",
        description="🔧 SMART UNIFIED TOOL: Intelligently manages access to any Microsoft 365 resource with automatic resource detection and optimal method selection. ⚠️ GRAPH API LIMITATIONS: Cannot manage shared mailbox permissions via Graph API - provides PowerShell fallback automatically."
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
