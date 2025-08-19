# M365 MCP Server

A Model Context Protocol (MCP) server for Microsoft 365 administration using the Microsoft Graph API.

## 🚀 Features

- **User Management**: Create, read, update, delete users and their mailboxes
- **Microsoft 365 Groups**: Full CRUD operations for Microsoft 365 Groups, including member and owner management
- **Shared Mailboxes**: List and manage shared mailboxes (read-only via Graph API, PowerShell required for member management)
- **Mailbox Operations**: Delegate mailboxes and convert user mailboxes to shared (PowerShell instructions provided)
- **Enhanced Error Handling**: Comprehensive error reporting with troubleshooting guidance and Graph API limitation explanations

## 🛠️ Setup

### 1. Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) → "Azure Active Directory" → "App registrations"
2. Click "New registration"
3. Name your application
4. Supported account types: "Accounts in this organizational directory only"
5. Click "Register"

### 2. Configure Permissions

1. Go to "API permissions" in your app registration
2. Click "Add a permission" → "Microsoft Graph" → "Application permissions"
3. Add these **required** permissions:

   **User Management:**
   - `User.ReadWrite.All` - Required for all user operations (create, read, update, delete users)
   - `Directory.ReadWrite.All` - Required for user directory operations and license management

   **Group Management:**
   - `Group.ReadWrite.All` - Required for Microsoft 365 Groups operations (create, manage members/owners)

   **Mailbox Operations:**
   - `MailboxSettings.ReadWrite.All` - Required for mailbox settings and calendar permissions

   **License Management:**
   - `User.ReadWrite.All` (already listed above) - Covers license assignment operations

4. Click "Grant admin consent for [Your Organization]"

### 3. Important Notes About Permissions

**⚠️ Graph API Limitations:**
- **Shared Mailboxes**: Cannot add/remove members via Graph API - requires Exchange Online PowerShell
- **Mailbox Delegation**: Cannot delegate mailbox access via Graph API - requires Exchange Online PowerShell  
- **Distribution Lists**: Cannot manage members via Graph API - requires Exchange Online PowerShell
- **Mailbox Type Conversion**: Cannot convert mailbox types via Graph API - requires Exchange Online PowerShell

**✅ What Works with Graph API:**
- User creation, updates, deletion
- Microsoft 365 Groups management
- Basic mailbox settings retrieval
- License assignment and removal
- User account enable/disable

### 4. Create Client Secret

1. Go to "Certificates & secrets" in your app registration
2. Click "New client secret"
3. Description: `MCP Server Secret`
4. Expiration: Choose appropriate duration
5. Copy the **Value** (not the ID) - this is your `CLIENT_SECRET`

### 5. Get Credentials

Copy these values from your app registration:
- **Application (client) ID** → `CLIENT_ID`
- **Directory (tenant) ID** → `TENANT_ID`
- **Client Secret Value** → `CLIENT_SECRET`

## 🚀 Installation

### 1. Install Dependencies

```bash
pip3 install -r requirements.txt
```

### 2. Configure Environment

Create a `.env` file:
```env
TENANT_ID=your_tenant_id_here
CLIENT_ID=your_client_id_here
CLIENT_SECRET=your_client_secret_here
```

### 3. Run the Server

```bash
python3 m365_mcp_server.py
```

## 🖥️ Claude Desktop Integration

Add this to your Claude Desktop configuration:

**macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
**Linux:** `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "m365-admin": {
      "command": "python3",
      "args": ["/path/to/your/m365-mcp/m365_mcp_server.py"],
      "env": {
        "TENANT_ID": "your_tenant_id_here",
        "CLIENT_ID": "your_client_id_here",
        "CLIENT_SECRET": "your_client_secret_here"
      }
    }
  }
}
```

## 🔍 Troubleshooting

### PowerShell Requirements

**Many operations require Exchange Online PowerShell:**
- Shared mailbox member management
- Mailbox delegation
- Distribution list member management
- Mailbox type conversion

**To install Exchange Online PowerShell:**
```powershell
Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline
```

### Common Error Codes

- **401 Unauthorized**: Token expired or insufficient permissions
- **403 Forbidden**: Missing API permissions or admin consent not granted
- **404 Not Found**: Resource doesn't exist or email address is incorrect
- **400 Bad Request**: Invalid request format or missing required properties
- **409 Conflict**: Resource already exists

### Authentication Issues

1. **Check Environment Variables**: Ensure all three variables are set
2. **Verify App Registration**: Confirm app exists and has correct permissions
3. **Grant Admin Consent**: Ensure admin consent is granted for all permissions
4. **Check Client Secret**: Verify secret is valid and not expired

### API Permission Issues

1. **Review Permissions**: Ensure all required permissions are added:
   - `User.ReadWrite.All`
   - `Directory.ReadWrite.All` 
   - `Group.ReadWrite.All`
   - `MailboxSettings.ReadWrite.All`
2. **Check Permission Type**: Use "Application permissions" not "Delegated permissions"
3. **Grant Admin Consent**: Click "Grant admin consent" in Azure Portal
4. **Wait for Propagation**: Changes may take up to 15 minutes to propagate

### Graph API Limitations

**Expected Behaviors (Not Errors):**
- **Shared mailbox member management fails**: This is expected - use PowerShell `Add-MailboxPermission`
- **Mailbox delegation fails**: This is expected - use PowerShell `Add-MailboxPermission`
- **Distribution list member management fails**: This is expected - use PowerShell `Add-DistributionGroupMember`
- **Mailbox type conversion fails**: This is expected - use PowerShell `Set-Mailbox -Type Shared`

### Common Issues

- **"pip command not found"**: Use `pip3` instead of `pip`
- **"Permission denied" (403)**: Ensure admin consent is granted for all permissions
- **"Resource not found" (404)**: Check if the user/group exists and verify email addresses
- **"Unsupported query" (400)**: Usually indicates Graph API limitation - use PowerShell instead
- **"Authorization denied" (403)**: Check if the operation is supported by Graph API
- **Microsoft 365 Group membership restrictions**: Some groups may have restricted membership management
- **Shared mailbox operations**: Most shared mailbox operations require Exchange Online PowerShell
