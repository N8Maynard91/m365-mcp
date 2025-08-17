# M365 MCP Server

A Model Context Protocol (MCP) server for Microsoft 365 administration using the Microsoft Graph API.

## 🚀 Features

- **User Management**: Create, read, update, delete users and their mailboxes
- **Distribution Lists**: Full CRUD operations for distribution lists, including member management
- **Microsoft 365 Groups**: Add users to Microsoft 365 Groups (which function as both shared mailboxes and distribution lists)
- **Shared Mailboxes**: Create and manage shared mailboxes, including member management
- **Mailbox Operations**: Delegate mailboxes and convert user mailboxes to shared
- **Member Management**: List members of distribution lists and shared mailboxes
- **Enhanced Error Handling**: Comprehensive error reporting with troubleshooting guidance

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
3. Add these permissions:
   - `User.ReadWrite.All` - Required for all user operations
   - `Group.ReadWrite.All` - Required for distribution lists and Microsoft 365 Groups
   - `MailboxSettings.ReadWrite` - Required for mailbox operations
   - `Directory.ReadWrite.All` - Required for creating and managing users
   - `User.ManageIdentities.All` - Required for user identity management
   - .All permissions do not need to be  used, but reducing these scopes may limit mcp server capbilites
4. Click "Grant admin consent for [Your Organization]"

### 3. Create Client Secret

1. Go to "Certificates & secrets" in your app registration
2. Click "New client secret"
3. Description: `MCP Server Secret`
4. Expiration: Choose appropriate duration
5. Copy the **Value** (not the ID) - this is your `CLIENT_SECRET`

### 4. Get Credentials

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

1. **Review Permissions**: Ensure all required permissions are added
2. **Check Permission Type**: Use "Application permissions" not "Delegated permissions"
3. **Grant Admin Consent**: Click "Grant admin consent" in Azure Portal
4. **Wait for Propagation**: Changes may take up to 15 minutes to propagate

### Common Issues

- **"pip command not found"**: Use `pip3` instead of `pip`
- **"Permission denied"**: Ensure admin consent is granted for all permissions
- **"Resource not found" (404)**: Check if the user/group exists and verify email addresses
- **Microsoft 365 Group membership restrictions**: Some groups may have restricted membership management that prevents API-based member addition
- **Microsoft 365 Groups vs Distribution Lists**: The server automatically detects group types and uses the appropriate API endpoints for each
