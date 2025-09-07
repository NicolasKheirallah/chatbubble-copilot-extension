# ChatBubble Copilot Extension Troubleshooting Guide

## Common Issues and Solutions

### 1. DirectLine Connection Issues

#### Symptoms:
- "DirectLine token not available" error
- 502 Bad Gateway errors when sending messages
- Stuck at "Loading chatbot..." screen

#### Root Causes & Solutions:

**A. Bot URL Configuration**
```javascript
// Check your configuration in SharePoint list or ConfigurationService.ts
BotURL: "https://your-bot.azurewebsites.net/api/messages"
```
- Ensure the bot URL is accessible and responding
- Verify the bot is properly deployed and running
- Test the bot URL directly in browser (should return a 405 Method Not Allowed for GET requests, which is normal)

**B. Azure AD App Registration Issues**
1. **Missing API Permissions:**
   ```
   Required Permissions:
   - Microsoft Graph: User.Read
   - Your Custom API: user_impersonation (or your custom scope)
   ```
   
2. **Incorrect Custom Scope:**
   ```
   Format: api://YOUR-APP-ID/SPO.Read
   Example: api://12345678-1234-1234-1234-123456789abc/user_impersonation
   ```

3. **Redirect URI Issues:**
   - Add your SharePoint site URL: `https://yourtenant.sharepoint.com/sites/yoursite`
   - Include both with and without trailing slash

**C. Token Exchange Configuration**
- In Azure AD > Authentication > Advanced settings
- Token Exchange URL: `https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token`
- Ensure the tenant ID is correct

**D. Bot Framework Configuration**
- Verify the bot has DirectLine channel enabled
- Check if the bot requires authentication
- Ensure the bot can handle SSO token exchange

### 2. Session Management Issues

#### Symptoms:
- Blank chat when clicking "New Chat" button
- Unable to switch between previous sessions
- Sessions not being saved

#### Solutions:
✅ **Fixed in v3.0.0:**
- Added proper DirectLine connection cleanup on session switch
- Implemented WebChat reinitialization using key prop
- Enhanced session persistence with proper state management

### 3. Header Visibility Issues

#### Symptoms:
- White text on light background making header invisible
- Unable to see close button or other header elements

#### Solutions:
✅ **Fixed in v3.0.0:**
- Added comprehensive CSS overrides with `!important` declarations
- Implemented solid blue background (#0078d4) instead of CSS variables
- Enhanced button visibility with proper contrast

### 4. Build and Deployment Issues

#### Symptoms:
- TypeScript compilation errors
- ESLint warnings about deprecated features
- Build failures

#### Solutions:
✅ **Fixed in v3.0.0:**
- Updated to use `disableFileUpload` instead of deprecated `hideUploadButton`
- Fixed all TypeScript strict mode issues
- Eliminated all build warnings and errors

## Diagnostic Steps

### Step 1: Check Browser Console
```javascript
// Open browser developer tools and check for these errors:
1. "DirectLine token not available"
2. 502 Bad Gateway errors
3. MSAL authentication errors
4. CORS issues
```

### Step 2: Verify Configuration
```javascript
// Run this in browser console to check configuration:
console.log('Bot URL:', window.TenantWideConfig?.BotURL);
console.log('Client ID:', window.TenantWideConfig?.ClientID);
console.log('Custom Scope:', window.TenantWideConfig?.CustomScope);
```

### Step 3: Test Bot Endpoint
```bash
# Test bot endpoint directly:
curl -X POST https://your-bot.azurewebsites.net/api/messages \
  -H "Content-Type: application/json" \
  -d '{"type":"message","text":"test"}'
```

### Step 4: Check Azure AD Logs
1. Go to Azure AD > Sign-in logs
2. Look for authentication requests from your application
3. Check for any failed authentications or missing permissions

### Step 5: Verify App Catalog Deployment
1. Go to SharePoint Admin Center
2. Check Apps > App Catalog
3. Verify the app is deployed and API permissions are approved

## Configuration Checklist

### Azure AD App Registration:
- [ ] App registration created with correct name
- [ ] Redirect URIs include SharePoint site URL
- [ ] API permissions granted and admin consent provided
- [ ] Custom scope properly defined (if using custom API)
- [ ] Token exchange URL configured

### Bot Configuration:
- [ ] Bot deployed and accessible
- [ ] DirectLine channel enabled
- [ ] SSO configuration set up (if required)
- [ ] Bot can handle token exchange requests

### SharePoint Configuration:
- [ ] App deployed to tenant app catalog
- [ ] API permissions approved in SharePoint admin center
- [ ] Configuration list created with correct values
- [ ] Extension activated on target sites

### Network and Security:
- [ ] Bot endpoint accessible from SharePoint Online
- [ ] No firewall blocking DirectLine communications
- [ ] HTTPS properly configured
- [ ] CORS policies allow SharePoint domain

## Support Resources

- [Bot Framework Documentation](https://docs.microsoft.com/en-us/azure/bot-service/)
- [SharePoint Framework Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/)
- [Azure AD Authentication](https://docs.microsoft.com/en-us/azure/active-directory/develop/)
- [DirectLine API Reference](https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-direct-line-3-0-concepts)

## Getting Help

If you continue to experience issues:

1. **Check the GitHub Issues**: https://github.com/microsoft/BotFramework-WebChat/issues
2. **SharePoint Developer Community**: https://techcommunity.microsoft.com/t5/sharepoint-developer/bd-p/SharePointDev
3. **Stack Overflow**: Use tags `botframework`, `sharepoint-framework`, `azure-ad`

## Known Limitations

- **SharePoint Server**: This solution only works with SharePoint Online, not on-premises
- **Internet Explorer**: Not supported (deprecated browser)
- **Mobile Apps**: Limited support in SharePoint mobile apps
- **Network Requirements**: Requires internet access for DirectLine communication