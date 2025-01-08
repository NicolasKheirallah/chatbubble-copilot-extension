# Setup Guide

This document provides a step-by-step guide for setting up the chatbot solution in your SharePoint environment. It covers everything from prerequisites to deployment and troubleshooting to ensure a seamless setup experience.

---

## **Prerequisites**

Before proceeding with the setup, ensure you have the following:

1. **Tenant App Catalog**:
   - A Tenant App Catalog site in your SharePoint Online environment.

2. **Permissions**:
   - Permissions to create lists and manage App Catalog configurations.

3. **Azure AD App Registration**:
   - An App Registration with appropriate API permissions for MSAL authentication.

4. **Development Tools**:
   - A local development environment with Node.js, NPM, Gulp, and Visual Studio Code installed.

---

## **Steps**

### **1. Deploying the Application Customizer**

1. Open the project in your development environment.
2. Use the following command to bundle your solution:
   ```bash
   gulp bundle --ship
   ```
3. Package your solution using:
   ```bash
   gulp package-solution --ship
   ```
4. Upload the generated `.sppkg` file from the `sharepoint/solution` folder to your Tenant App Catalog.
5. Deploy the package and ensure the tenant-wide deployment option is enabled.

---

### **2. Grant API Permissions**

1. Navigate to the Azure Portal and locate your App Registration.
2. Under **API Permissions**, ensure the following permissions are granted:
   - `User.Read`
   - Any custom scopes required for accessing your chatbot backend.
3. Grant admin consent for the permissions.

---

### **3. Configure the Extension**

After deploying the extension, configure its settings by creating the following resources.

#### **A. Configuration List**

Create a SharePoint list to store your chatbot configurations.

1. Navigate to the **Tenant App Catalog Site**:
   - The URL typically follows the pattern: `https://<tenant>-admin.sharepoint.com/sites/AppCatalog`.
   - Replace `<tenant>` with your actual tenant name.

2. Create a New SharePoint List:
   - **Name:** `TenantWideExtensionsConfig`
   - **Columns:**
     - **Title** (Default): Use for identification or leave it empty.
     - **BotURL** (Single line of text): The endpoint URL of your PVA chatbot.
     - **BotName** (Single line of text): The display name of your chatbot.
     - **ButtonLabel** (Single line of text): Label for the chat toggle button (e.g., "Chat with us").
     - **BotAvatarImage** (Hyperlink or Picture): URL to the chatbot's avatar image.
     - **BotAvatarInitials** (Single line of text): Initials for the chatbot's avatar if no image is provided.
     - **Greet** (Yes/No): Determines if the chatbot should greet the user upon opening.
     - **CustomScope** (Single line of text): The scope for MSAL authentication.
     - **ClientID** (Single line of text): Azure AD App Registration's Client ID.
     - **Authority** (Single line of text): Azure AD Authority URL (e.g., `https://login.microsoftonline.com/<tenant-id>`).

3. Add a Configuration Item:
   - Click **New** to add a new item.
   - Fill in all the fields with appropriate values.
   - Save the item.

   > **Note:** Ensure that only **one item** exists in this list, as the Application Customizer fetches the **first item** for configuration.

#### **B. JSON Backup Configuration**

If the configuration list is not available, the solution will fallback to a hardcoded JSON configuration. Update the JSON file with the following fields:

```json
{
  "BotURL": "<PVA chatbot URL>",
  "BotName": "<Chatbot Display Name>",
  "ButtonLabel": "<Chat Toggle Button Label>",
  "BotAvatarImage": "<URL to Avatar Image>",
  "BotAvatarInitials": "<Avatar Initials>",
  "Greet": true,
  "CustomScope": "<Custom Scope for MSAL>",
  "ClientID": "<Azure AD Client ID>",
  "Authority": "<Azure AD Authority URL>"
}
```

> **Note:** This JSON serves as a fallback mechanism and should match the structure of the SharePoint list.

---

### **4. Configuration Setup**

Before diving into the code, set up a SharePoint list to store your chatbot configurations.

1. Create the Configuration List

Navigate to the Tenant App Catalog Site:

Typically, the App Catalog site URL follows the pattern: `https://<tenant>-admin.sharepoint.com/sites/AppCatalog`. Replace `<tenant>` with your actual tenant name.

Create a New SharePoint List:

- **Name:** `TenantWideExtensionsConfig`
- **Columns:**
  - **Title (Default):** Use for identification or leave it empty.
  - **BotURL (Single line of text):** The endpoint URL of your PVA chatbot.
  - **BotName (Single line of text):** The display name of your chatbot.
  - **ButtonLabel (Single line of text):** Label for the chat toggle button (e.g., "Chat with us").
  - **BotAvatarImage (Hyperlink or Picture):** URL to the chatbot's avatar image.
  - **BotAvatarInitials (Single line of text):** Initials for the chatbot's avatar if no image is provided.
  - **Greet (Yes/No):** Determines if the chatbot should greet the user upon opening.
  - **CustomScope (Single line of text):** The scope for MSAL authentication.
  - **ClientID (Single line of text):** Azure AD App Registration's Client ID.
  - **Authority (Single line of text):** Azure AD Authority URL (e.g., `https://login.microsoftonline.com/<tenant-id>`).

Add a Configuration Item:

- Click **New** to add a new item.
- Fill in all the fields with appropriate values.
- Save the item.

   > **Note:** Ensure that only **one item** exists in this list, as the Application Customizer fetches the **first item** for configuration.

---

### **5. Verify the Setup**

1. Open a SharePoint site where the extension is deployed.
2. Verify that the chatbot toggle button appears as expected.
3. Click the button to open the chatbot window and validate functionality.

---

### **6. Troubleshooting**

#### **Common Issues**

1. **Chatbot not loading:**
   - Ensure the configuration list is properly set up.
   - Verify the JSON fallback configuration if the list is unavailable.

2. **Authentication errors:**
   - Confirm that the Azure AD App Registration permissions are correctly configured.

3. **UI Issues:**
   - Verify that the SCSS styles are properly compiled and included in the project.

#### **Debugging Tips**

- Use browser developer tools to inspect network requests and console logs.
- Check Azure Portal logs for failed authentication attempts.
- Ensure the SPFx web part is correctly configured in the elements.xml file.

---

This completes the setup process for your chatbot solution. If you encounter issues, refer to the troubleshooting section or reach out to the support team for assistance.

