# QuickBooks Desktop Enterprise to Google Sheets Integration

## **Project Overview**

This project enables automatic synchronization of QuickBooks Desktop Enterprise account balances to Google Sheets.

## **Architecture**

```
QuickBooks Desktop Enterprise → Rust Service → Google Apps Script → Google Sheets
```

## **Components**

### 1. **Google Apps Script** (`Google_Sheet_Functions/src/Code.ts`)
- Receives data from Rust service via HTTP POST
- Updates Google Sheets cells with account balances
- Handles authentication and error logging
- Deployed as a Web App to receive external requests

### 2. **Rust Service** (`quickbooks-rust-service/`)
- **Final program name: qb_sync.exe**
- Connects to QuickBooks Desktop Enterprise using COM interface
- Extracts account balance for specified account (9445)
- Sends data to Google Apps Script Web App
- Runs on configurable schedule (default: hourly)

## **Setup Instructions**

### **Phase 1: Google Apps Script Setup**

1. **Navigate to Google_Sheet_Functions
   - Open Code.ts

2. **Access the updated code**:
   - Open Google Apps Script
   - Copy the Code.ts file via Editor | Code.gs
   - You should now see functions like `setupPermissions`, `checkPermissions`, etc.

3. **Set up permissions**:
   - In the editor, select `setupPermissions` from the function dropdown
   - Click the ▶️ "Run" button
   - When prompted, click "Review permissions" → "Allow"

4. **Generate API Key**:
   - Select `setupQuickBooksIntegration` from the dropdown
   - Click the ▶️ "Run" button
   - Copy the generated API key from the execution log
   - Paste this into the proper place in the config.toml file in the config directory of quickbooks-rust-service

5. **Deploy as Web App**:
   - Click "Deploy" > "New Deployment"
   - Choose "Web app" type
   - Set "Execute as: Me" and "Who has access: Anyone"
   - Click "Deploy" and copy the Web App URL to the proper place in the config.toml file in the config directory of quickbooks-rust-service

6. **Test the setup**:
   - Select `testWebAppEndpoint` from the dropdown
   - Click the ▶️ "Run" button
   - Check that it shows success in the execution log

7. **TODO Add instuctions for wiring up to Google Cloud API Logging & OAuth

### **Phase 2: Rust Service Setup**

1. **Navigate to quickbooks-rust-service**:

2. **Configure the service**:
   Edit `config/config.toml`:
   ```toml
   [quickbooks]
   company_file = "C:\\Path\\To\\Your\\Company.qbw"
   
   [google_sheets]
   webapp_url = "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec"
   api_key = "YOUR_API_KEY_FROM_STEP_1"
   ```

3. **Make sync_blocks**
   - For each account that you want to sync with a sheet make a sync block
   Edit `config/config.toml`:
   ```toml
   [[sync_blocks]]
   spreadsheet_id = "THE SPREADHSEET ID PART OF THE URL TO YOUR SPREADSHEET"
   account_full_name = "THE FULL NAME OF THE ACCOUNT IN QUICKBOOKS(*)"
   sheet_name = "THE EXACT NAME OF THE TAB WHERE THE DATA WILL GO"
   cell_address = "THE CELL ON THAT SHEET WHERE THE DATA WILL GO"
   ```

(*) The full name of the account in Quickbooks can usually be found by looking at the balance sheet. The name will not include account numbers and each level of the account's heirarchy will be separated by colons (:). For example:

"Cash Accounts:BoA Accounts:PRINT RESERVE Vault"

4. **Run the service as a test**:
   - In the quickbooks-rust-service directory:
   ```bash
   cargo run -- --verbose
   ```

## 🔧 **Requirements**

### **Google Apps Script Side**
- Google Apps Script deployed as Web App
- API key configured for authentication
- Target Google Sheets document accessible

### **Rust Service Side (Windows)**
- QuickBooks Desktop Enterprise installed and running
- Company file accessible at the indicated location in config.toml
- Rust toolchain installed (for building on the Windows machine; not necessary to run the compiled program)
- Access to the internet

## **Data Flow**

1. **Rust service starts** and connects to QuickBooks
2. The service:
   - Queries QuickBooks for the list of all accounts & balances
   - Sends HTTP POST to Google Apps Script Web App for each sync_block
   - Google Apps Script updates the specified cell

## **Monitoring**

- **Google Apps Script logs** show received data and update status
- Recommend enable Google Cloud API logging

## **Scheduling**

- Use Windows Task Scheduler to schedule the program to run every hour
- You may have to give the account that you want to use special permissions
-- Run secpol.msc /s as Administrator on the Windows machine
--- Navigate to Local Policies | User Rights Assignment | Log on as a batch job and add the user that runs the task

## **Troubleshooting**

### **Common Issues:**

1. **"Invalid API key"**: 
   - Open Google Apps Script editor
   - Select `setupQuickBooksIntegration` from dropdown, click ▶️ "Run"
   - Copy the generated API key to your config.toml

2. **"Sheet not found"**:
   - Sheets means Tabs
   - Update your config.toml with the correct Tab name

3. **"Permission denied"**:
   - Select `setupPermissions` from dropdown, click ▶️ "Run"
   - Click "Review permissions" → "Allow" when prompted

4. **"Failed to create QuickBooks session manager"**:
   - Install QuickBooks Desktop SDK

5. **"STATUS_ACCESS_VIOLATION"**:
   - Run as Administrator
   - Install QuickBooks SDK properly
  
6. **"Page Not Found"**:
   - Deploy Google Apps Script as Web App
   - Use the correct Web App URL (ends with `/exec`)
   - Set permissions to "Anyone"

## 📈 **Current Status**

✅ **Google Apps Script**: Deployed and ready  
✅ **Rust Service**: Built and configured  

## **Expected Result**

Once the Google Functions are created, depolyed and authorized, and the Rust program is installed along with its config directory and config.toml file, your Google Sheets will automatically update with the current balance for the accounts you have indicated, in the sheet/cells you have indicated every hour