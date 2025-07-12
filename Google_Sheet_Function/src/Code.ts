"use strict";
/**
 * Google Sheets integration for QuickBooks data
 *
 * This script receives QuickBooks account data from an external Windows service (via HTTP POST),
 * and updates the specified cell in a Google Sheet. Account selection is handled by the Windows
 * service (via command-line argument), and scheduling is managed by Windows Task Scheduler.
 *
 * No internal scheduling or config-driven account selection is used in this script.
 */
/**
 * Updates a specific cell with QuickBooks account data
 * @param {string} accountNumber - The QuickBooks account number (provided by the Windows service)
 * @param {number} accountValue - The account balance/value
 * @param {string} cellAddress - The cell address (e.g., "A1", "B2")
 * @param {string} spreadsheetId - The spreadsheet ID (optional, uses active if not provided)
 * @param {string} sheetName - The name of the sheet (optional)
 * @return {string} Success message
 * @customfunction
 *
 * Note: Account selection is handled by the Windows service, not this script.
 */
function UPDATE_QB_ACCOUNT(accountNumber, accountValue, cellAddress, spreadsheetId, sheetName) {
    try {
        // Use specific spreadsheet if ID provided, otherwise use active spreadsheet
        const spreadsheet = spreadsheetId ?
            SpreadsheetApp.openById(spreadsheetId) :
            SpreadsheetApp.getActiveSpreadsheet();
        const sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : spreadsheet.getActiveSheet();
        if (!sheet) {
            console.error(`[UPDATE_QB_ACCOUNT] Sheet not found: ${sheetName}`);
            throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${allSheets.join(', ')}`);
        }
        // Update the cell with the account value
        const range = sheet.getRange(cellAddress);
        range.setValue(accountValue);
        const msg = `Account ${accountNumber} updated: ${accountValue} at ${new Date().toLocaleString()}`;
        return msg;
    }
    catch (error) {
        console.error('[UPDATE_QB_ACCOUNT] Error:', error);
        // Re-throw the error instead of returning an error message
        throw error;
    }
}
/**
 * Web App endpoint to receive QuickBooks data from the Windows service
 * This function handles POST requests from the Rust service running as a Windows service.
 *
 * The Windows service is responsible for selecting the account and scheduling execution.
 */
function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);
        // Validate required fields
        if (!data.accountNumber || data.accountValue === undefined || !data.cellAddress) {
            console.error('[doPost] Missing required fields:', data);
            throw new Error('Missing required fields: accountNumber, accountValue, cellAddress');
        }
        // Validate API key for security
        const scriptApiKey = PropertiesService.getScriptProperties().getProperty('QB_API_KEY');
        if (!data.apiKey || data.apiKey !== scriptApiKey) {
            console.error('[doPost] Invalid API key:', data.apiKey);
            throw new Error('Invalid API key');
        }
        // Update the QuickBooks account data
        const result = UPDATE_QB_ACCOUNT(data.accountNumber, data.accountValue, data.cellAddress, data.spreadsheetId, // Pass spreadsheet ID if provided
            data.sheetName);
        return ContentService
            .createTextOutput(JSON.stringify({ success: true, message: result }))
            .setMimeType(ContentService.MimeType.JSON);
    }
    catch (error) {
        console.error('[doPost] Error:', error);
        return ContentService
            .createTextOutput(JSON.stringify({ success: false, error: error instanceof Error ? error.message : String(error) }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}
/**
 * Test function to verify Web App deployment
 *
 * This function simulates a POST request as if sent by the Windows service.
 */
function testWebAppEndpoint() {
    try {
        // Create a test POST request payload
        const testPayload = {
            accountNumber: 'TEST',
            accountValue: 999.99,
            spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
            cellAddress: 'A1',
            sheetName: 'Sheet1',
            apiKey: PropertiesService.getScriptProperties().getProperty('QB_API_KEY')
        };
        console.log('Test payload:', JSON.stringify(testPayload, null, 2));
        // Simulate the doPost function
        const mockEvent = {
            postData: {
                contents: JSON.stringify(testPayload)
            }
        };
        const response = doPost(mockEvent);
        console.log('Response:', response.getContent());
    }
    catch (error) {
        console.error('Test failed:', error);
    }
}
/**
 * Set up the API key for QuickBooks integration
 * Run this once to set up authentication and test the setup
 *
 * The API key is used to authenticate requests from the Windows service.
 */
function setupQuickBooksIntegration() {
    try {
        // Generate a unique API key
        const apiKey = Utilities.getUuid();
        // Store the API key in script properties
        PropertiesService.getScriptProperties().setProperty('QB_API_KEY', apiKey);
        // Log setup completion
        console.log('QuickBooks integration setup completed!');
        console.log('API Key:', apiKey);
        console.log('Copy this API key to your Rust service configuration.');
        // Test the UPDATE_QB_ACCOUNT function
        const testResult = UPDATE_QB_ACCOUNT('TEST', 1234.56, 'A1');
        console.log('Test result:', testResult);
        return apiKey;
    }
    catch (error) {
        console.error('Setup failed:', error);
        throw error;
    }
}
/**
 * Get the current API key (for setup purposes)
 * @return {string} The current API key
 */
function getQuickBooksApiKey() {
    return PropertiesService.getScriptProperties().getProperty('QB_API_KEY') || 'Not set';
}
/**
 * Manually trigger QuickBooks account update (for testing)
 * @param {number} accountValue - The account balance to set
 * @param {string} cellAddress - The cell address to update
 * @return {string} Result message
 * @customfunction
 *
 * Note: For production, updates are triggered by the Windows service, not manually.
 */
function TEST_QB_UPDATE(accountValue, cellAddress) {
    return UPDATE_QB_ACCOUNT('9445', accountValue, cellAddress);
}
/**
 * Setup function to configure permissions and test access
 * Run this function first to set up all necessary permissions
 *
 * No internal scheduling or cron logic is present; scheduling is handled externally.
 */
function setupPermissions() {
    try {
        console.log('🔧 Setting up permissions for QuickBooks integration...');
        // Test basic Google Sheets access
        console.log('📊 Testing Google Sheets access...');
        const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const activeSheet = activeSpreadsheet.getActiveSheet();
        console.log('✅ Active spreadsheet ID:', activeSpreadsheet.getId());
        console.log('✅ Active sheet name:', activeSheet.getName());
        // Test writing to a cell
        console.log('📝 Testing write access...');
        const testCell = activeSheet.getRange('A1');
        const originalValue = testCell.getValue();
        testCell.setValue('Permission Test - ' + new Date().toISOString());
        console.log('✅ Successfully wrote to cell A1');
        // Test reading from the cell
        const readValue = testCell.getValue();
        console.log('✅ Successfully read from cell A1:', readValue);
        // Restore original value if it wasn't empty
        if (originalValue !== '') {
            testCell.setValue(originalValue);
            console.log('✅ Restored original cell value');
        }
        // Test accessing script properties
        console.log('🔑 Testing script properties access...');
        const properties = PropertiesService.getScriptProperties();
        const testKey = 'PERMISSION_TEST';
        const testValue = 'SUCCESS_' + Date.now();
        properties.setProperty(testKey, testValue);
        const retrievedValue = properties.getProperty(testKey);
        if (retrievedValue === testValue) {
            console.log('✅ Script properties access working');
            properties.deleteProperty(testKey);
        }
        else {
            console.log('❌ Script properties access failed');
        }
        console.log('');
        console.log('🎉 Permission setup completed successfully!');
        console.log('📋 Next steps:');
        console.log('   1. Run setupQuickBooksIntegration() to generate API key');
        console.log('   2. Deploy as Web App with "Anyone" access');
        console.log('   3. Test with testWebAppEndpoint()');
    }
    catch (error) {
        console.error('❌ Permission setup failed:', error);
        console.log('');
        console.log('🔧 Troubleshooting:');
        console.log('   1. Make sure you have edit access to this spreadsheet');
        console.log('   2. Check that the Google Apps Script project has the necessary scopes');
        console.log('   3. Try running the script manually first');
        throw error;
    }
}
/**
 * Check current permissions and access levels
 */
function checkPermissions() {
    try {
        console.log('🔍 Checking current permissions...');
        // Check basic access
        const user = Session.getActiveUser();
        console.log('👤 Active user:', user.getEmail());
        // Check spreadsheet access
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        console.log('📊 Spreadsheet ID:', spreadsheet.getId());
        console.log('📊 Spreadsheet name:', spreadsheet.getName());
        console.log('📊 Spreadsheet URL:', spreadsheet.getUrl());
        // Check sheet access
        const sheets = spreadsheet.getSheets();
        console.log('📄 Available sheets:');
        sheets.forEach((sheet, index) => {
            console.log(`   ${index + 1}. ${sheet.getName()} (${sheet.getMaxRows()}x${sheet.getMaxColumns()})`);
        });
        // Check script properties
        const apiKey = PropertiesService.getScriptProperties().getProperty('QB_API_KEY');
        console.log('🔑 API Key configured:', apiKey ? 'YES' : 'NO');
        // Check if we can create a trigger (advanced permission)
        try {
            const triggers = ScriptApp.getProjectTriggers();
            console.log('⚡ Existing triggers:', triggers.length);
        }
        catch (e) {
            console.log('⚡ Trigger access: Limited (may need manual authorization)');
        }
        console.log('');
        console.log('✅ Permission check completed');
    }
    catch (error) {
        console.error('❌ Permission check failed:', error);
        throw error;
    }
}
