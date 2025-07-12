use anyhow::{Context, Result};
use log::{info, warn};
use reqwest::Client;
use serde::{Deserialize, Serialize};

use crate::{AccountData, config::{Config, GoogleSheetsConfig}};

#[derive(Debug, Clone)]
pub struct GoogleSheetsClient {
    client: Client,
    config: GoogleSheetsConfig,
}

#[derive(Debug, Serialize)]
struct SheetsUpdateRequest {
    #[serde(rename = "accountNumber")]
    account_number: String,
    #[serde(rename = "accountValue")]
    account_value: f64,
    #[serde(rename = "spreadsheetId")]
    spreadsheet_id: String,
    #[serde(rename = "cellAddress")]
    cell_address: String,
    #[serde(rename = "sheetName")]
    sheet_name: Option<String>,
    #[serde(rename = "apiKey")]
    api_key: String,
}

#[derive(Debug, Deserialize)]
struct SheetsUpdateResponse {
    success: bool,
    message: Option<String>,
    error: Option<String>,
}

impl GoogleSheetsClient {
    pub fn new(config: &GoogleSheetsConfig) -> Result<Self> {
        let client = Client::builder()
            .timeout(std::time::Duration::from_secs(30))
            .user_agent("QuickBooks-Sheets-Sync/1.0")
            .build()
            .context("Failed to create HTTP client")?;

        Ok(Self {
            client,
            config: config.clone(),
        })
    }

    pub async fn test_connection(&self) -> Result<()> {
        info!("Testing Google Sheets connection...");
        
        // Create a test request with a mock account balance
        let test_request = SheetsUpdateRequest {
            account_number: "TEST".to_string(),
            account_value: 0.0,
            spreadsheet_id: self.config.spreadsheet_id.clone(),
            cell_address: self.config.cell_address.clone(),
            sheet_name: self.config.sheet_name.clone(),
            api_key: self.config.api_key.clone(),
        };

        // Send test request
        let response = self.client
            .post(&self.config.webapp_url)
            .json(&test_request)
            .send()
            .await
            .context("Failed to send test request to Google Sheets")?;

        let status = response.status();
        let response_text = response.text().await
            .context("Failed to read response body")?;

        if status.is_success() {
            info!("Google Sheets connection test successful");
            Ok(())
        } else {
            warn!("Google Sheets connection test failed with status: {}", status);
            warn!("Response: {}", response_text);
            Err(anyhow::anyhow!("Google Sheets connection test failed: {}", response_text))
        }
    }

    pub async fn update_account_data(&self, account_data: &AccountData, config: &Config) -> Result<()> {
        info!("Updating Google Sheets with account data: {} = {}", 
              account_data.account_number, 
              account_data.balance);

        let request = SheetsUpdateRequest {
            account_number: account_data.account_number.clone(),
            account_value: account_data.balance,
            spreadsheet_id: config.google_sheets.spreadsheet_id.clone(),
            cell_address: config.google_sheets.cell_address.clone(),
            sheet_name: config.google_sheets.sheet_name.clone(),
            api_key: config.google_sheets.api_key.clone(),
        };

        let response = self.client
            .post(&config.google_sheets.webapp_url)
            .json(&request)
            .send()
            .await
            .context("Failed to send update request to Google Sheets")?;

        let status = response.status();
        
        if status.is_success() {
            let response_body: SheetsUpdateResponse = response
                .json()
                .await
                .context("Failed to parse response from Google Sheets")?;

            if response_body.success {
                info!("Successfully updated Google Sheets: {}", 
                      response_body.message.unwrap_or("No message".to_string()));
            } else {
                let error_msg = response_body.error.unwrap_or("Unknown error".to_string());
                return Err(anyhow::anyhow!("Google Sheets update failed: {}", error_msg));
            }
        } else {
            let error_text = response.text().await
                .unwrap_or_else(|_| "Failed to read error response".to_string());
            return Err(anyhow::anyhow!("HTTP error {}: {}", status, error_text));
        }

        Ok(())
    }

    #[allow(dead_code)]
    pub async fn get_current_balance(&self, _config: &Config) -> Result<f64> {
        info!("Getting current balance from Google Sheets");

        // This would require implementing a GET endpoint in your Google Apps Script
        // For now, we'll return a placeholder
        warn!("get_current_balance not implemented - returning 0.0");
        Ok(0.0)
    }
}
