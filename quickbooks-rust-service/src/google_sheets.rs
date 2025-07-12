use anyhow::{Result, Context};
use serde::Serialize;

pub struct GoogleSheetsClient {
    pub webapp_url: String,
    pub api_key: String,
    pub spreadsheet_id: String,
    pub sheet_name: Option<String>,
    pub cell_address: String,
}

#[derive(Serialize)]
struct GoogleSheetsPayload<'a> {
    accountNumber: &'a str,
    accountValue: f64,
    cellAddress: &'a str,
    spreadsheetId: &'a str,
    #[serde(skip_serializing_if = "Option::is_none")]
    sheetName: Option<&'a str>,
    apiKey: &'a str,
}

impl GoogleSheetsClient {
    pub fn new(webapp_url: String, api_key: String, spreadsheet_id: String, sheet_name: Option<String>, cell_address: String) -> Self {
        Self { webapp_url, api_key, spreadsheet_id, sheet_name, cell_address }
    }

    pub async fn send_balance(&self, account_number: &str, account_value: f64, sheet_name: Option<&str>, cell_address: Option<&str>) -> Result<()> {
        let payload = GoogleSheetsPayload {
            accountNumber: account_number,
            accountValue: account_value,
            cellAddress: cell_address.unwrap_or(&self.cell_address),
            spreadsheetId: &self.spreadsheet_id,
            sheetName: sheet_name.or(self.sheet_name.as_deref()),
            apiKey: &self.api_key,
        };
        let client = reqwest::Client::new();
        let res = client.post(&self.webapp_url)
            .json(&payload)
            .send()
            .await
            .context("Failed to send POST to Google Sheets Web App")?;
        if !res.status().is_success() {
            let status = res.status();
            let text = res.text().await.unwrap_or_default();
            anyhow::bail!("Google Sheets Web App returned error: {} - {}", status, text);
        }
        Ok(())
    }
}
