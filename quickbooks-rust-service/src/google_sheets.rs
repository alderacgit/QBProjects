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
    #[serde(rename = "accountNumber")]
    account_number: &'a str,
    #[serde(rename = "accountValue")]
    account_value: f64,
    #[serde(rename = "cellAddress")]
    cell_address: &'a str,
    #[serde(rename = "spreadsheetId")]
    spreadsheet_id: &'a str,
    #[serde(rename = "sheetName", skip_serializing_if = "Option::is_none")]
    sheet_name: Option<&'a str>,
    #[serde(rename = "apiKey")]
    api_key: &'a str,
    #[serde(rename = "stringValue", skip_serializing_if = "Option::is_none")]
    string_value: Option<&'a str>,
}

impl GoogleSheetsClient {
    pub fn new(webapp_url: String, api_key: String, spreadsheet_id: String, sheet_name: Option<String>, cell_address: String) -> Self {
        Self { webapp_url, api_key, spreadsheet_id, sheet_name, cell_address }
    }

    pub async fn send_balance(&self, account_number: &str, account_value: f64, sheet_name: Option<&str>, cell_address: Option<&str>, string_value: Option<&str>) -> Result<()> {
        let payload = GoogleSheetsPayload {
            account_number: account_number,
            account_value: account_value,
            cell_address: cell_address.unwrap_or(&self.cell_address),
            spreadsheet_id: &self.spreadsheet_id,
            sheet_name: sheet_name.or(self.sheet_name.as_deref()),
            api_key: &self.api_key,
            string_value: string_value,
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
