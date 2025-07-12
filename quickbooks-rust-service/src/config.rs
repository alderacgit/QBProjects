use anyhow::{Context, Result};
use figment::{Figment, providers::{Format, Toml}};
use serde::{Deserialize, Serialize};
use std::path::Path;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Config {
    pub quickbooks: QuickBooksConfig,
    pub google_sheets: GoogleSheetsConfig,
    pub sync_blocks: Vec<AccountSyncConfig>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AccountSyncConfig {
    pub spreadsheet_id: String,
    pub account_full_name: String,
    pub sheet_name: String,
    pub cell_address: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct QuickBooksConfig {
    pub enabled: Option<bool>,
    pub company_file: String,
    pub connection_mode: Option<String>,
    pub application_name: Option<String>,
    pub application_id: Option<String>,
    pub connection_timeout: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GoogleSheetsConfig {
    pub webapp_url: String,
    pub api_key: String,
    pub sheet_name: Option<String>,
    pub cell_address: Option<String>,
    // spreadsheet_id is now per sync_block, not required here
    // pub spreadsheet_id: Option<String>,
}



impl Config {
    pub fn load_from_file<P: AsRef<Path>>(path: P) -> Result<Self> {
        let figment = Figment::from(Toml::file(path));
        figment.extract().context("Failed to parse config file")
    }
}
