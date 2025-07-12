mod file_mode;
mod config;
mod qbxml_safe;

use anyhow::{Result, Context};
use log::info;
use std::env;

use crate::config::Config;
use crate::file_mode::FileMode;
use crate::qbxml_safe::qbxml_request_processor::QbxmlRequestProcessor;
mod google_sheets;
use google_sheets::GoogleSheetsClient;

#[derive(Debug, Clone)]
pub struct AccountData {
    pub account_full_name: String,
    pub number: String,
    pub account_type: String,
    pub balance: f64,
}

fn print_instructions() {
    println!("QuickBooks Account Query Service v4");
    println!("===================================");
    println!();
    println!("This service reads configuration from config/config.toml and queries");
    println!("the specified account to retreive its balance from QuickBooks Desktop.");
    println!();
    println!("Prerequisites:");
    println!("   1. QuickBooks Desktop and the QuickBooks SDK v16 (or higher) must be installed and running");
    println!("   2. A company file must be open in QuickBooks");
    println!("   3. The FullName of the account in config.toml must exist in QuickBooks");
    println!();
    println!("Usage: main_account_query [--verbose]");
    println!("All account sync blocks are now read from config/config.toml; no account_full_name, sheet_name, or cell_address parameter is required.");
    println!();
}


#[tokio::main]
async fn main() {
    match real_main().await {
        Err(e) => {
            eprintln!("Error: {:#}", e);
            std::process::exit(1);
        },
        Ok(()) => {
            std::process::exit(0);
        }
    }
}

async fn real_main() -> anyhow::Result<()> {
    // Parse arguments
    let args: Vec<String> = env::args().collect();
    let verbose = args.iter().any(|a| a == "--verbose" || a == "-v");

    if verbose {
        print_instructions();
        env_logger::builder().filter_level(log::LevelFilter::Debug).init();
    } else {
        env_logger::builder().filter_level(log::LevelFilter::Info).init();
    }
    // Load configuration
    let config = Config::load_from_file("config/config.toml")
        .context("Failed to load configuration file")?;
    run_qbxml(config).await
}

async fn run_qbxml(config: Config) -> Result<()> {
    unsafe {
        let hr = winapi::um::combaseapi::CoInitializeEx(std::ptr::null_mut(), winapi::um::objbase::COINIT_APARTMENTTHREADED);
        if hr < 0 {
            return Err(anyhow::anyhow!("Failed to initialize COM system: HRESULT=0x{:08X}", hr));
        }
    }

    let processor = QbxmlRequestProcessor::new().context("Failed to create QBXML request processor")?;

    let app_id = config.quickbooks.application_id.as_deref().unwrap_or("QuickBooks-Sheets-Sync");

    let app_name = config.quickbooks.application_name.as_deref().unwrap_or("QuickBooks Sheets Sync");

    processor.open_connection(app_id, app_name)?;

    let company_file = match config.quickbooks.company_file.as_str() { "AUTO" => "", path => path };
    println!("[DEBUG] Company file: {}", company_file);
    let ticket = processor.begin_session(company_file, crate::FileMode::DoNotCare)?;
    match processor.get_account_xml(&ticket) {
        Ok(Some(response_xml)) => {
            let gs_cfg = &config.google_sheets;
            for sync in &config.sync_blocks {
                match processor.get_account_balance(&response_xml, &sync.account_full_name) {
                    Ok(Some(account_balance)) => {
                        info!("[QBXML] Account '{}' balance is: {:?}", sync.account_full_name, account_balance);
                        // Create a new GoogleSheetsClient for each sync block with correct spreadsheet_id and cell_address
                        let gs_client = GoogleSheetsClient::new(
                            gs_cfg.webapp_url.clone(),
                            gs_cfg.api_key.clone(),
                            sync.spreadsheet_id.clone(),
                            Some(sync.sheet_name.clone()),
                            sync.cell_address.clone(),
                        );
                        gs_client.send_balance(
                            &sync.account_full_name,
                            account_balance,
                            Some(&sync.sheet_name),
                            Some(&sync.cell_address),
                        ).await?;
                    },
                    Ok(None) => {
                        info!("[QBXML] No valid balance for account '{}'.", sync.account_full_name);
                    },
                    Err(e) => {
                        eprintln!("[QBXML] Error parsing balance for '{}': {:#}", sync.account_full_name, e);
                    }
                }
            }
        },
        Ok(None) => {
            info!("[QBXML] No response_xml received");
        },
        Err(e) => {
            eprintln!("[QBXML] Error querying Quickbooks: {:#}", e);
        }
    }
    processor.end_session(&ticket)?;
    processor.close_connection()?;
    unsafe { winapi::um::combaseapi::CoUninitialize(); }
    Ok(())
}
