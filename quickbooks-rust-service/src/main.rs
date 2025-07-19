mod file_mode;
mod config;
mod qbxml_safe;

use anyhow::{Result, Context};
use log::info;
use quickbooks_sheets_sync::config::{AccountSyncConfig, TimestampConfig};
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
    println!("QuickBooks Account Query Service v5");
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

async fn process_sync_blocks(processor: &QbxmlRequestProcessor, response_xml: &str, sync_block: &AccountSyncConfig, config: &Config) -> result<()> {
    let gs_cfg = &config.google_sheets;
    match processor.get_account_balance(&response_xml, &sync_block.account_full_name) {
    Ok(Some(account_balance)) => {
        info!("[QBXML] Account '{}' balance is: {:?}", sync_block.account_full_name, account_balance);
        let gs_client = GoogleSheetsClient::new(
            gs_cfg.webapp_url.clone(),
            gs_cfg.api_key.clone(),
            sync.spreadsheet_id.clone(),
            );
        gs_client.send_balance(
            account_balance,
            Some(&sync_block.sheet_name),
            Some(&sync_block.cell_address),
            ).await?;
            },
        Ok(None) => {
          info!("[QBXML] No valid balance for account '{}'.", sync_block.account_full_name);
            },
        Err(e) => {
            eprintln!("[QBXML] Error parsing balance for '{}': {:#}", sync_block.account_full_name, e);
            }
    }
    Ok(())
}

async fn process_timestamp_blocks(timestamp_block: &TimestampConfig, config: &Config, ) -> Result<()> {
    use chrono::Local;
    let gs_cfg = &config.google_sheets;
    let now = Local::now();
    let formatted_time = now.format("%d-%m-%Y:%H:%M").to_string();
    let gs_client = GoogleSheetsClient::new(
        gs_cfg.webapp_url.clone(),
        gs_cfg.api_key.clone(),
        gs_cfg.spreadsheet_id.clone(),
        );
    gs_client.send_timestamp(
        Some(&formatted_time), 
        Some(&timestamp_block.sheet_name),
        Some(&timestamp_block.cell_address),
        ).await?;
    Ok(())
}

async fn process_qbxml(processor: &QbxmlRequestProcessor, response_xml: &str, config: Config) -> Result<()> {
    for sync_block in &config.sync_blocks {
        process_sync_blocks(&processor, &response_xml, &sync_block, &config).await?;
    }
    // Inject timestamp after all sync_blocks processed
    for timestamp_block in &config.timestamp_blocks {
        process_timestamp_blocks(&timestamp_block, &config, ).await?;
        }
    Ok(())
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
    
    // AppID isn't used by the QBSDK
    let app_id = config.quickbooks.application_id.as_deref().unwrap_or(""); 
    // If we ever change the name of the service we register with Quickbooks we'll have to change this default too
    let app_name = config.quickbooks.application_name.as_deref().unwrap_or("QuickBooks Sync Service"); 
    
    processor.open_connection(app_id, app_name)?;

    let company_file = match config.quickbooks.company_file.as_str() { 
        "AUTO" => "",
        path => {
            println!("[DEBUG] Company file: {}", path);
            path }
        };
    
    let ticket = processor.begin_session(company_file, crate::FileMode::DoNotCare)?;

    match processor.get_account_xml(&ticket) {
        Ok(Some(response_xml)) => {
            process_qbxml(&processor, &response_xml, config).await?;
        },
        Ok(None) => {
            eprintln!("[QBXML] No response_xml received");
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
