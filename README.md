# QuickBooks-Sheets Integration Project

This project enables automated synchronization between QuickBooks Desktop Enterprise and Google Sheets, consisting of two main components that work together to extract account data from QuickBooks and update specific cells in Google Sheets.

## Project Components

### 1. QuickBooks Rust Program (`quickbooks-rust-service`)

A Windows program written in Rust that:
- Runs on Windows
- Connects to QuickBooks Desktop Enterprise using the QuickBooks SDK v16
- Extracts account balances from an open company file
- Sends the data to Google Sheets via a web app endpoint
- Supports multiple account synchronization blocks configured via TOML

Key Features:
- Uses QBXML for QuickBooks communication (QBFC implementation was deprecated due to limitations)
- Safe wrapping of Windows COM & OLE interfaces
- Configurable through `config/config.toml`
- Supports multiple account-to-cell mappings
- Provides a Timestamp service to inject timestamps into Google Sheets

### 2. Google Sheets Integration (`Google_Sheet_Functions`)

A Google Apps Script project that:
- Provides a web endpoint to receive QuickBooks data
- Updates specific cells in Google Sheets with account balances
- Updates specific cells in Google Sheets with timestamps
- Includes security measures with API key authentication
- Supports multiple spreadsheets and sheets

Key Features:
- Secure API endpoint for receiving QuickBooks data
- Functions for manual testing and updates
- Comprehensive setup and permission management
- Error handling and logging

## Prerequisites

### Windows Requirements
- QuickBooks Desktop Enterprise v24 (64-bit)
- QuickBooks SDK v16 (or higher)
- Application must register with QuickBooks which requires company to be in single-user mode and user to be Administrator

### Google Sheets Requirements
- Google Workspace account with appropriate permissions
- Google Apps Script enabled
- API key configured (generated during setup)

## Setup Instructions

### Google Sheets Setup

1. Deploy the Google Apps Script project:
   ```
   1. Open Google Apps Script editor
   2. Copy the contents of Google_Sheet_Function/src/Code.ts
   3. Run setupPermissions()
   4. Run setupQuickBooksIntegration() to generate API key
   5. Deploy as Web App with "Anyone" access
   6. Test with testWebAppEndpoint()
   ```

2. Note the generated API key and Web App URL for the Rust program configuration

### QuickBooks Rust Program Setup

1. Build and deploy:
   ```
   1. Pull the project from github to a working directory (https://github.com/alderacgit/QBProjects)
   2. Build the service using MSVC toolchain and cargo build
   3. Deploy qb_sync.exe to the Windows server (typically in \Program Files (x86)\Alderac\qb_sync
   4. Configure Windows Task Scheduler for periodic execution
   5. Be sure that you have the program start in a directory with a child directory called config which contains the config.toml file
   ```

2. Configure the program:
   ```
   1. Copy config/config.example.toml to config/config.toml
   2. Update configuration toml with:
      - QuickBooks application details
      - Google Sheets API key
      - Web App URL
      - Account sync blocks
      - Timestamp blocks [Optional]
   ```

## Configuration (config\config.toml)

Spreadsheet_id values are the gibberish code in a standard Google Sheet URL after /d/ and before /edit?...
Sheet_name is the name of the tab on the sheet you want to work with

```toml
[quickbooks]
application_id = "QuickBooks-Sheets-Sync"
application_name = "QuickBooks Sheets Sync"
company_file = "AUTO"  # or specify path

[google_sheets]
webapp_url = "Your-Google-Web-App-URL"
api_key = "Your-API-Key"

[[sync_blocks]]
account_full_name = "Account Name in QuickBooks"
spreadsheet_id = "Google-Spreadsheet-ID"
sheet_name = "Sheet Name"
cell_address = "A1"

[[timestamp_blocks]]
spreadseet_id = "Google-Spreadsheet-ID"
sheet_name = "Sheet Name"
cell_address = "A1"


```

## Development Notes

### QuickBooks SDK Considerations
- Use QBXML exclusively (QBFC has known limitations)
- Parameter ordering may differ from documentation
- Use safe wrappers from qbxml_safe directory for COM/OLE interactions
- Reference QBFC16 COM OLE Data.IDL for API definitions

## Security Notes

- API key authentication required for all requests

## Troubleshooting

1. QuickBooks Connectivity:
   - Verify SDK registration

2. QuickBooks Configuration
   - You can only have one instance of QBW.EXE running on your computer; if the system tries to open a "Second Quickbooks" you have at least two copies running and/or you are trying to run the program from an account other than the one you have already opened QuickBooks in
   - You can run the Rust program without having QuickBooks open if you have provided the proper permissions in Quickbooks when the program registers itself
   - Don't run the program as SYSTEM, it has to run as a regular Windows user account

2. Google Sheets Issues:
   - Verify API key configuration
   - Check sheet and cell permissions
   - Review Apps Script logs

## Contributing

When contributing to this project:
- Follow existing code patterns and practices
- Use safe wrappers for COM/OLE interactions
- Test thoroughly before deployment
- Document any SDK quirks encountered

## License

Code made available using the MIT license
