This is an umbrella project with two separate but related projects

# Google_Sheet_Functions
This is a project that will build a Google Script function which will be used to update a specific cell on a specific Google Sheet

# quickbooks-rust-service
This is a project that will run on a Windows 2019 Server computer as a service and periodically extract values from a Quickbooks Enterprise Desktop company file, and then use the Google_Sheet_Functions tools to get those values onto their target Google Sheets and cells

# AI Instructions
Two computers are being used in this project. One is a Macintosh the other is a Windows 2019 Server. When building on the Macintosh we are targeting the Windows 2019 Server using MSVC and it is ok to build even though the build process will fail when it is time to link; the most important part of the build process is error checking not linking

When issuing terminal commands do not use &&; issue each command as a separate command

The build process can take several minutes; don't assume that the project built successfully - you may need to wait to analyze the results of the build for some minutes. When in doubt, ask if the build has completed, don't rely on echoing strings to the Terminal to detect the completion of the build

# Windows 2019 Server Environment

The Windows 2019 Server has the following specifications:
1. QuickBooks Desktop Enterprise v24 64-bit is installed
2. A company file is open
3. A user is logged in with Administrative rights
4. The Quickbooks SDK v16 is installed and all the dlls have been successfully registered
5. The qb_sync.exe program that gets built in the quickbooks-rust-serivce directory has been tested and is working

Generally speaking if something is not working always assume the problem lies with our code not with some element of the software or services installed on the Windows 2019 Server

# Quirks of the QuickBooks SDK
During our work to create the qb_sync.exe tool we discovered many things about the QuickBooks SDK including
* The documentation that used to be online from inuit has been removed and is no longer avilable for many parts of the API; do not guess or extrapolate from other sources when asked questions about the API documentation. If you cannot source the documentation from intuit for an answer, don't provide an answer just say you don't have a source
* We have extracted data on the COM and OLE objects and that information is in the QBFC16 COM OLE Data.IDL file at the root of this repository; it contains many useful definitions for the API functions and parameters
* There are two mechanisms to use the SDK - QBFC and QBXML. We have discovered that there is a fatal and unfixable flaw in QBFC which prohibits our use of that part of the SDK. We have to exclusively use QBXML
* Sometimes but not always when passing parameters to the QB SDK functions we have to pass them in the reverse order as specified by the documentation. The only way to determine when that is necessary is by debugging working code that is failing to receive proper responses from QuickBooks. You may be tempted to recommend providing parameters in the reverse order based on your theory of how COM or OLE works but you will be wrong; of the five functions we have tested only one required reverse paramater ordering
* There are C-style unions (specifically VARIANT) in the Windows COM & OLE system that need to be wrapped so Rust can use them safely. We worked hard to get the windows crate to do this and failed; our fallback was to use the winapi crate which succeeded. In general when working with COM or OLE you need to use a wrapper from the qbxml_safe directory and the qbxml_safe_variants.rs file. You should always wrap and never write code from scratch. WHen in doubt ask for direction. In particular if you find yourself using the construction Anonymous.Anonymous or Anoynous.Anonymous.Anonymous you have strayed from the righteous path and need to seek enlightement in qbxml_safe_variant.rs
