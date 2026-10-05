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
* IMPORTANT CAVEAT ON THAT IDL: it is the type library for **QBFC16Lib**, and it says nothing about QBXMLRP2. It contains no reference to `QBXMLRP2` or `RequestProcessor` anywhere. Since we call `QBXMLRP2.RequestProcessor` and not QBFC, the IDL is evidence about QBFC only. Do not carry a constant, an enum value or a parameter list from it into a QBXMLRP2 call and treat the result as sourced -- that is extrapolation between two different type libraries, which is the very thing the rule above forbids. Our `FileMode` numbers are an example of the mistake: the comment in `begin_session` used to cite `omDontCare = 2` from this IDL, but `omDontCare` belongs to QBFC's `ENOpenMode`. Only `DoNotCare => 2` is actually established, by the fact that it works. To settle the rest, dump the QBXMLRP2 type library on the Windows host and read it from there
* There are two mechanisms to use the SDK - QBFC and QBXML. We have discovered that there is a fatal and unfixable flaw in QBFC which prohibits our use of that part of the SDK. We have to exclusively use QBXML
* Parameter order through `IDispatch::Invoke` is reversed, and this is a rule rather than a surprise. `DISPPARAMS::rgvarg` is ordered **last argument first**: `rgvarg[0]` holds the rightmost parameter of the documented signature. This is the standard COM calling convention and it applies to every multi-argument call, not to some of them.
  This note used to say the opposite -- that reversal was needed "sometimes but not always", that only one of five tested functions required it, and that reasoning from how COM works would lead you astray. That was a misreading of our own code. All five call sites were in fact passing arguments reversed; `invoke_method` handed its `params` slice straight to `rgvarg` without reordering, so a call site that looked "reversed" was simply correct, and `EndSession`/`CloseConnection` looked unreversed only because they take one argument or none. There was never a function that behaved differently from the others.
  `invoke_method` now performs the reversal itself, so call sites pass arguments in documented left-to-right order and new SDK calls are correct the first time. The `rgvarg` layout sent to QuickBooks is unchanged by that refactor. If you add a call, write the parameters as the SDK documents them and do not reverse them by hand
* There are C-style unions (specifically VARIANT) in the Windows COM & OLE system that need to be wrapped so Rust can use them safely. We worked hard to get the windows crate to do this and failed; our fallback was to use the winapi crate which succeeded. In general when working with COM or OLE you need to use a wrapper from the qbxml_safe directory and the qbxml_safe_variants.rs file. You should always wrap and never write code from scratch. WHen in doubt ask for direction. In particular if you find yourself using the construction Anonymous.Anonymous or Anoynous.Anonymous.Anonymous you have strayed from the righteous path and need to seek enlightement in qbxml_safe_variant.rs
