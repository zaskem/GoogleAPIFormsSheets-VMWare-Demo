# GoogleAPIFormsSheets-VMWare-Demo
Example files for running a demo between Google Forms, Sheets and VMWare to manage the lifecycle of a "throwaway" virtual machine. Primarily Powershell-based, with generalized Google Apps Script source for Form (input) and Sheet (response) management.

With some additions (credentials and setting variables), this is production-ready code for reading, processing against VMWare, and writing VMWare data from a Google sheet via the Google REST API. **This code will possibly DO IRREVOCABLE THINGS WITHOUT DIRECT CONSENT within a valid VMWare environment!**
## Dependencies
Requires any current version of the following Powershell modules:
* `UMN-Google`
* `VMWare.PowerCLI`
The VMWare PowerCLI module does require a bit of initial set-up, depending on your VMWare infrastructure situation. Your Mileage May Vary.
## The Idea:
The driving force behind all of this is to take action in VMWare based on the results of Google form submissions. Essentially, using Google forms to "standardize" a process to create and dispose of might be called "throwaway" or "test" Virtual Machines.
## The Execution:
Each script of the repo handles a different action, identified in its title. These particular scripts have been configured to all use the same base Google sheet for simplicity, but function across various tabs/sheets within the document. Customizing variables appropriately, you could cross multiple Google sheets/documents as your needs determine.
* `VMWare-CreateRequestedVMsFromVM.ps1`
    * Clones a VM from an existing VM
* `VMWare-RemoveRequestedVMs.ps1`
    * Permanently removes a VM
* `GoogleAppsScriptExamples.gs`
    * Code examples to use with Google Form/Sheet triggers to self-clean and self-update form responses and input options
## Credential Files
No credential files are included. _You will need to obtain your own project-specific key_ via https://console.developers.google.com and reference as necessary. The Google API key used in this demo must be the `.p12` version and not the `.json` version. Additionally, the referenced `svcacct-example` service account credential **must be generated** via a separate process (TL;DR: use the Powershell cmdlet `ConvertTo-SecureString` and `ConvertFrom-SecureString` to create an input file for using service account credentials to access VMWare).
