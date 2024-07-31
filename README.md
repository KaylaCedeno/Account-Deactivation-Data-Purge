# PowerShell Script for Managing Office 365 and Azure AD Users 

This PowerShell script automates the process of updating custom attributes for Office 365 mailboxes and removing Azure AD users based on the end date specified in an Excel sheet. It utilizes the ImportExcel, ExchangeOnlineManagement, and AzureAD modules.

# Notes

- Ensure that your Excel file has the columns named Account and End Date
- Modify the $Thresh variable to set the threshold date for your specific use case
- The script includes basic error handling (try & catch), including the $_ command which prints errors in the console with the relevant account information
