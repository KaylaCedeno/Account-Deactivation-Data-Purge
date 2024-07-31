# Importing modules so that its cmdlets/functions can be used in current session

Install-Module -Name ImportExcel 
Import-Module ImportExcel 

Install-Module -Name ExchangeOnlineManagement 
Connect-ExchangeOnline 

Install-Module -Name AzureAD 
Import-Module AzureAD 
Connect-AzureAD 


# Importing excel sheet for session

$ExcelLink = path_file 
$Sheet = Import-Excel –Path $ExcelLink 

foreach($col in $Sheet) {

# Variables holding the user's account, end date, and threshold date 

        $Account = $col.'Account'
        $EndDate = [DateTime]$col.'End Date'
        $Tresh = [DateTime]"07/23/2024"

        if($EndDate -ge $Tresh) {
          try {
          
# Updating users mailbox attribute so that it contains their end date
          
           Set-Mailbox $Account -CustomAttribute7 $EndDate
           Write-Host "Updated for $Account"

           }

           catch {
            Write-Host "Error for $Account : $_"

           }
        }

# Removes the user's account from the system 

        else {
            try {
             Remove-AzureADUser -ObjectId $Account
             Write-Host "Deleted Azure AD User: $Account"

            }

            catch {
             Write-Host "Error deleting user $Account : $_"

            }
        }
  }
