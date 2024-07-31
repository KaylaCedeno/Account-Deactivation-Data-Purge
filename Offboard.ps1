Install-Module -Name ImportExcel 
Import-Module ImportExcel 

Install-Module -Name ExchangeOnlineManagement 
Connect-ExchangeOnline 

Install-Module -Name AzureAD 
Import-Module AzureAD 
Connect-AzureAD 

$ExcelLink = path_file 
$Sheet = Import-Excel –Path $ExcelLink 


foreach($col in $Sheet) {

        $Account = $col.'Account'
        $EndDate = [DateTime]$col.'End Date'
        $Tresh = [DateTime]"07/23/2024"

        if($EndDate -ge $Tresh) {
          try {
           Set-Mailbox $Account -CustomAttribute7 $EndDate
           Write-Host "Updated for $Account"

           }

           catch {
            Write-Host "Error for $Account : $_"

           }
        }

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
