Import-Module Microsoft.Graph.Users.Actions
 
$clientID = ""
$tenantId = ""
$certificateName = ""
 
# Authenticate with Graph API
Connect-MgGraph -ClientID $clientID -TenantId $tenantId -CertificateName $certificateName
 
# Import the licence details CSV
$plansCSV = Import-Csv -Path "C:\source\scripts\M365 Reporting\licenceDetails.csv"
 
# Get the users
$users = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable licensedUserCount -All -Select UserPrincipalName,DisplayName,AssignedLicenses,JobTitle,Department
 
Write-Output("The total amount of licenced users is " + $licensedUserCount)

$customUsers = New-Object System.Collections.ArrayList
 
# Get the licence names
foreach ($user in $users) {
    $licenceNames = ""
 
    # Map the licence SKU to the licence's name
    foreach ($licence in $user.AssignedLicenses) {
        $licenceDetails = $plansCSV | Where-Object { $_.GUID -eq $licence.skuId } | Select-Object -Unique
        $licenceNames += $licenceDetails.Product_Display_Name + ", "
    }
 
    # Create a custom user object
    $customUser = @{
        DisplayName = $user.DisplayName
        UPN = $user.UserPrincipalName
        JobTitle = $user.JobTitle
        Department = $user.Department
        Licences = $licenceNames
    }
    $customUserObject = new-object psobject -Property $customUser
 
    # Add it to the collection
    $customUsers.Add($customUserObject)
}
 
# Display a table of all the licenced users
$customUsers | Format-Table -Property UPN,DisplayName,JobTitle,Department,Licences

# Create a csv of all the licenced users
$reportCSV = $customUsers | Export-Csv -Path "C:\source\scripts\M365 Reporting\Exports\report.csv"
 
 
# Send the email
$params = @{
    Message = @{
        Subject = "M365 Licence Report"
        Body = @{
            ContentType = "Text"
            Content = "Hi, I hope that you're well. Please find attached a report of all current users and the licences that they have assigned. Kind regards,"
        }
        ToRecipients = @(
            @{
                EmailAddress = @{
                    Address = ""
                }
            }
        )
        CcRecipients = @(
            @{
                EmailAddress = @{
                    Address = ""
                }
            }
        )
    }
    SaveToSentItems = "false"
}
Send-MgUserMail -UserId "" -BodyParameter $params
 
 
# Disconnect from Graph API
Disconnect-MgGraph
