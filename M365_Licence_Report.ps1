Import-Module Microsoft.Graph.Users.Actions
 
$clientID = ""
$tenantId = ""
$certificateName = ""

$licenceDetailsCSVPath = ""
$reportExportPath = ""
 
# Authenticate with Graph API
Connect-MgGraph -ClientID $clientID -TenantId $tenantId -CertificateName $certificateName
 
# Import the licence details CSV
$plansCSV = Import-Csv -Path $licenceDetailsCSVPath
 
# Get the users
$users = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable licensedUserCount -All -Select UserPrincipalName,DisplayName,AssignedLicenses,JobTitle,Department
 
Write-Output("The total amount of licenced users is " + $licensedUserCount)

$customUsers = New-Object System.Collections.ArrayList
 
# Get the licence names
foreach ($user in $users) {
    $licenceNames = ""
    $licenceCount = 0
 
    # Map the licence SKU to the licence's name
    foreach ($licence in $user.AssignedLicenses) {
        $licenceDetails = $plansCSV | Where-Object { $_.GUID -eq $licence.skuId } | Select-Object -Unique
        $licenceNames += $licenceDetails.Product_Display_Name + ", "
        $licenceCount += 1
    }

    # Add the licence count to the end
    $licenceNames += "(" + $licenceCount + ")"
 
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
$customUsers | select DisplayName, UPN, JobTitle, Department, licences | Export-Csv -Path $reportExportPath -NoTypeInformation

# Convert the csv file to base64
$base64string = [Convert]::ToBase64String([IO.File]::ReadAllBytes($reportExportPath))

# Send the email
$params = @{
    Message = @{
        Subject = "M365 Licence Report"
        Body = @{
            ContentType = "Text"
            Content = "Hi,`r`rI hope that you're well.`r`rPlease find attached a report of all current users and the licences that they have assigned.`r`rKind regards,`r`r"
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
        Attachments = @(
			@{
				"@odata.type" = "#microsoft.graph.fileAttachment"
				Name = "usersReport.csv"
				ContentType = "text/plain"
				ContentBytes = $base64string
			}
		)
    }
    SaveToSentItems = "false"
}
Send-MgUserMail -UserId "" -BodyParameter $params
 
 
# Disconnect from Graph API
Disconnect-MgGraph
