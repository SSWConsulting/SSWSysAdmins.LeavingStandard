. "C:\inetpub\wwwroot\SSWLeavingStandard\Functions.ps1"

Import-Module universaldashboard.community

$Theme = New-UDTheme -Name "Basic" -Definition @{
    #UDDashboard = @{
    #BackgroundColor = "rgb(204,65,65)"
    #FontColor = "rgb(51,51,51)"
    #}
    '.btn'    = @{
        'background-color' = 'rgb(204,65,65)'
    }
    UDNavBar  = @{
        BackgroundColor = "rgb(204,65,65)"
        FontColor       = "rgb(51,51,51)"
    }
    UDFooter  = @{
        BackgroundColor = "rgb(204,65,65)"
        FontColor       = "rgb(51,51,51)"
    }
    UDCard    = @{
        BackgroundColor = "rgb(204,65,65)"
        FontColor       = "rgb(51,51,51)"
    }
    UDChart   = @{
        BackgroundColor = "rgb(204,65,65)"
        FontColor       = "rgb(51,51,51)"
    }
    UDCounter = @{
        BackgroundColor = "rgb(204,65,65)"
        FontColor       = "rgb(51,51,51)"
    }
    UDMonitor = @{
        #BackgroundColor = "rgb(204,65,65)"
        #FontColor = "rgb(51,51,51)"
    }
    UDGrid    = @{
        #BackgroundColor = "rgb(204,65,65)"
        #FontColor = "rgb(51,51,51)"
    }
    UDInput   = @{
        BackgroundColor = "rgb(204,65,65)"
        FontColor       = "rgb(51,51,51)"
    }
    UDTable   = @{
        #BackgroundColor = "rgb(204,65,65)"
        #FontColor = "rgb(51,51,51)"
    }
    UDButton  = @{
        BackgroundColor = "rgb(204,65,65)"
        FontColor       = "rgb(51,51,51)"
    }
}

$Page1 = New-UDPage -Name "Home" -Icon Home -Content {
   
    New-UDRow {
        New-UDHeading -Size 4 -Content {
            "*IMPORTANT* - Authenticate below before you do anything! And follow the SugarLearning procedure: https://my.sugarlearning.com/companies/SSW/items/13045/Disable-your-Accounts "
        }
        New-UDHeading -Size 6 -Content {
            "Use your admin account like this (non-case-sensitive): SSW2000\adminkaiquebiancatti"
        }
        New-UDHeading -Size 6 -Content {
            "Search for an AD user like this (non-case-sensitive, and given name and surname together): adamcogan"
        }
    }
    
    New-UDElement -Tag "div" -Attributes @{
        style = @{
            height = '25px'
        }
    }
    New-UDRow -Columns {
        New-UDColumn -Size 10 -SmallOffset 1 -Content {
            New-UDRow -Columns {
                New-UDColumn -Size 10 -Content {
                    New-UDTextbox -Id "Username" -Label "Username (AD Access - Necessary for Search)" -Placeholder "e.g. adminkaiquebiancatti" -Icon search
                }
                New-UDColumn -Size 10 -Content {
                    New-UDTextbox -Id "Password" -Type 'password' -Label "Password" -Placeholder "Password" -Icon search
                }
                New-UDColumn -Size 10 -Content {
                    New-UDTextbox -Id "txtSearch" -Label "Search" -Placeholder "e.g. adamcogan" -Icon search
                }
                New-UDColumn -Size 2 -Content {
                    New-UDButton -Id "btnSearch" -Text "Search User" -Icon search -OnClick {
                        $User = Get-UDElement -Id "Username"
                        $AdminUsername = $User.Attributes["value"]

                        $Pass = Get-UDElement -Id "Password"
                        $AdminPassword = ConvertTo-SecureString $Pass.Attributes["value"] -AsPlainText -Force 

                        $Element = Get-UDElement -Id "txtSearch" 
                        $Value = $Element.Attributes["value"]
                        
                        $Session:Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList ($AdminUsername, $AdminPassword)

                        Set-UDElement -Id "results" -Content {
                            New-UDGrid -Title "Search Results for: $Value" -Headers @("Name", "More Info") -Properties @("Name", "MoreInfo") -Endpoint {
                                $Objects = Get-ADUser -Filter "Name -like '$Value' -or samAccountName -like '$Value'" -Server "ssw-dc4.sydney.ssw.com.au" -Credential $Session:Credentials
                                $Objects | ForEach-Object {
                                    [PSCustomObject]@{
                                        Name     = $_.Name
                                        #SAMAccountName = $_.samAccountName
                                        MoreInfo = New-UDButton -Text "Select" -OnClick {
                                            $Session:SelectedUser = $_
                                            Show-UDToast -Message "Selected"
                                        }
                                    }
                                } | Out-UDGridData 
                            } 
                        }
                    }
                }
            }
        }
    }

    New-UDRow -Columns {
        New-UDColumn -SmallSize 10 -SmallOffset 1 {
            New-UDElement -Tag "div" -Id "results"
        }
    }
    
    New-UDRow -Columns {
        New-UDColumn {
            New-UDHtml -Markup "<h5><b>*IMPORTANT*</b> - Before anything else, go to <a href=https://protection.office.com/contentsearchbeta>https://protection.office.com/contentsearchbeta</a> | New Search | Specific Locations | Modify... | Choose users, groups or teams... | Search for the email | Export Results... | Download the mailbox to \\fileserver\DataSSW\ExEmployees\ExchangeBackup\{username}.pst</h5>"
        }
    }
    
    # First Row
    New-UDRow -Endpoint {
        New-UDColumn -Size 3 -Endpoint {
            New-UDButton -Text "2. Backup Search (Start Here)" -OnClick {
                Set-UDElement -id "icon2" -Content {
                    New-UDIcon -icon pause
                    "Started fileserver backups"
                }
                try {
                    New-backupsearch -User $Session:SelectedUser.samAccountName
                    Set-UDElement -id "icon2" -Content {
                        New-UDIcon -icon check -Color green
                        "Finished fileserver backups - SugarLearning Step 2"
                    }
                }
                catch {
                    Set-UDElement -id "icon2" -Content {
                        $LastError = $Error[0]
                        New-UDIcon -icon ban -color red
                        "Error (You probably forgot to search for an user above) - $LastError"
                    }
                }
            }
            New-UDElement -Id "icon2" -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDColumn -Size 3 -Endpoint {
            New-UDButton -Text "3. Remove Groups" -OnClick {
                Set-UDElement -id "icon3" -Content {
                    New-UDIcon -icon pause
                    "Started removal of user groups"
                }
                try {
                    Remove-UserFromAllADGroups -User $Session:SelectedUser.samAccountName
                    Set-UDElement -id "icon3" -Content {
                        New-UDIcon -icon check -color green
                        "Removed user from all AD groups - SugarLearning Step 3"
                    }
                }
                catch {
                    Set-UDElement -id "icon3" -Content {
                        $LastError = $Error[0]
                        New-UDIcon -icon ban -color red
                        "Error - $LastError"
                    }
                }
            }
            New-UDElement -Id "icon3" -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDColumn -Size 3 -Endpoint {
            New-UDButton -Text "4/5. Disable AD User" -OnClick {
                Set-UDElement -id "icon45" -Content {
                    New-UDIcon -icon pause
                    "Started disabling user"
                }
                try {
                    Disable-User -User $Session:SelectedUser
                    Set-UDElement -id "icon45" -Content {
                        New-UDIcon -icon check -color green
                        "Finished disabling user - SugarLearning Step 4 and 5"
                    }
                }
                catch {
                    Set-UDElement -id "icon45" -Content {
                        $LastError = $Error[0]
                        New-UDIcon -icon ban -color red
                        "Error - $LastError"
                    }
                }
            }
            New-UDElement -Id "icon45" -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDColumn -Size 3 -Endpoint {
            New-UDButton -Text "6. Disable S4B User" -OnClick {
                Set-UDElement -id "icon6" -Content {
                    New-UDIcon -icon pause
                    "Started disabling Skype for Business user"
                }
                try {
                    Disable-S4BUser -User $Session:SelectedUser.UserPrincipalName
                    Set-UDElement -id "icon6" -Content {
                        New-UDIcon -icon check -color green
                        "Finished disabling Skype for Business user - SugarLearning Step 6"
                    }
                }
                catch {
                    Set-UDElement -id "icon6" -Content {
                        $LastError = $Error[0]
                        New-UDIcon -icon ban -color red
                        "Error - $LastError"
                    }
                }
            }
            New-UDElement -Id "icon6" -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
    }
 
    # Second Row
    New-UDRow -Endpoint {
        New-UDColumn -Size 3 -Endpoint {
            New-UDButton -Text "7. Hide from Email List" -OnClick {
                Set-UDElement -id "icon7" -Content {
                    New-UDIcon -icon pause
                    "Started hiding from email list"
                }
                try {
                    Hide-Email -User $Session:SelectedUser
                    Set-UDElement -id "icon7" -Content {
                        New-UDIcon -icon check -color green
                        "Finished hiding from email list - SugarLearning Step 7"
                    }
                }
                catch {
                    Set-UDElement -id "icon7" -Content {
                        $LastError = $Error[0]
                        New-UDIcon -icon ban -color red
                        "Error - $LastError"
                    }
                }
            }
            New-UDElement -Id "icon7" -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDColumn -Size 3 -Endpoint {
            New-UDButton -Text "8. Create Exch Rule" -OnClick {
                Set-UDElement -id "icon8" -Content {
                    New-UDIcon -icon check -color green
                    "Finished creating redirect rule in Exchange and sent target an email letting them know we created this rule - SugarLearning Step 8"
                }
                try {
                    $Target = Get-UDElement -Id "targetExch"
                    $Session:TargetExch = $Target.Attributes["value"]
                    New-RedirectRule -username $Session:SelectedUser -target $Session:TargetExch
                    Set-UDElement -id "ico8" -Content {
                        New-UDIcon -icon check -color green
                        "Finished creating redirect rule in Exchange - SugarLearning Step 8"
                    }
                }
                catch {
                    Set-UDElement -id "icon8" -Content {
                        $LastError = $Error[0]
                        New-UDIcon -icon ban -color red
                        "Error - $LastError"
                    }
                }
            }
            New-UDElement -Id "icon8" -tag "span" -Content {
                New-UDIcon -icon ban
            }
            New-UDTextbox -Id "targetExch" -Label "Exchange Rule Target Email Address" -Placeholder "e.g. uly@ssw.com.au"
        }
        New-UDColumn -size 3 -Endpoint {
            New-UDButton -Text "9. Check In Assets" -OnClick {
                Set-UDElement -id "icon9" -Content {
                    New-UDIcon -icon pause
                    "Started checking assets and disabling Snipe user"
                }
                try {
                    CheckIn-SnipeAssets -User $Session:SelectedUser
                    Set-UDElement -id "icon9" -Content {
                        New-UDIcon -icon check -color green
                        "Finished checking assets and disabling Snipe user - SugarLearning Step 9"
                    }
                }
                catch {
                    Set-UDElement -id "icon9" -Content {
                        $LastError = $Error[0]
                        New-UDIcon -icon ban -color red
                        "Error - $LastError"
                    }
                }
            }
            New-UDElement -Id "icon9" -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDColumn -Size 3 -Endpoint {
            New-UDButton -Text "10. Downgrade Zendesk Agent" -OnClick {
                Set-UDElement -id "icon10" -Content {
                    New-UDIcon -icon pause
                    "Started downgrading Zendesk agent to end-user"
                }
                try {
                    Disable-ZendeskUser -User $Session:SelectedUser
                    Set-UDElement -id "icon10" -Content {
                        New-UDIcon -icon check -color green
                        "Finished downgrading Zendesk agent to end-user - SugarLearning Step 10"
                    }
                }
                catch {
                    Set-UDElement -id "icon10" -Content {
                        $LastError = $Error[0]
                        New-UDIcon -icon ban -color red
                        "Error - $LastError"
                    }
                }
            }
            New-UDElement -Id "icon10" -tag "span" -Content {
                New-UDIcon -icon ban
            }
        } 
    }
    
    # Third Row
    New-UDRow -Endpoint {
        New-UDColumn -Size 3 -Endpoint {
            New-UDButton -Text "11. Delete Azure Devops User" -OnClick {
                Set-UDElement -id "icon11" -Content {
                    New-UDIcon -icon pause
                    "Started deleting Azure DevOps users in SSW1 and SSW2"
                }
                try {
                    Disable-AzureDevopsUser -User $Session:SelectedUser
                    Set-UDElement -id "icon11" -Content {
                        New-UDIcon -icon check -color green
                        "Finished deleting Azure DevOps users in SSW1 and SSW2 - SugarLearning Step 11"
                    }
                }
                catch {
                    Set-UDElement -id "icon11" -Content {
                        $LastError = $Error[0]
                        New-UDIcon -icon ban -color red
                        "Error - $LastError"
                    }
                }
            }
            New-UDElement -Id "icon11" -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDColumn -Size 3 -Endpoint {
            New-UDButton -Text "12. Check Azure Resource Groups" -OnClick {
                Set-UDElement -id "icon12" -Content {
                    New-UDIcon -icon pause
                    "Started checking Azure Resource Groups"
                }
                try {
                    $Returned = Search-AzureTags -User $Session:SelectedUser
    
                    $FirstName = $Session:SelectedUser.givenname
                    $Surname = $Session:SelectedUser.surname

                    Set-UDElement -id "icon12" -Content {
                        New-UDIcon -icon check -color green
                        "Found the following Resource Groups owned by $FirstName $Surname (You need to change the Owner tag): $Returned"
                    }
                }
                catch {
                    Set-UDElement -id "icon12" -Content {
                        $LastError = $Error[0]
                        New-UDIcon -icon ban -color red
                        "Error - $LastError"
                    }
                }
            }
            New-UDElement -Id "icon12" -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDColumn -Size 3 -Endpoint {
            New-UDButton -Text "13. Move OU (Last Step)" -OnClick {
                Set-UDElement -id "icon13" -Content {
                    New-UDIcon -icon pause
                    "Started move to disabled users OU"
                }
                try {
                    Move-UserToDisabledUserOU -User $Session:SelectedUser
                    Set-UDElement -id "icon13" -Content {
                        New-UDIcon -icon check -color green
                        "Moved user to ztDisabledUsers_ToClea - - SugarLearning Step 13"
                    }
                }
                catch {
                    Set-UDElement -id "icon13" -Content {
                        $LastError = $Error[0]
                        New-UDIcon -icon ban -color red
                        "Error - $LastError"
                    }
                }
            }
            New-UDElement -Id "icon13" -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDColumn -size 3 -Endpoint {
            New-UDButton -Text "Send Finish Email (After All Steps Done)" -OnClick {
                $EmailToSend = Get-UDElement -id "ownEmail"
                $FinalEMailToSend = $EmailToSend.Attributes["value"]

                Set-UDElement -id "iconemail" -Content {
                    New-UDIcon -icon pause
                    "Started sending finished email to yourself"
                }
                try {
                    Set-UDElement -id "iconemail" -Content {
                        Send-FinishEmail $Session:TargetExch $FinalEMailToSend
                        New-UDIcon -icon check -color green
                        "Finished sending finished email to yourself"
                    }

                }
                catch {
                    Set-UDElement -id "iconemail" -Content {
                        $LastError = $Error[0]
                        New-UDIcon -icon ban -color red
                        "Error - $LastError"
                    }
                }
            }   
            New-UDElement -Id "iconemail" -tag "span" -Content {
                New-UDIcon -icon ban
            }
            New-UDTextbox -Id "ownEmail" -Label "Your Own Email" -Placeholder "e.g. stevenandrews@ssw.com.au"
        }
    }
    

    # Fourth Row
    New-UDRow -Endpoint {
        New-UDColumn -size 3 -Endpoint {
            New-UDButton -Text "Show Session" -OnClick {
                
                $upn = $Session:SelectedUser.UserPrincipalName
                $sam = $Session:SelectedUser.samAccountName
                $fir = $Session:SelectedUser.givenname
                $sur = $Session:SelectedUser.surname
                $mess = "UPN = $upn `n"
                $mess += "SAMAccountName = $sam `n"
                $mess += "First Name = $fir `n"
                $mess += "Surname = $sur"

                Set-UDElement -id "iconsess" -Content {
                    New-UDIcon -icon check -color green
                    "$mess"
                }
            }
            New-UDElement -Id "iconsess" -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDColumn -size 3 -Endpoint {
            New-UDButton -Text "Show Log Location" -OnClick {

                Set-UDElement -id "iconlog" -Content {
                    New-UDIcon -icon check -color green
                    "Clicked - Log can be found in \\fileserver\Backups\SSWLeavingStandard.log"
                }
            }   
            New-UDElement -Id "iconlog" -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
    }
}

$Dashboard = New-UDDashboard -theme $theme -Pages @($Page1) -Title "SSW Leaving Standard"
Start-UDDashboard -port 10000 -Dashboard $Dashboard -Wait
