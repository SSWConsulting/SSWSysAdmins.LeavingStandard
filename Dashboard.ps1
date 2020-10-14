# Importing my other file with sweet, sweet functions
. "$PSScriptRoot\Functions.ps1"

# Importing the configuration file
$config = Import-PowerShellDataFile $PSScriptRoot\Config.PSD1

# Creating variables to determine magic strings and getting them from the configuration file
$LogFile = $config.LogFile
$DCServer = $config.DCServer
$BackupFolders = $config.BackupFolders
$BackupOnPremisesLocation = $config.BackupOnPremisesLocation
$OriginEmail = $config.OriginEmail
$TargetEmail = $config.TargetEmail
$LogModuleLocation = $config.LogModuleLocation
$O365Account = "$PSScriptRoot/$($config.O365Account)"
$O365AccountPass = "$PSScriptRoot/$($config.O365AccountPass)"
$O365AccountKey = "$PSScriptRoot/$($config.O365AccountKey)"
$SMTPServer = $config.SMTPServer
$SnipeKey = "$PSScriptRoot/$($config.SnipeKey)"
$SnipeURL = $config.SnipeURL
$ZendeskUsername = "$PSScriptRoot/$($config.ZendeskUsername)"
$ZendeskToken = "$PSScriptRoot/$($config.ZendeskToken)"
$ZendeskUri1 = $config.ZendeskUri1
$ZendeskUri2 = $config.ZendeskUri2
$AzureDevopsTok1 = "$PSScriptRoot/$($config.AzureDevopsTok1)"
$AzureDevopsTok2 = "$PSScriptRoot/$($config.AzureDevopsTok2)"
$AzureDevopsURI1 = $config.AzureDevopsURI1
$AzureDevopsURI2 = $config.AzureDevopsURI2
$AzureSecretKey = "$PSScriptRoot/$($config.AzureSecretKey)"
$AzureClientSecret = "$PSScriptRoot/$($config.AzureClientSecret)"
$AzureServicePrincipal = $config.AzureServicePrincipal
$AzureTenantId = $config.AzureTenantId
$AzureSubscriptionId = $config.AzureSubscriptionId
$DisabledUsersOU = $config.DisabledUsersOU
$SugarlearningLoginURI = $config.SugarlearningLoginURI
$SugarlearningAccount = $config.SugarlearningAccount
$SugarlearningPassword = "$PSScriptRoot/$($config.SugarlearningPassword)"
$SugarlearningUserURI = $config.SugarlearningUserURI
$SugarlearningDisableUserURI = $config.SugarlearningDisableUserURI
$SharepointIntranetLink1 = $config.SharepointIntranetLink1 
$LeavingStandardGithub = $config.LeavingStandardGithub  
$Office365Powershell = $config.Office365Powershell

# Importing the SSW Write-Log module
Import-Module -Name $LogModuleLocation

$Theme = @{
    palette    = @{
        primary = @{
            main = '#CC4141'
        }    
        grey    = @{
            '300' = '#333333'
        }    
    }
    typography = @{
        fontFamily = "Helvetica"
    }    
}

<#
.SYNOPSIS
This function builds the "Catch" block of all buttons in the page.
    
.DESCRIPTION
This function builds the "Catch" block of all buttons in the page.
Every button uses the same bit of code, so just easier to build a function that does it instead of redoing it every time.
    
.PARAMETER Step
The step variable above that corresponds to the Icon ID.
    
.EXAMPLE
PS> New-CatchErrorBlock -Step $StepNumber
    
.NOTES
Only use this inside this page and in the buttons.
#>
function New-CatchErrorBlock {
    Param(
        [Parameter(Mandatory = $True)]
        $Step
    )
    Clear-UDElement -Id $Step
    Add-UDElement -ParentId $Step -Content {
        $LastError = $Error[0]
        New-UDIcon -icon ban -color red
        "Error - $LastError"
        Write-Log -File $LogFile -Message "Error - $LastError"
    }
}

$Page1 = New-UDPage -Name "LeavingStandard" -Content { 

    # Variables for the IDs below
    $SLStep2 = "icon2"
    $SLStep3 = "icon3"
    $SLStep4and5 = "icon45"
    $SLStep6 = "icon6"
    $SLStep7 = "icon7"
    $SLStep8 = "icon8"
    $SLStep9 = "icon9"
    $SLStep10 = "icon10"
    $SLStep11 = "icon11"
    $SLStep12 = "icon12"
    $SLStep13 = "icon13"
    $SLStep14 = "icon14"
    $SLStep15 = "icon15"
    $SLStepEmail = "iconEmail"
    $SLStepLog = "iconLog"

    # Whole button grid that will be shown at the end of the page
    $WholeGrid = New-UDGrid -Container -Content {
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "2. Backup Search (Start Here)" -OnClick {
                Clear-UDElement -Id $SLStep2
                Add-UDElement -ParentId $SLStep2 -Content {
                    New-UDIcon -icon pause
                    "Started fileserver backups..."
                    Write-Log -File $LogFile -Message "Started fileserver backups..."
                }
                try {
                    New-backupsearch -User $Session:SelectedUser.samAccountName
                    Clear-UDElement -Id $SLStep2
                    Add-UDElement -ParentId $SLStep2 -Content {
                        New-UDIcon -icon check -Color green
                        "Finished fileserver backups - SugarLearning Step 2"
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep2
                }
            }
            New-UDElement -Id $SLStep2 -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "3. Remove Groups" -OnClick {
                Clear-UDElement -Id $SLStep3
                Add-UDElement -ParentId $SLStep3 -Content {
                    New-UDIcon -icon pause
                    "Started removal of user groups..."
                    Write-Log -File $LogFile -Message "Started removal of user groups..."
                }
                try {
                    Remove-UserFromAllADGroups -User $Session:SelectedUser.samAccountName
                    Clear-UDElement -Id $SLStep3
                    Add-UDElement -ParentId  $SLStep3 -Content {
                        New-UDIcon -icon check -color green
                        "Removed user from all AD groups (except Domains Users and any Reporting groups) - SugarLearning Step 3"
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep3
                }
            }
            New-UDElement -Id $SLStep3 -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "4/5. Disable AD User" -OnClick {
                Clear-UDElement -Id $SLStep4and5
                Add-UDElement -ParentId $SLStep4and5 -Content {
                    New-UDIcon -icon pause
                    "Started disabling user..."
                    Write-Log -File $LogFile -Message "Started disabling user..."
                }
                try {
                    Disable-User -User $Session:SelectedUser.samAccountName
                    Clear-UDElement -Id $SLStep4and5
                    Add-UDElement -ParentId $SLStep4and5 -Content {
                        New-UDIcon -icon check -color green
                        "Finished disabling user - SugarLearning Step 4 and 5"
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep4and5
                }
            }
            New-UDElement -Id $SLStep4and5 -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "6. Disable Admin AD User" -OnClick {
                Clear-UDElement -Id $SLStep6
                Add-UDElement -ParentId $SLStep6 -Content {
                    New-UDIcon -icon pause
                    "Started disabling user..."
                    Write-Log -File $LogFile -Message "Started disabling user..."
                }
                try {
                    $ReturnedAdmin = Disable-AdminUser -User $Session:SelectedUser.UserPrincipalName
                    Clear-UDElement -Id $SLStep6
                    Add-UDElement -ParentId $SLStep6 -Content {
                        New-UDIcon -icon check -color green
                        if ($ReturnedAdmin -ne $null) {
                            "Found and disabled the following Admin user: $ReturnedAdmin - SugarLearning Step 6"
                        }
                        else {
                            "Not found any admin users for $($Session:SelectedUser.UserPrincipalName)"
                        }
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep6
                }
            }
            New-UDElement -Id $SLStep6 -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "7. Remove AD Attributes" -OnClick {
                Clear-UDElement -Id $SLStep7
                Add-UDElement -ParentId $SLStep7 -Content {
                    New-UDIcon -icon pause
                    "Started clearing Manager and extensionAttribute1 fields in AD..."
                    Write-Log -File $LogFile -Message "Started clearing Manager and extensionAttribute1 fields in AD..."
                }
                try {
                    Remove-ADAttributes -User $Session:SelectedUser.samAccountName
                    Clear-UDElement -Id $SLStep7
                    Add-UDElement -ParentId $SLStep7 -Content {
                        New-UDIcon -icon check -color green
                        "Finished clearing Manager and extensionAttribute1 fields in AD - SugarLearning Step 7"
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep7
                }
            }
            New-UDElement -Id $SLStep7 -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "8. Hide from Email List" -OnClick {
                Clear-UDElement -Id $SLStep8
                Add-UDElement -ParentId $SLStep8 -Content {
                    New-UDIcon -icon pause
                    "Started hiding from email list..."
                    Write-Log -File $LogFile -Message "Started hiding from email list..."
                }
                try {
                    Hide-Email -User $Session:SelectedUser.samAccountName
                    Clear-UDElement -Id $SLStep8
                    Add-UDElement -ParentId $SLStep8 -Content {
                        New-UDIcon -icon check -color green
                        "Finished hiding from email list - SugarLearning Step 8"
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep8
                }
            }
            New-UDElement -Id $SLStep8 -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "9. Create Exchange Rule" -OnClick {
                Clear-UDElement -Id $SLStep9
                Add-UDElement -ParentId $SLStep9 -Content {
                    New-UDIcon -icon pause
                    "Started creating redirect rule in Exchange..."
                    Write-Log -File $LogFile -Message "Started creating redirect rule in Exchange..."
                }
                try {
                    $Session:TargetExch = (Get-UDElement -Id "targetExch").value
                    New-RedirectRule -User $Session:SelectedUser -Target $Session:TargetExch
                    Clear-UDElement -Id $SLStep9
                    Add-UDElement -ParentId $SLStep9 -Content {
                        New-UDIcon -icon check -color green
                        "Finished creating redirect rule in Exchange - SugarLearning Step 9"
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep9
                }
            }
            New-UDTextbox -Id "targetExch" -Label "Exchange Rule Target Email Address" -Placeholder "e.g. uly@ssw.com.au"
            New-UDElement -Id $SLStep9 -tag "span" -Content {
                New-UDIcon -icon ban
            }            
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "10. Check In Assets" -OnClick {
                Clear-UDElement -Id $SLStep10
                Add-UDElement -ParentId $SLStep10 -Content {
                    New-UDIcon -icon pause
                    "Started checking assets and disabling Snipe user..."
                    Write-Log -File $LogFile -Message "Started checking assets and disabling Snipe user..."
                }
                try {
                    CheckIn-SnipeAssets -User $Session:SelectedUser
                    Clear-UDElement -Id $SLStep10
                    Add-UDElement -ParentId $SLStep10 -Content {
                        New-UDIcon -icon check -color green
                        "Finished checking assets and disabling Snipe user - SugarLearning Step 10"
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep10
                }
            }
            New-UDElement -Id $SLStep10 -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "11. Downgrade Zendesk Agent" -OnClick {
                Clear-UDElement -Id $SLStep11
                Add-UDElement -ParentId $SLStep11 -Content {
                    New-UDIcon -icon pause
                    "Started downgrading Zendesk agent to end-user..."
                    Write-Log -File $LogFile -Message "Started downgrading Zendesk agent to end-user..."
                }
                try {
                    Disable-ZendeskUser -User $Session:SelectedUser
                    Clear-UDElement -Id $SLStep11
                    Add-UDElement -ParentId $SLStep11 -Content {
                        New-UDIcon -icon check -color green
                        "Finished downgrading Zendesk agent to end-user - SugarLearning Step 11"
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep11
                }
            }
            New-UDElement -Id $SLStep11 -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "12. Delete Azure DevOps User" -OnClick {
                Clear-UDElement -Id $SLStep12
                Add-UDElement -ParentId $SLStep12 -Content {
                    New-UDIcon -icon pause
                    "Started deleting Azure DevOps users in SSW1 and SSW2..."
                    Write-Log -File $LogFile -Message "Started deleting Azure DevOps users in SSW1 and SSW2..."
                }
                try {
                    Disable-AzureDevopsUser -User $Session:SelectedUser
                    Clear-UDElement -Id $SLStep12
                    Add-UDElement -ParentId $SLStep12 -Content {
                        New-UDIcon -icon check -color green
                        "Finished deleting Azure DevOps users in SSW1 and SSW2 - SugarLearning Step 12"
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep12
                }
            }
            New-UDElement -Id $SLStep12 -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "13. Check Azure Resource Groups" -OnClick {
                Clear-UDElement -Id $SLStep13
                Add-UDElement -ParentId $SLStep13 -Content {
                    New-UDIcon -icon pause
                    "Started checking Azure Resource Groups..."
                    Write-Log -File $LogFile -Message "Started checking Azure Resource Groups..."
                }
                try {
                    $Returned = Search-AzureTags -User $Session:SelectedUser
    
                    $FirstName = $Session:SelectedUser.givenname
                    $Surname = $Session:SelectedUser.surname

                    Clear-UDElement -Id $SLStep13
                    Add-UDElement -ParentId $SLStep13 -Content {
                        New-UDIcon -icon check -color green
                        if ($Returned -ne $null) {
                            "Found the following Resource Groups owned by $FirstName $Surname (You need to change the Owner tag): $Returned"
                        }
                        else {
                            "Not found any Azure Resource Groups owned by $FirstName $Surname"
                        }
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep13
                }
            }
            New-UDElement -Id $SLStep13 -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "14. Move OU (Last Step)" -OnClick {
                Clear-UDElement -Id $SLStep14
                Add-UDElement -ParentId $SLStep14 -Content {
                    New-UDIcon -icon pause
                    "Started move to Disabled users OU..."
                    Write-Log -File $LogFile -Message "Started move to Disabled users OU..."
                }
                try {
                    Move-UserToDisabledUserOU -User $Session:SelectedUser.ObjectGUID
                    Clear-UDElement -Id $SLStep14
                    Add-UDElement -ParentId $SLStep14 -Content {
                        New-UDIcon -icon check -color green
                        "Moved user to DisabledUsers - SugarLearning Step 14"
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep14
                }  
            }              
            New-UDElement -Id $SLStep14 -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "15. Disable Sugarlearning Account" -OnClick {
                Clear-UDElement -Id $SLStep15
                Add-UDElement -ParentId $SLStep15 -Content {
                    New-UDIcon -icon pause
                    "Started disabling Sugarlearning's account..."
                    Write-Log -File $LogFile -Message "Started disabling Sugarlearning's account..."
                }
                try {
                    Disable-SugarlearningUser -User $Session:SelectedUser.userprincipalname
                    Clear-UDElement -Id $SLStep15
                    Add-UDElement -ParentId $SLStep15 -Content {
                        New-UDIcon -icon check -color green
                        "Disabled Sugarlearning user - SugarLearning Step 15"
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStep15
                }  
            }              
            New-UDElement -Id $SLStep15 -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "Send Finish Email (After All Steps Done)" -OnClick {
                Clear-UDElement -Id $SLStepEmail
                Add-UDElement -ParentId $SLStepEmail -Content {
                    New-UDIcon -icon pause
                    "Started sending finished email to yourself..."
                    Write-Log -File $LogFile -Message "Started sending finished email to yourself..."
                }
                try {
                    $Session:FinalEmailToSend = (Get-UDElement -Id "targetEmail").value
                    Send-FinishEmail $Session:TargetExch $Session:FinalEmailToSend $Session:SelectedUser.Name
                    Clear-UDElement -Id $SLStepEmail
                    Add-UDElement -ParentId $SLStepEmail -Content {
                        New-UDIcon -icon check -color green
                        "Finished sending finished email to yourself..."
                    }
                }
                catch {
                    New-CatchErrorBlock -Step $SLStepEmail
                }  
            }       
            New-UDTextbox -Id "targetEmail" -Label "Send Email To" -Placeholder "e.g. yourself@ssw.com.au"       
            New-UDElement -Id $SLStepEmail -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
        New-UDGrid -Item -ExtraSmallSize 3 -Content {
            New-UDButton -Text "SHow Log Location" -OnClick {
                Clear-UDElement -Id $SLStepLog
                Add-UDElement -ParentId $SLStepLog -Content {
                    New-UDIcon -icon check -color green
                    "Log can be found at $LogFile"
                } 
            }            
            New-UDElement -Id $SLStepLog -tag "span" -Content {
                New-UDIcon -icon ban
            }
        }
    }

    $SearchButton = New-UDRow -Columns {
        New-UDColumn -Size 6 -Content {
            New-UDTextbox -Id "txtSearch" -Label "Search" -Placeholder "e.g. adamcogan"
        }
        New-UDColumn -Size 2 -Content {
            New-UDButton -Id "btnSearch" -Text "Search User" -OnClick {
                $Session:Value = (Get-UDElement -id "txtSearch").value
                $Session:Objects = Get-ADUser -Filter "Name -like '*$Session:Value*' -or samAccountName -like '*$Session:Value*' -or UserPrincipalName -like '*$Session:Value*@ssw.com.au'" -Server $DCServer
                Sync-UDElement -id "Dynamic1"                
                Start-Sleep -s 1
                Invoke-UDJavaScript -JavaScript 'document.getElementById("Table1").parentElement.style.display = "block";'      
            }
        }   
    }  
    New-UDStepper -Steps {
        New-UDStep -OnLoad {
            New-UDTypography -Variant "h3" -Text "Welcome to the SSW Leaving Standard!" -GutterBottom
            New-UDElement -Tag 'p' -Content {}
            New-UDHtml -markup "<span style='font-family:Helvetica'><h3>Are you following the 'Disable your accounts' SugarLearning item? <a href=https://my.sugarlearning.com/companies/SSW/items/13045/Disable-your-Accounts target='_blank'>https://my.sugarlearning.com/companies/SSW/items/13045/Disable-your-Accounts</a></h3>" 
            New-UDCheckbox -Id 'checkStep1' -Label "Yes, I am following this item." -LabelPlacement end
        } -Label "Step 1 - Following Sugarlearning"
        New-UDStep -OnLoad {
            New-UDTypography -Variant "h4" -Text "Before proceeding, it is necessary to backup the employee's mailbox to our on-premises file server." -GutterBottom
            New-UDHtml -Markup "<span style='font-family:Helvetica'><h3>Go to <a href=https://protection.office.com/contentsearchbeta target='_blank'>https://protection.office.com/contentsearchbeta</a>:<ol><li>New Search | Specific Locations | Modify... | Choose users, groups or teams... | Search for the email | Export Results...</li><li>Download the mailbox to \\fileserver\DataSSW\ExEmployees\ExchangeBackup\{username}.pst</li></ol></h3>"
            New-UDCheckbox -Id 'checkStep2' -Label "Yes, I have backed up the email." -LabelPlacement end
        } -Label "Step 2 - Backing up Email"
    } -OnFinish {
        $Context = ConvertFrom-Json $Body
        New-UDTypography -Text "Search the leaving employee in AD below:" -Variant h3
        Add-UDElement -ParentId "Search1" -Content { $SearchButton }
        Add-UDElement -ParentId "Grid1" -Content { $WholeGrid }
    } -OnValidateStep {
        $Context = ConvertFrom-Json $Body
        if ($Context.CurrentStep -eq 0 -and $Context.Context.checkStep1 -eq $false) {
            New-UDValidationResult -ValidationError "You need to make sure you are following the Sugarlearing item!"
        }
        elseif ($Context.CurrentStep -eq 1 -and $Context.Context.checkStep2 -eq $false) {
            New-UDValidationResult -ValidationError "You need to make sure you have backed up the user's email first!" 
        }
        else {
            New-UDValidationResult -Valid 
        }
    }
  
    New-UDElement -tag "div" -Id "Search1" -Content {
    }
    
    New-UDDynamic -id "Dynamic1" -content {  
        New-UDElement -Tag "div" -Id "Style1" -Attributes @{
            style = @{
                display = 'none'
            }
        } -Content {        
            $Columns = @(
                New-UDTableColumn -Property ObjectGUID -Title Select -Render { 
                    $Item = $Body | ConvertFrom-Json 
                    New-UDButton -Text "Select" -OnClick { 
                        $Session:SelectedUser = $Item
                        Show-UDToast -Message "Selected Name: $($Session:SelectedUser.Name), SAMAccountName: $($Session:SelectedUser.SAMAccountName), UserPrincipalName: $($Session:SelectedUser.UserPrincipalName), ObjectGUID: $($Session:SelectedUser.ObjectGUID)" 
                    }          
                }
                New-UDTableColumn -Property Name -Title Name
                New-UDTableColumn -Property GivenName -Title GivenName
                New-UDTableColumn -Property Surname -Title Surname
                New-UDTableColumn -Property SAMAccountName -Title SAMAccountName 
                New-UDTableColumn -Property UserPrincipalName -Title UserPrincipalName
                New-UDTableColumn -Property ObjectGUID -Title GUID
            ) 
            New-UDTable -ID "Table1" -Columns $Columns -LoadData {
                $Query = $Body | ConvertFrom-Json
                $Session:Objects | Out-UDTableData -Page 0 -TotalCount 5 -Properties $Query.Properties 
            }
        }    
    }   
    New-UDElement -tag "div" -Id "Grid1" -Content {
    }
}

$Dashboard = New-UDDashboard -Theme $Theme -Pages @($Page1) -Title "SSW Leaving Standard"
$Dashboard