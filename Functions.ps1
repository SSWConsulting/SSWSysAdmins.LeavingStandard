<#
.SYNOPSIS
   All functions that will be used by the SSW Leaving Standard.
.DESCRIPTION
    All functions that will be used by the SSW Leaving Standard.
    It contains all functions that will be called by the main dashboard script.
.EXAMPLE
    This script is imported in the main dashboard script and its functions are called as necessary by the buttons on the screen.
.INPUTS
    Configuration file: Config.psd1 (called from the main dashboard script)
.OUTPUTS
    Based on each function.
.NOTES
    Only use this script file with the main dashboard file, they both use the same set of variables and session variables, acting like one big file.
    Dependencies, you will need to install all these in the new server if migrating:
    - 7zip CLI: for New-BackupSearch - needs to be installed and available in the "Path" system environment variable.
    - ExchangeOnline Module: for New-RedirectRule - Install-Module -Name ExchangeOnlineManagement -RequiredVersion 1.0.1
    - Snipe Module for Powershell (https://www.powershellgallery.com/packages/SnipeitPS): for CheckIn-SnipeAssets - Install-Module -Name SnipeitPS
    - Azure CLI (https://docs.microsoft.com/en-us/cli/azure/install-azure-cli-windows?view=azure-cli-latest&tabs=azure-cli): for Disable-AzureDevopsUser - Invoke-WebRequest -Uri https://aka.ms/installazurecliwindows -OutFile .\AzureCLI.msi; Start-Process msiexec.exe -Wait -ArgumentList '/I AzureCLI.msi /quiet'; rm .\AzureCLI.msi
    - PowerShell Az module (https://docs.microsoft.com/en-us/powershell/azure/new-azureps-module-az): for Search-AzureTags  - Install-Module -Name Az

    Created by Kaique "Kiki" Biancatti for SSW.
#>

<#
.SYNOPSIS
Searches the file share for files and compresses them to another location.

.DESCRIPTION
Searches the file share for files and compresses them to another location.
Uses 7zip CLI and -aoa switch, always overwriting everything.

.PARAMETER username
The username to search in folder locations.

.EXAMPLE
PS> New-BackupSearch -User $Session:SelectedUser.samAccountName

.NOTES
7zip CLI needs to be installed and available in the "Path" system environment variable for this to work. 
#>
function New-BackupSearch {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )
   
    $FullUser = "Data" + $username

    $BackupFolders | ForEach-Object {
        $UserBackupFiles = Get-ChildItem $_ -Filter $FullUser
        if ($UserBackupFiles -eq $null) {
            Write-Log -File $LogFile -Message "Found nothing in $_"
        }
        else {
            Write-Log -File $LogFile -Message "Folder found in $_$FullUser. Backing up..."
            Write-Log -File $LogFile -Message "Backup location: $BackupOnPremisesLocation$fulluser.zip"
            $Test = Start-Job -ScriptBlock {               
                $FullUser = $args[0]
                $BackupOnPremisesLocation = $args[1]
                $Current = $args[2]
                $UserBackupFiles = $args[3]
                $LogFile = $args[4]
                $LogModuleLocation = $args[5]

                Import-Module -Name $LogModuleLocation
                Write-Log -File $LogFile -Message "From: $Current$UserBackupFiles\"
                Write-Log -File $LogFile -Message "To: $BackupOnPremisesLocation$FullUser.zip"
                
                7z a -aoa -bsp1 $BackupOnPremisesLocation$FullUser".zip" $Current$UserBackupFiles"\" | out-string -Stream | Select-String -Pattern "\d{1,3}%"
            } -ArgumentList $FullUser, $BackupOnPremisesLocation, $_, $UserBackupFiles, $LogFile, $LogModuleLocation
            
            While (Get-Job -State "Running") {
                $results = Receive-Job -Keep -Job $Test | Select -Last 1
                Clear-UDElement -Id "icon2"
                Add-UDElement -ParentId "icon2" -Content {
                    New-UDProgress -Circular -Color Blue -Size Small 
                    "Running fileserver backups: $results"
                }
                Write-Log -File $LogFile -Message "Running fileserver backups:$results"
                Start-sleep -s 3                
            }            
            Write-Log -File $LogFile -Message "Finished fileserver backups from $BackupOnPremisesLocation$FullUser.zip"
        }

    }
}

<#
.SYNOPSIS
Removes a user from all AD groups.

.DESCRIPTION
Removes a user from all AD groups.
Removes a user from all groups he is member of, with the exception of 'Domain Users' (all users need to be part of at least one group) and any Reporting groups (Reporting groups are/were used by CRM).

.PARAMETER username
The Samaccountname of the user.

.EXAMPLE
PS> Remove-UserFromAllADGroups -User $Session:SelectedUser.samAccountName
#>
function Remove-UserFromAllADGroups {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )
    $FoundGroups = Get-ADprincipalGroupMembership "$username" | Where-Object Name -notlike '*Report*' | where-object Name -notlike 'Domain Users'
    if ($FoundGroups -eq $null) {
        Write-Log -File $LogFile -Message "Found no groups where $username is part of (except Domain Users and any Reporting groups)..."
    }
    else {
        $FoundGroups | ForEach-Object {
            try {
                Remove-ADGroupMember -Identity $_ -Member $username -Confirm:$false
                Write-Log -File $LogFile -Message "Successfully removed $username from AD group $($_.Name)"
            } 
            catch {
                $LastError = $Error[0]
                Write-Log -File $LogFile -Message "Failed to remove $username from AD group $($_.Name) - $LastError"
            }        
        }
        Write-Log -File $LogFile -Message "Removed all groups from $Username..."
    }    
}

<#
.SYNOPSIS
Clears the Manager and extensionAttribute1 fields in AD.

.DESCRIPTION
Clears the Manager and extensionAttribute1 fields in AD.
This needs to be cleared so the user will not show in Delve after he leaves, or will be contacted in his personal email address in other SSW automations.

.PARAMETER username
The Samaccountname of the user.

.EXAMPLE
PS> Remove-ADAttributes -User $Session:SelectedUser.samAccountName
#>
function Remove-ADAttributes {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )
    Set-ADUser -Identity "$username" -Clear Manager
    Write-Log -File $LogFile -Message "Cleared $username Manager field in AD..."
    Set-ADUser -Identity "$username" -Clear extensionAttribute1
    Write-Log -File $LogFile -Message "Cleared $username extensionAttribute1 (personal email) field in AD..."
}

<#
.SYNOPSIS
Moves the user to a OU.

.DESCRIPTION
Moves the user to a OU.
Uses the user GUID to move it somewhere else.

.PARAMETER username
The objectGUID of the user.

.EXAMPLE
PS> Move-UserToDisabledUserOU -User $Session:SelectedUser.ObjectGUID
#>
function Move-UserToDisabledUserOU {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )
    Move-ADObject -Identity "$username" -TargetPath $DisabledUsersOU -Server $DCServer
    Write-Log -File $LogFile -Message "Moved to $DisabledUsersOU..."
}

<#
.SYNOPSIS
Disables the user in AD and sets its description to the current date.

.DESCRIPTION
Disables the user in AD and sets its description to the current date.
Also writes in the log what is happening.

.PARAMETER username
The Samaccountname of the user.

.EXAMPLE
PS> Disable-User -User $Session:SelectedUser.samAccountName
#>
function Disable-User {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )
    Disable-ADAccount -Identity "$username"
    Write-Log -File $LogFile -Message "Disabled $username..."
    $date = Get-Date -Format "dd/MM/yyyy"
    Set-ADUser -Identity "$username" -Description "Disabled - $date"
    Write-Log -File $LogFile -Message "Set $username's description to Disabled - $date..."
}

<#
.SYNOPSIS
Hides the user's email address from Exchange address lists.

.DESCRIPTION
Hides the user's email address from Exchange address lists.

.PARAMETER username
The Samaccountname of the user.

.EXAMPLE
PS> Hide-Email -User $Session:SelectedUser.samAccountName
#>
function Hide-Email {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )
    Set-ADUser -Identity "$username" -Replace @{ msexchhidefromaddresslists = $true }
    Write-Log -File $LogFile -Message "Hid $username's email from address list..."
}

<#
.SYNOPSIS
Creates a redirect rule in Exchange Online.

.DESCRIPTION
Creates a redirect rule in Exchange Online.
It connects to Exchange Online, uses New-TransportRule cmdlet to redirect emails from the leaving employee to someone else.
The nominated receiver will then receive an email saying that they will receive said emails.

.PARAMETER username
The user object that is leaving the company and therefore will have its emails redirected to someone else.

.PARAMETER target
The target of the redirection, the one who will receive the emails.

.EXAMPLE
PS> New-RedirectRule -User $Session:SelectedUser -Target $Session:TargetExch

.NOTES
Needs ExchangeOnline Module - Install-Module -Name ExchangeOnlineManagement -RequiredVersion 1.0.1
#>
function New-RedirectRule {
    Param(
        [Parameter(Mandatory = $True, Position = 0)]
        $username,
        [Parameter(Mandatory = $True, Position = 1)]
        $target
    )
    try {
        $FunctionAccount = Get-Content $O365Account
        $FunctionPassword = Get-Content $O365AccountPass | ConvertTo-SecureString -Key (Get-Content $O365AccountKey)
        $O365Credential = New-Object System.Management.Automation.PsCredential($FunctionAccount, $FunctionPassword )

        Connect-ExchangeOnline -ConnectionUri $Office365Powershell -Credential $O365Credential

        $Email = $username.userPrincipalName
        $FirstName = $username.GivenName
        $LastName = $username.Surname
        Write-Log -File $LogFile -Message "Connected to Exchange Online..."
    }
    catch {
        $LastError = $Error[0]
        Write-Log -File $LogFile -Message "Failed to connect to Exchange Online - $LastError"
    }    
    try {
        $ChangeDate = (Get-date).AddMonths(6)
        $ChangeDate = $ChangeDate.ToString("dd/MM/yy")

        New-TransportRule "Leaving Standard - $FirstName $LastName - autoreply on $ChangeDate" -SentTo $Email -RedirectMessageTo $target

        $bodyhtml = @"
        <div style='font-family:Helvetica;'>
        <p>We just added a Redirect Rule in our Exchange Online Server, as per Step 8 on the <a href=https://my.sugarlearning.com/SSW/items/13045/disable-your-accounts target='_blank'>SSW Leaving Standard</a>.</p>
        <p>You will now receive all emails from $Email (this means $FirstName $LastName left SSW). <br>On $ChangeDate, this redirect rule will be changed to an auto-reply stating that $FirstName $LastName left SSW for good, and you will stop receiving the emails.</p>
        
        If you think this is wrong, please contact SysAdmins.<br>
        Tip: You can find a log file with more information at <a href=$LogFile> $LogFile </a>
        <p></p>
        <p>-- Powered by SSWSysAdmins.LeavingStandard<br> 
        GitHub: <a href=$LeavingStandardGithub>SSWSysAdmins.LeavingStandard</a><br> 
        Server: $env:computername 
        </p>
"@
        
        Send-MailMessage -From $OriginEmail -to $target -Subject "SSW.LeavingStandard - Mails are now being redirected from $Email to you" -Body $bodyhtml -SmtpServer $SMTPServer -BodyAsHtml
        Write-Log -File $LogFile -Message "Created redirect rule in exchange..."
    }
    catch {
        $LastError = $Error[0]
        Write-Log -File $LogFile -Message "Failed to create Transport Rule on Exchange Online - $LastError"
    }    
}

<#
.SYNOPSIS
Resets the owner of assets that are checked-in for the user.

.DESCRIPTION
Resets the owner of assets that are checked-in for the user.
Connects to our Snipe-IT instance and searches for all the assets that are assigned to a particular user, then checks it in.

.PARAMETER username
The Samaccountname of the user.

.EXAMPLE
PS> CheckIn-SnipeAssets -User $Session:SelectedUser

.NOTES
Needs the Snipe Module for Powershell (https://www.powershellgallery.com/packages/SnipeitPS) - Install-Module -Name SnipeitPS
#>
function CheckIn-SnipeAssets {
    Param(
        [Parameter(Mandatory = $True, Position = 0)]
        $username
    )
    $SnipeApiKey = get-content $SnipeKey
    $SnipeUser = $username.samaccountname

    # Here we load the Snipe Module for Powershell (https://www.powershellgallery.com/packages/SnipeitPS) to easily access and use it
    Import-Module SnipeITPS
    Set-Info -URL $SnipeURL -apiKey $SnipeApiKey
    
    $UserAssets = Get-Asset -search "$SnipeUser" | where { $_.assigned_to.username -eq "$SnipeUser" } 
    if ($UserAssets -eq $null) {
        Write-Log -File $LogFile -Message "No assets to check-in in Snipe..."
    }
    else {
        $UserAssets | foreach { Write-Log -File $LogFile -Message "Checked in asset ID $($_.id)/Asset tag $($_.asset_tag) in Snipe..." }
        $UserAssets | Foreach { Reset-AssetOwner -Id $_.id }
    }    
}

<#
.SYNOPSIS
Turns a Zendesk agent to end-user.

.DESCRIPTION
Turns a Zendesk agent to end-user.

.PARAMETER username
The user object.

.EXAMPLE
PS> Disable-ZendeskUser -User $Session:SelectedUser
#>
function Disable-ZendeskUser {
    Param(
        [Parameter(Mandatory = $True, Position = 0)]
        $username
    )
    $LocalUsername = $username.userPrincipalName

    $ZendeskUser = get-content $ZendeskUsername
    $token = get-content $ZendeskToken

    $params = @{
        Uri     = $ZendeskUri1
        Headers = @{ Accept = 'application/json'
            'Authorization' = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($ZendeskUser):$($Token)")); 
        }
        Method  = 'GET'
    }
    $ZendeskUsers = Invoke-RestMethod @params
    $ZendeskEmployee = $ZendeskUsers.users | where email -eq $LocalUsername
    $ZendeskEmployeeEmail = $ZendeskEmployee.email

    #Build the json payload to send to Zendesk
    $JsonBody = @{
        user = @{
            email = $ZendeskEmployeeEmail
            role  = "end-user"
        }
    }

    #Have to do all these conversions for it to work in Zendesk
    $JsonBody2 = $JsonBody | ConvertTo-Json -Compress -Depth 4
    $JsonBody2 = $JsonBody2.Replace("\u0026", "&").Replace("`r", "").Replace("`n", "")

    $params2 = @{
        Uri         = $ZendeskUri2
        Headers     = @{ Accept = 'application/json'
            'Authorization' = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($ZendeskUser):$($Token)")); 
        }
        Body        = $JsonBody2
        ContentType = 'application/json'
    }
    try {
        Invoke-RestMethod -Method "POST" @params2
        Write-Log -File $LogFile -Message "Successfully downgraded agent to end-user in Zendesk..."
    }
    catch {
        Write-Log -File $LogFile -Message "User not found or other error - $($_.Exception.Message)"
    }  
}

<#
.SYNOPSIS
Disables and removes the user from SSW1 and SSW2 Azure DevOps.

.DESCRIPTION
Disables and removes the user from SSW1 and SSW2 Azure DevOps.
Uses Azure CLI to do the commands instead of Graph API, easier and work better and faster.

.PARAMETER username
The Samaccountname of the user.

.EXAMPLE
PS> Disable-AzureDevopsUser -User $Session:SelectedUser

.NOTES
Needs Azure CLI (https://docs.microsoft.com/en-us/cli/azure/install-azure-cli-windows?view=azure-cli-latest&tabs=azure-cli): for Disable-AzureDevopsUser - Invoke-WebRequest -Uri https://aka.ms/installazurecliwindows -OutFile .\AzureCLI.msi; Start-Process msiexec.exe -Wait -ArgumentList '/I AzureCLI.msi /quiet'; rm .\AzureCLI.msi

#>
function Disable-AzureDevopsUser {
    Param(
        [Parameter(Mandatory = $True, Position = 0)]
        $username
    )
    $AzureDevopsusername = $username.userPrincipalName

    # SSW1
    try {
        # Install without prompt
        az config set extension.use_dynamic_install=yes_without_prompt
        # Login with token to az devops
        Get-Content $AzureDevopsTok1 | az devops login --org $AzureDevopsURI1
        az devops user remove --org $AzureDevopsURI1 --user "$AzureDevopsusername" --yes
        Write-Log -File $LogFile -Message "Deleted $AzureDevopsusername from $AzureDevopsURI1..."
    }
    catch {
        $LastError = $Error[0]
        Write-Log -File $LogFile -Message "Error - $LastError"
    }
    
    # SSW2
    try {
        az config set extension.use_dynamic_install=yes_without_prompt
        Get-Content $AzureDevopsTok2 | az devops login --org $AzureDevopsURI2
        az devops user remove --org $AzureDevopsURI2 --user "$AzureDevopsusername" --yes
        Write-Log -File $LogFile -Message "Deleted $AzureDevopsusername from $AzureDevopsURI2..."
    }
    catch {
        $LastError = $Error[0]
        Write-Log -File $LogFile -Message "Error - $LastError"
    }   
}


<#
.SYNOPSIS
Sends an email with steps taken on the Leaving Standard procedure.

.DESCRIPTION
Sends an email with steps taken on the Leaving Standard procedure.
Steps are based on SSW's Leaving Standard.

.PARAMETER primaryTarget
The "exchange rule" (step 8) email target.

.PARAMETER ownTarget
The target that this email will be sent to.

.PARAMETER nameTarget
The name of the user.

.EXAMPLE
PS> Send-FinishEmail $Session:TargetExch $Session:FinalEmailToSend $Session:SelectedUser.Name
#>
function Send-FinishEmail {
    Param(
        [Parameter(Mandatory = $False, Position = 0)]
        $primaryTarget,
        [Parameter(Mandatory = $True, Position = 1)]
        $ownTarget,
        [Parameter(Mandatory = $True, Position = 2)]
        $nameTarget
    )
    if ($primaryTarget -eq $null ) {
        $primaryTarget = "XXX (whoever you think is best)"
    }

    $bodyhtml2 = @"
    <div style='font-family:Helvetica;'>
    <p>SSW Leaving Standard finished for $nameTarget</p>
    <ul><li>1. Email - Backup my email (in Office 365 Compliance Center) to fileserver</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>2. File Server - Go through this guide <a href=$SharepointIntranetLink1>here<a> in our intranet and totally clean my folders in fileserver </li></ul>
    <p><strong>Done (automatically)</strong></p>
    <ul><li>3. Active Directory - Remove me from groups in Active Directory (Leave CRM and default groups)</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>4. Active Directory - Disable my account</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>5. Active Directory - Edit my account description with 'Disabled - &lt; The Current Date&gt; '</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>6. Active Directory - Disable my Admin account, if any</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>7. Active Directory - Remove ExtensionAttribute1 and Manager field from AD</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>8. Exchange - Hide my email from address lists</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>9. Exchange - Forward my emails to $primaryTarget without leaving a copy on the recipients mailbox. Email the new owner to let him know. Also, add a Mail Flow rule in our EAC to forward my emails to the person above.</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>10. Snipe - Check in all assets checked out to me on Snipe (Go to <a href=$SnipeURL>$SnipeURL</a> | Look for my name on the asset list | Check in all assets so we know they are available for other people)</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>11. Zendesk - Turn agent to end-user in Zendesk</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>12. Azure DevOps - Remove my access from ssw.visualstudio.com and ssw2.visualstudio.com</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>13. Azure - Check if any Azure Resource Groups are still owned by me - if yes, they need to be handed over to someone else</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>14. Active Directory - Move my account to SSW/ztDisabledUsers</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>15. SugarLearning - Remove my access from SugarLearning</li></ul>
    <p><strong>Done (automatically)</strong></p><ul>
    <li>16. Github - Remove my access from SSW Consulting GitHub</li></ul>
    <ul><li>17. Microsoft Partner Center - Remove my MSDN subscription</li></ul>
    <ul><li>18. Invixium - Remove my fingerprint from Control 4 [Sydney Office Only]</li></ul>
    <ul><li>19. Control 4 - Remove Control 4 Account: https://customer.control4.com/account/users</li></ul>
    <ul><li>20. OneDrive - Go into my OneDrive and backup the important files in there, they will be lost after 30 days (Office 365 Admin | Active Users | Select the user | OneDrive tab | Get access to files | Download important files to fileserver) </li></ul> 
    <ul><li>21. CRM - Input the correct date in my user CRM account in the field 'Date Finished'</li></ul>
    <ul><li>22. CRM - Disable my Dynamics 365 (CRM) User account</li></ul>
    <p>Note: Thank you, hopefully we will see you around the user groups!</p>
    <p>-- Partially powered by SSW.LeavingStandard<br> 
    <br>GitHub: <a href=$LeavingStandardGithub>SSWSysAdmins.LeavingStandard</a><br>
    Server: $env:computername <br>
    </p></div>
"@
    Send-MailMessage -From $OriginEmail -to $ownTarget -Subject "Leaving SSW - Disable accounts" -Body $bodyhtml2 -SmtpServer $SMTPServer -BodyAsHtml
}

<#
.SYNOPSIS
Searches Azure Resource Groups for the Owner tag.

.DESCRIPTION
Searches Azure Resource Groups for the Owner tag.
If the Owner tag in the RGs match, it will be shown in the UI.

.PARAMETER username
The user object.

.EXAMPLE
PS> Search-AzureTags -User $Session:SelectedUser

.NOTES
Needs PowerShell Az module (https://docs.microsoft.com/en-us/powershell/azure/new-azureps-module-az): - Install-Module -Name Az
#>
function Search-AzureTags {
    Param(
        [Parameter(Mandatory = $True, Position = 0)]
        $username
    )
    $UserFirstName = $username.GivenName
    $UserSurname = $username.Surname

    try {
        Import-Module Az
        Write-Log -File $LogFile -Message "Imported Module Az..."
    }
    catch {
        $LastError = $Error[0]
        Write-Log -File $LogFile -Message "Failed to import Module Az - $LastError"
    }
    
    $AesKEy = get-content $AzureSecretKey
    $AzureUser = $AzureServicePrincipal
    $AzureSecret = get-content $AzureClientSecret | ConvertTo-SecureString -Key $AesKey
    $AzureCredential = New-Object System.Management.Automation.PSCredential($AzureUser, $AzureSecret)
    
    $Session:FinalAzureGroups = @()
    try {
        $Connected = Connect-AzAccount -Credential $AzureCredential -TenantId $AzureTenantId -SubscriptionId $AzureSubscriptionId -ServicePrincipal
        Write-Log -File $LogFile -Message "Connected to Azure Subscription..."
    }
    catch {
        $LastError = $Error[0]
        Write-Log -File $LogFile -Message "Failed to connect to Azure Subscription - $LastError"
    }
    
    $AzureGroups = Get-AzResourceGroup -Tag @{owner = "$UserFirstName $UserSurname" } | ForEach { $Session:FinalAzureGroups += "" + $_.ResourceGroupName + " " }
    $Session:FinalAzureGroups

    if ($Session:FinalAzureGroups -ne $null) {
        Write-Log -File $LogFile -Message "Found Azure Resource Groups owned by $UserFirstName $UserSurname : $Session:FinalAzureGroups"
    }
    else {
        Write-Log -File $LogFile -Message "Not found any Azure Resource Groups owned by $UserFirstName $UserSurname"
    }
}

<#
.SYNOPSIS
Disables the Sugarlearning user.

.DESCRIPTION
Disables the Sugarlearning user.
Makes 3 requests, one to login and get authentication cookies, the second one to fetch the user's GUID, and the third one to disable the account based on the GUID.

.PARAMETER username
The userprincipalname of the user.

.EXAMPLE
PS> Disable-SugarlearningUser -User $Session:SelectedUser.userprincipalname
#>
function Disable-SugarlearningUser {
    Param(
        [Parameter(Mandatory = $True, Position = 0)]
        $username
    )
    $SlPassword = (Get-Content $SugarlearningPassword)

    # Params for Login request to SL, we need to login first to get the cookies
    $params = @{
        Uri     = $SugarlearningLoginURI
        Headers = @{ 
            Accept = 'application/json'
        }
        Method  = 'POST'
        Body    = @{
            IsPersistent = "false"
            Password     = $SlPassword
            Username     = $SugarlearningAccount
        }
    }
    $LoginRequest = Invoke-WebRequest @params -UseBasicParsing

    # Building first cookie, this is not super dynamic
    # Also please don't mind the +'s and -'s done below, this will be revised in a next version to be properly dynamic
    # TODO: Make the cookie acquisition better 
    $FirstCookie = $LoginRequest.RawContent
    $FirstIndex = $FirstCookie.IndexOf("Set-Cookie:")
    $LastIndex = $FirstCookie.IndexOf(".AspNet")
    $CorrectIndex = $LastIndex - $FirstIndex
    $TheCookie = $FirstCookie.Substring(($FirstIndex + 11), ($CorrectIndex - 12))

    # First Cookie Name
    $CookieNameIndex = $TheCookie.IndexOf("=")
    $CookieName = $TheCookie.substring(1, $CookieNameIndex - 1)

    # First Cookie Value
    $CookieValueIndex = $TheCookie.IndexOf(";")
    $CookieValue = $TheCookie.substring(($CookieNameIndex + 1), ($CookieValueIndex - $CookieNameIndex - 1))

    # First Cookie expires
    $CookieExpiryIndex = $TheCookie.IndexOf(";", ($CookieValueIndex + 1))
    $CookieExpiry = $TheCookie.substring((($CookieValueIndex + 1) + 9), (($CookieExpiryIndex - $CookieValueIndex - 1) - 9))

    # First Cookie path
    $CookiePathIndex = $TheCookie.IndexOf(";", ($CookieExpiryIndex + 1))
    $CookiePath = $TheCookie.substring(($CookieExpiryIndex + 1 + 6), ($CookiePathIndex - $CookieExpiryIndex - 1 - 6))

    # First Cookie domain
    $CookieDomainIndex = $TheCookie.IndexOf(";", ($CookiePathIndex + 1))
    $CookieDomain = $TheCookie.substring(($CookiePathIndex + 1 + 8), ($CookieDomainIndex - $CookiePathIndex - 1 - 8))

    # Cookie HTTP Only
    $CookieHttpIndex = $TheCookie.IndexOf(";", ($CookieDomainIndex + 1))
    $CookieHttp = $TheCookie.substring(($CookieDomainIndex + 1 + 1), ($CookieHttpIndex - $CookieDomainIndex - 1 - 1))

    # This session will be used by both requests
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession   
    
    # Creating the cookie to be added to the session
    $cookie1 = New-Object System.Net.Cookie 

    $cookie1.Name = $CookieName
    $cookie1.Value = $CookieValue
    $cookie1.Domain = $CookieDomain
    $cookie1.Path = $CookiePath
    $cookie1.HttpOnly = $true
    $cookie1.Expires = $CookieExpiry
    $session.Cookies.Add($cookie1);

    # Building second cookie
    $AspCookie = $LoginRequest.RawContent
    $FirstIndex = $AspCookie.IndexOf(".AspNet.ApplicationCookie")
    $LastIndex = $AspCookie.IndexOf("Request-Context")

    $CorrectIndex = $LastIndex - $FirstIndex
    $TheCookie = $AspCookie.Substring($FirstIndex, $CorrectIndex)

    # Cookie Name
    $CookieNameIndex = $TheCookie.IndexOf("=")
    $CookieName = $TheCookie.substring(0, $CookieNameIndex)

    #Cookie Value
    $CookieValueIndex = $TheCookie.IndexOf(";")
    $CookieValue = $TheCookie.substring(($CookieNameIndex + 1), ($CookieValueIndex - $CookieNameIndex - 1))

    # Cookie path
    $CookiePathIndex = $TheCookie.IndexOf(";", ($CookieValueIndex + 1))
    $CookiePath = $TheCookie.substring(($CookieValueIndex + 1 + 6), ($CookiePathIndex - $CookieValueIndex - 1 - 6))

    $cookie2 = New-Object System.Net.Cookie 

    $cookie2.Name = $CookieName
    $cookie2.Value = $CookieValue
    $cookie2.Domain = "my.sugarlearning.com"
    $cookie2.Path = $CookiePath
    $cookie2.HttpOnly = $true
    $cookie2.Secure = $true
    $session.Cookies.Add($cookie2);
    
    # Params for second request to get users ID, after cookies have been taken above
    $params2 = @{
        Uri     = $SugarlearningUserURI
        Headers = @{ Accept = '*/*'
        }
        Method  = 'GET'
    }
    $SugarlearningResponse = Invoke-WebRequest @params2 -WebSession $session -UseBasicParsing
    $SLJson = $SugarlearningResponse | ConvertFrom-Json
    $SLJson | ForEach {         
        if ($_.UserName -eq $username) {
            Write-Log -File $LogFile -Message $_.UserName
            $GUID = $_.Id
            Write-Log -File $LogFile -Message $GUID
        } 
    }

    # Params for third request to actually disable the user
    $EmptyArray = ConvertTo-Json @()
    $params3 = @{
        Uri     = $SugarlearningDisableUserURI
        Headers = @{ 
            Accept = 'application/json'
        }
        Body    = @{
            Message         = ""
            AssignedGroups  = $EmptyArray
            RoleGroup       = 1
            SelectedModules = $EmptyArray
            Identifier      = $GUID
            IsRemoving      = $true
            Note            = $null
        }        
        Method  = 'POST'
    }
    $SugarlearningDisable = Invoke-WebRequest @params3 -WebSession $session -UseBasicParsing
    Write-Log -File $LogFile -Message "Disabled $username with GUID $GUID in Sugarlearning..."
}

<#
.SYNOPSIS
Disables the user's Admin account in AD, sets its description to the current date and moves it to the correct OU.

.DESCRIPTION
Disables the user's Admin account in AD, sets its description to the current date and moves it to the correct OU.
Also writes in the log what is happening.

.PARAMETER username
The UserPrincipalName of the user.

.EXAMPLE
PS> Disable-AdminUser -User $Session:SelectedUser.UserPrincipalName
#>
function Disable-AdminUser {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )
    $AdminUsername = "Admin" + $username
    $FoundAdminUsername = Get-ADUser -Filter "UserPrincipalName -eq '$AdminUsername'" -Server $DCServer
    if ($FoundAdminUsername -eq $null) {
        Write-Log -File $LogFile -Message "Not found any admin users for $FoundAdminUsername..."
    }
    else {
        Write-Log -File $LogFile -Message "Found admin user: $FoundAdminUsername"
        Remove-UserFromAllADGroups -User $FoundAdminUsername.Samaccountname
        Disable-User -User $FoundAdminUsername.Samaccountname
        Move-ADObject -Identity $FoundAdminUsername -TargetPath $DisabledUsersOU -Server $DCServer
        Write-Log -File $LogFile -Message "Moved to $DisabledUsersOU..."
    }
    $FoundAdminUsername.Samaccountname
}
