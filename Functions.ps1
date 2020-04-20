<#
   .SYNOPSIS
   All functions that will be used by LeavingDashboard.ps1
      
   #>

$LogFile = "\\fileserver\Backups\SSWLeavingStandard.log"
# Function to write logs on the screen and server
Function LogWrite {
    $username = $env:USERNAME   
    $PcName = $env:computername
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $PcName $Username $args"
   
    Add-content $Logfile -value $Line
    Write-Host $Line
    $Session:log += $Line
}
   
# ------- BEGIN FILE SHARE SEARCH IN \\FILESERVER -------
function New-BackupSearch {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )
      
    $fulluser = "Data" + $username
    $UserBackupFiles = Get-ChildItem "\\fileserver.sydney.ssw.com.au\UserBackups\" -filter $fulluser
    $UserBackupFileName = "\\fileserver.sydney.ssw.com.au\UserBackups\" + $fulluser
    $DataSSWFiles = Get-ChildItem "\\fileserver.sydney.ssw.com.au\DataSSW\" -filter $fulluser
    $DataSSWFileName = "\\fileserver.sydney.ssw.com.au\DataSSW\" + $fulluser
   
    $folderExistInUserBackups = $true
    $folderExistInDataSSW = $true
   
    if ($UserBackupFiles -eq $null) {
        $folderExistInUserBackups = $false
        Logwrite "Found nothing in \\fileserver.sydney.ssw.com.au\UserBackups\"
    }
    else {
        Logwrite "Found files in \\fileserver.sydney.ssw.com.au\UserBackups\"
    }
   
    if ($DataSSWFiles -eq $null) {
        $folderExistInDataSSW = $false
        Logwrite "Found nothing in \\fileserver.sydney.ssw.com.au\DataSSW\"
    }
    else {
        Logwrite "Found files in \\fileserver.sydney.ssw.com.au\DataSSW\"
    }
   
    if (($folderExistInUserBackups -eq $true) -or ($folderExistInUserBackups -eq $true)) {
        Logwrite "Found something in DataSSW or UserBackups. Backing it all up..."
        7z a \\fileserver\DataSSW\ExEmployees\PcBackup\$fulluser".zip" $DataSSWFileName
        7z a \\fileserver\DataSSW\ExEmployees\PcBackup\$fulluser".zip" $UserBackupFileName
        Logwrite "Added all found content in \\fileserver\DataSSW\ExEmployees\PcBackup\$fulluser.zip"
    }
}
# ------- END FILE SHARE SEARCH IN \\FILESERVER -------
   
<# ------- BEGIN CONTENT SEARCH IN OFFICE 365 -------
   function New-ContentSearch {
       Param(
           [Parameter(Mandatory=$True,ParameterSetName="User")]
           $username
       )
       # Use another .ps1 to connect to Exchange Online with MFA (https://o365reports.com/2019/10/05/connect-all-office-365-services-powershell/)
       try {
           & "$PWD\ConnectO365Services.ps1" -Services SecAndCompCenter -MFA
           LogWrite "Connected correctly to Office 365 Security & Compliance Center..."
       }
       catch {
           Logwrite "Could not connect to Office 365 Security & Compliance Center..."
       }
   
       # Make a new Content search in Office 365 Compliance Center, then start it
       New-ComplianceSearch -ExchangeLocation $username@ssw.com.au -Name "$Username Leaving Standard" | Start-ComplianceSearch
       LogWrite "Successfully created and started the Search $Username Leaving Standard..."
   
       # Start Internet Explorer on the correct page
       Start-Process iexplore.exe "https://protection.office.com/contentsearchbeta"
       LogWrite "Internet Explorer just opened in your machine. Log into https://protection.office.com/contentsearchbeta | $Username Leaving Standard | More | Export Results"
   }
   # ------- END CONTENT SEARCH IN OFFICE 365 -------
   #>
#------- BEGIN GROUP REMOVAL IN ACTIVE DIRECTORY -------
   
function Remove-UserFromAllADGroups {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )
    get-adgroup -filter "Name -notlike '*Report*' -and Name -notlike 'Domain Users'" -Server "ssw-dc4.sydney.ssw.com.au" -Credential $Session:Credentials | Remove-ADGroupMember -member $username -Confirm:$false
    LogWrite "Removed all groups from $Username..."
}
# ------- FINISH GROUP REMOVAL IN ACTIVE DIRECTORY -------
   
# ------- BEGIN OU MOVE IN ACTIVE DIRECTORY -------
function Move-UserToDisabledUserOU {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )
    Move-ADObject -Identity $username -TargetPath "OU=ztDisabledUsers_ToClea,OU=DisabledUsers,OU=yyActive Directory Clean,DC=sydney,DC=ssw,DC=com,DC=au" -Server "ssw-dc4.sydney.ssw.com.au" -Credential $Session:Credentials
    LogWrite "Moved to DisabledUsers OU..."
}
# ------- FINISH OU MOVE IN ACTIVE DIRECTORY -------
   
# ------- BEGIN USER DISABLE IN ACTIVE DIRECTORY -------
function Disable-User {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )
    Disable-ADAccount -Identity $username -Server "ssw-dc4.sydney.ssw.com.au" -Credential $Session:Credentials
    LogWrite "Disabled user..."
    $date = Get-Date -Format "dd/MM/yyyy"
    Set-ADUser $username -Description "Disabled - $date"
    LogWrite "Set user description to today's date..."
}
# ------- FINISH USER DISABLE IN ACTIVE DIRECTORY -------
   
# ------- BEGIN USER DISABLE IN SKYPE FOR BUSINESS -------
function Disable-S4BUser {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )    
    Invoke-Command -ComputerName "SydLync2013P01.sydney.ssw.com.au" -Credential $Session:Credentials -Authentication Credssp -ArgumentList $username -ScriptBlock {
        Import-Module SkypeForBusiness
   
        Disable-CsUser -identity $args[0]
    }
    LogWrite "Disabled Skype For Business user..."
}
# ------- FINISH USER DISABLE IN SKYPE FOR BUSINESS -------
   
# ------- BEGIN HIDE EMAIL FROM ADDRESS LIST -------
function Hide-Email {
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "User")]
        $username
    )
   
    Set-ADUser -Identity $username -Server "ssw-dc4.sydney.ssw.com.au" -Credential $Session:Credentials -Replace @{msexchhidefromaddresslists = $true }
    LogWrite "Hid email from address list..."
}
# ------- FINISH HIDE EMAIL FROM ADDRESS LIST -------
   
# ------- BEGIN CREATE REDIRECT RULE IN EXCHANGE -------
function New-RedirectRule {
    Param(
        [Parameter(Mandatory = $True, Position = 0)]
        $username,
        [Parameter(Mandatory = $True, Position = 1)]
        $target
    )
       
    $O365Account = Get-Content "C:\inetpub\wwwroot\SSWLeavingStandard\O365Account.key"
    $O365Password = Get-Content "C:\inetpub\wwwroot\SSWLeavingStandard\O365AccountPass.key" | ConvertTo-SecureString -Key (Get-Content "C:\inetpub\wwwroot\SSWLeavingStandard\O365AccountKey.key")
    $O365Credential = New-Object System.Management.Automation.PsCredential($O365Account, $O365Password)
   
    Connect-ExchangeOnline -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Credential
    $email = $username.userPrincipalName
    $name = $username.samaccountname
    $FirstName = $username.GivenName
    $LastName = $username.Surname
   
    New-TransportRule "Leaving Standard - $FirstName $LastName" -SentTo $Email -RedirectMessageTo $target
   
    $bodyhtml = "<div style='font-family:Calibri;'>"
    $bodyhtml += "</H3>"
    $bodyhtml += "<p>We just added a Redirect Rule in our Exchange Server. </p>"
    $bodyhtml += "<p>You will now receive all emails from $Email, as per SSW's Leaving Standard (this means $FirstName left SSW). </p>"
    $bodyhtml += "<p>Tip: You can find a log file with more information at <a href=$LogFile> $LogFile </a></p>"
    $bodyhtml += "If you think this is wrong, please contact SysAdmins.</p>"
    $bodyhtml += "<p></p>"
    $bodyhtml += "<p>-- Powered by SSW.LeavingStandard<br/> Server: $env:computername </p>"
       
    Send-MailMessage -From "info@ssw.com.au" -to $target -Subject "SSW.LeavingStandard - Mails are now being redirected from $Email to you" -Body $bodyhtml -SmtpServer "ssw-com-au.mail.protection.outlook.com" -BodyAsHtml
    LogWrite "Created redirect rule in exchange..."
}
# ------- FINISH CREATE REDIRECT RULE IN EXCHANGE -------
   
# ------- BEGIN CHECKIN ASSETS IN SNIPE -------
function CheckIn-SnipeAssets {
    Param(
        [Parameter(Mandatory = $True, Position = 0)]
        $username
    )
    $SnipeApiKey = get-content "C:\inetpub\wwwroot\SSWLeavingStandard\snipeapi.key"
    $SnipeUser = $username.samaccountname
   
    #Here we save the Snipe Module for Powershell (https://www.powershellgallery.com/packages/SnipeitPS) to easily access and use it
    Import-Module SnipeITPS
    #Save-Module -Name SnipeITPS -Path $env:TEMP
    #$snipeModulePath = (Get-ChildItem -Path "$env:TEMP\SnipeITPS" -Directory).FullName
    #Import-Module -FullyQualifiedName $snipeModulePath
   
    Set-Info -URL 'https://snipe.ssw.com.au' -apiKey $SnipeApiKey
       
    $UserAssets = Get-Asset -search $SnipeUser | where { $_.assigned_to.username -eq $SnipeUser } 
    if ($UserAssets -eq $null) {
        LogWrite "No assets to check-in in Snipe..."
    }
    else {
        $UserAssets | foreach { LogWrite "Checked in asset ID "$_.id" / Asset tag "$_.asset_tag" in Snipe..." }
        $UserAssets | Foreach { Reset-AssetOwner $_.id }
    }
       
}
# ------- FINISH CHECKIN ASSETS IN SNIPE -------
   
# ------- BEGIN USER DEPROVISION IN ZENDESK -------
function Disable-ZendeskUser {
    Param(
        [Parameter(Mandatory = $True, Position = 0)]
        $username
    )
    $ZendeskUsername = $username.userPrincipalName
   
    #Get token and username
    $ZendeskUser = get-content "C:\inetpub\wwwroot\SSWLeavingStandard\ZendeskUser.key"
    $token = get-content "C:\inetpub\wwwroot\SSWLeavingStandard\ZendeskTok.key"
   
    $params = @{
        Uri     = 'https://pdi-ssw.zendesk.com/api/v2/users.json?role=agent'
        Headers = @{ Accept = 'application/json'
            'Authorization' = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($ZendeskUser):$($Token)")); 
        }
        Method  = 'GET'
    }
    $ZendeskUsers = Invoke-RestMethod @params
    $ZendeskEmployee = $ZendeskUsers.users | where email -eq $ZendeskUsername
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
        Uri         = 'https://pdi-ssw.zendesk.com/api/v2/users/create_or_update.json'
        Headers     = @{ Accept = 'application/json'
            'Authorization' = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($ZendeskUser):$($Token)")); 
        }
        Body        = $JsonBody2
        ContentType = 'application/json'
    }
    Invoke-RestMethod -Method "POST" @params2
    LogWrite "Successfully downgrade agent to end-user in Zendesk..."
}
# ------- FINISH USER DEPROVISION IN ZENDESK -------
   
# ------- BEGIN USER DEPROVISION IN AZURE DEVOPS -------
function Disable-AzureDevopsUser {
    Param(
        [Parameter(Mandatory = $True, Position = 0)]
        $username
    )
    $AzureDevopsusername = $username.userPrincipalName
   
    #Start SSW1 API Request
    $AzureDevopsToken = Get-Content "C:\inetpub\wwwroot\SSWLeavingStandard\AzureDevopsTok1.key"
    $AzureDevopsParams = @{
        Uri     = "https://vsaex.dev.azure.com/ssw/_apis/userentitlements?top=10000&api-version=5.1-preview.2"
        Headers = @{ Accept = '*/*'
            'Authorization' = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($AzureDevopsToken)); 
        }
        Method  = 'GET'
    }
    $AzureDevopsUsers = Invoke-RestMethod @AzureDevopsParams
    $AzureDevopsEmployee = $AzureDevopsUsers.members
    $AzureDevopsId = $AzureDevopsEmployee | where { $_.user.mailAddress -eq "$AzureDevopsusername" }
    $AzureDevopsIdOnly = $AzureDevopsId.id
    $AzureDevopsEmailOnly = $AzureDevopsId.user.mailAddress
    if ($AzureDevopsId -eq $null) {
        LogWrite "Azure DevOps user not found in SSW1..."
    }
    else {
        LogWrite "Azure DevOps user found in SSW1: $AzureDevopsEmailOnly..."
    }
   
    $AzureDevopsParams2 = @{
        Uri     = "https://vsaex.dev.azure.com/ssw/_apis/userentitlements/" + "$AzureDevopsIdOnly" + "?api-version=5.1-preview.2"
        Headers = @{ Accept = '*/*'
            'Authorization' = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($AzureDevopsToken)); 
        }
        Method  = 'DELETE'
    }
    Invoke-RestMethod @AzureDevopsParams2
    LogWrite "Deleted Azure DevOps user in SSW1: $AzureDevopsEmailOnly..."
   
    #Start SSW2 API Request
    $AzureDevopsToken2 = Get-Content "C:\inetpub\wwwroot\SSWLeavingStandard\AzureDevopsTok2.key"
    $AzureDevopsParams3 = @{
        Uri     = "https://vsaex.dev.azure.com/ssw2/_apis/userentitlements?top=10000&api-version=5.1-preview.2"
        Headers = @{ Accept = '*/*'
            'Authorization' = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($AzureDevopsToken2)); 
        }
        Method  = 'GET'
    }
    $AzureDevopsUsers2 = Invoke-RestMethod @AzureDevopsParams3
    $AzureDevopsEmployee2 = $AzureDevopsUsers2.members
    $AzureDevopsId2 = $AzureDevopsEmployee2 | where { $_.user.mailAddress -eq "$AzureDevopsusername" }
    $AzureDevopsIdOnly2 = $AzureDevopsId2.id
    $AzureDevopsEmailOnly2 = $AzureDevopsId2.user.mailAddress
    if ($AzureDevopsId2 -eq $null) {
        LogWrite "Azure DevOps user not found in SSW2..."
    }
    else {
        LogWrite "Azure DevOps user found in SSW2: $AzureDevopsEmailOnly2..."
    }
   
    $AzureDevopsParams4 = @{
        Uri     = "https://vsaex.dev.azure.com/ssw2/_apis/userentitlements/" + "$AzureDevopsIdOnly2" + "?api-version=5.1-preview.2"
        Headers = @{ Accept = '*/*'
            'Authorization' = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($AzureDevopsToken2)); 
        }
        Method  = 'DELETE'
    }
    Invoke-RestMethod @AzureDevopsParams4
    LogWrite "Deleted Azure DevOps user in SSW2: $AzureDevopsEmailOnly2..."
}
# ------- FINISH USER DEPROVISION IN AZURE DEVOPS -------
   
# ------- BEGIN SEND FINISHED EMAIL -------
function Send-FinishEmail {
    Param(
        [Parameter(Mandatory = $True, Position = 0)]
        $primaryTarget,
        [Parameter(Mandatory = $True, Position = 1)]
        $ownTarget
    )
   
    $bodyhtml2 = @"
       <div style='font-family:Calibri;'>
       <p>SSW Leaving Standard finished.</p>
       <ul><li>1. Email - Backup my email (in Office 365 Compliance Center) to fileserver</li></ul>
       <p><strong>Done (automatically)</strong></p><ul>
       <li>2. File Server - Go through this guide <a href=https://sswcom.sharepoint.com/:w:/g/SysAdmin/EaI1Gk_sds5NqK0QBLqAMK4BR16CgA1JwzWbK2MGiJZs7Q?e=ij7FV8>here<a> in our intranet and totally clean my folders in fileserver </li></ul>
       <p><strong>Done </strong></p>
       <ul><li>3. Active Directory - Remove me from groups in Active Directory (Including Admin Account - leave CRM and default groups)</li></ul>
       <p><strong>Done (automatically)</strong></p><ul>
       <li>4. Active Directory - Disable my account</li></ul>
       <p><strong>Done (automatically)</strong></p><ul>
       <li>5. Active Directory - Edit my account description with 'Disabled - &lt; The Current Date&gt; '</li></ul>
       <p><strong>Done (automatically)</strong></p><ul>
       <li>6. S4B - Remove user from Skype for Business Server</li></ul>
       <p><strong>Done (automatically)</strong></p><ul>
       <li>7. Exchange - Hide my email from address lists</li></ul>
       <p><strong>Done (automatically)</strong></p><ul>
       <li>8. Exchange - Forward my emails to $primaryTarget without leaving a copy on the recipients mailbox. Email the new owner to let him know. Also, add a Mail Flow rule in our EAC to forward my emails to the person above (Go to https://mail.ssw.com.au/ecp | Mail Flow | &nbsp;'Create a new rule...' to do this)</li></ul>
       <p><strong>Done (automatically)</strong></p><ul>
       <li>9. Snipe - Check in all assets checked out to me on Snipe (Go to https://snipe.ssw.com.au/ | Look for my name on the asset list | Check in all assets so we know they are available for other people)</li></ul>
       <p><strong>Done (automatically)</strong></p><ul>
       <li>10. Zendesk - Turn agent to end-user in Zendesk</li></ul>
       <p><strong>Done (automatically)</strong></p><ul>
       <li>11. Azure DevOps - Remove my access from ssw.visualstudio.com (Microsoft Account) and ssw2.visualstudio.com as well</li></ul>
       <p><strong>Done (automatically)</strong></p><ul>
       <li>12. Azure - Check if any Azure Resource Groups are still owned by me - if yes, they need to be handed over to someone else</li></ul>
       <p><strong>Done (automatically)</strong></p><ul>
       <li>13. Active Directory - Move my account to SSW/ztDisabledUsers_ToClean</li></ul>
       <p><strong>Done (automatically)</strong></p><ul>
       <li>14. Active Directory - Remove ExtensionAttribute1 from AD</li></ul>
       <ul><li>15. Azure - Remove my access from Azure (Microsoft Account)</li></ul>
       <ul><li>16. VPN - In NPS, disable my VPN Access  </li></ul>
       <ul><li>17. Exchange - Disable unified messaging for my account.</li></ul>
       <ul><li>18. SugarLearning - Remove my access from SugarLearning</li></ul>
       <ul><li>19. Github - Remove my access from SSW Consulting GitHub</li></ul>
       <ul><li>20. Microsoft Partner Center - Remove my MSDN subscription</li></ul>
       <ul><li>21. Github - Remove from Employee page (https://github.com/SSWConsulting/People | Click on Employee name | Click on EmployeeName.md | Edit | Set currentEmployee to "false") | Merge directly to master)</li></ul>
       <ul><li>22. Invixium - Remove my fingerprint from Control 4 [Sydney Office Only]</li></ul>
       <ul><li>23. Control 4 - Remove Control 4 Account: https://customer.control4.com/account/users</li></ul>
       <ul><li>24. OneDrive - Go into my OneDrive and backup the important files in there, they will be lost after 30 days (Office 365 Admin | Active Users | Select the user | OneDrive tab | Get access to files | Download important files to fileserver) </li></ul> 
       <ul><li>25. CRM - Input the correct date in my user CRM account in the field 'Date Finished'</li></ul>
       <ul><li>26. CRM - Disable my Dynamics 365 (CRM) User account</li></ul>
       <p>Note: Thank you, hopefully we will see you around the user groups!</p>
       <p>-- Partially powered by SSW.LeavingStandard<br> Server: $env:computername </p>
"@
   
    Send-MailMessage -From "info@ssw.com.au" -to $ownTarget -Subject "Leaving SSW - Disable accounts" -Body $bodyhtml2 -SmtpServer "ssw-com-au.mail.protection.outlook.com" -BodyAsHtml
   
}
   
# ------- FINISH SEND FINISHED EMAIL -------
   
# ------- BEGIN SEARCH AZURE TAGS -------
function Search-AzureTags {
    Param(
        [Parameter(Mandatory = $True, Position = 0)]
        $username
    )
    $UserFirstName = $username.GivenName
    $UserSurname = $username.Surname
   
    Import-Module Az
       
    $AesKEy = get-content "C:\inetpub\wwwroot\SSWLeavingStandard\AzureSecret.key"
    $AzureUser = get-content "C:\inetpub\wwwroot\SSWLeavingStandard\AzureServicePrincipal.key"
    $AzureSecret = get-content "C:\inetpub\wwwroot\SSWLeavingStandard\AzureClientSecret.key" | ConvertTo-SecureString -Key $AesKey
    $AzureCredential = New-Object System.Management.Automation.PSCredential($AzureUser, $AzureSecret)
       
    $FinalAzureGroups = @()
    try {
        $Connected = Connect-AzAccount -Credential $AzureCredential -TenantId "ac2f7c34-b935-48e9-abdc-11e5d4fcb2b0" -SubscriptionId "b8b18dcf-d83b-47e2-9886-00c2e983629e" -ServicePrincipal
        LogWrite "Correctly connected to Azure Subscription"
    }
    catch {
        LogWrite "Could not connect to Azure Subscription"
    }
       
    $AzureGroups = Get-AzResourceGroup -Tag @{owner = "$UserFirstName $UserSurname" } | ForEach { $FinalAzureGroups += "" + $_.ResourceGroupName + " " }
    $FinalAzureGroups
   
    if ($FinalAzureGroups -ne $null) {
   
        LogWrite "Found Azure Resource Groups owned by $UserFirstName $UserSurname : $FinalAzureGroups"
    }
    else {
        LogWrite "Not found any Azure Resource Groups owned by $UserFirstName $UserSurname"
    }
   
}
   
# ------- FINISH SEARCH AZURE TAGS -------
   