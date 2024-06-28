<#
.SYNOPSIS
Update phone numbers for Office 365 users
Required Dependencies: Microsoft.Graph.Users, Microsoft.Graph.PersonalContacts, MSAL.PS

.DESCRIPTION
Update phone numbers for Office 365 users. The phone numbers are read from each user's properties.
The new entries will be created in a new subfolder inside Contacts. 
Every time this script runs, it will delete the existing folder, then create a new one with the same name.

.PARAMETER usersToUpdate
A list of UserPrincipalName to update. If left empty, all users with a phone number will be updated.

.EXAMPLE
Update-PhoneContacts
Update the phone contacts for all users

.EXAMPLE
Update-PhoneContacts User1 User2
Update the phone contacts for User1 and User2

.LINK
Github repo: https://github.com/cris-accasiam/powershell
#>


param(
    [Parameter(Position=0,ValueFromRemainingArguments=$true)][String[]]$usersToUpdate
    )

Import-Module MSAL.PS
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.PersonalContacts

function Read-ContactFolders {
    param ($userId)      
    $contactFolders = Get-MgUserContactFolder -UserId $userId -ErrorAction Stop                 
    return $contactFolders;
}

function Create-ContactFolder {
    param ($userId, $folderName)

    # create the new folder under the user's default contacts folder
    $params = @{        
        DisplayName = $folderName
    }
    $newFolder = New-MgUserContactFolder -UserId $userId -BodyParameter $params -ErrorAction Stop

    return $newFolder.Id
}

function Delete-ContactFolder {
    param($userId, $contactFolder)
    
    # workaround for issue https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/2743    
    foreach ($folder in $contactFolder) {
        # generate a unique name
        $newName = (New-Guid).Guid
        # rename the Contacts folder
        Update-MgUserContactFolder -ContactFolderId $folder.id -UserId $userId -DisplayName $newName -ErrorAction Stop | Out-Null
        # remove the folder
        Remove-MgUserContactFolder -ContactFolderId $folder.id -UserId $userId -ErrorAction Stop | Out-Null
    }    
}

function Create-Contact {
    param ($userId, $folderId, $contactData)

    $contactParam = @{
        GivenName      = $contactData.GivenName
        Surname        = $contactData.Surname
        FileAs         = $contactData.FileAs
        DisplayName    = $contactData.DisplayName
        CompanyName    = $contactData.CompanyName
        EmailAddresses = @(
            @{ Address = $contactData.Email; Name = $contactData.Email }
        )
        BusinessPhones = @($contactData.PhoneNumber) # we want our contacts to go under business phones, not private
    }
    # A UPN can also be used as -UserId.
    New-MgUserContactFolderContact -UserId $userId -ContactFolderId $folderId -BodyParameter $contactParam -ErrorAction Stop | Out-Null
}

function Get-Contacts {
    param ($userArray)

    $contactsToImport = @()

    # find all active users that are synced fron on-prem and have a mobilePhone entry
    foreach ($u in $userArray) {
        # discard all disabled accounts
        if ($u.AccountEnabled -eq $false) {
            continue
        }
        # discard guests and accounts that are not synced from on-prem AD
        if (($u.UserType -eq "guest") -or ($null -eq $u.OnPremisesDomainName) -or ($null -eq $u.OnPremisesSyncEnabled)) {
            continue
        }
        # discard entries without a phone number in AD
        if (($null -eq $u.MobilePhone) -and ($null -eq $u.BusinessPhones)) {
            continue
        }
        # get the phone number
        $phoneNumber = ''
        if ($null -ne $u.MobilePhone) {
            $phoneNumber = $u.MobilePhone
        }
        elseif ($null -ne $u.BusinessPhones) {
            $phoneNumber = $u.BusinessPhones[0]
        }
        # phoneNumber should hold a value at this point
        if ([string]::IsNullOrEmpty($phoneNumber)) {
            continue
        }
        # save the data in a custom object
        $contact = [PSCustomObject]@{
            Id                = $u.Id
            UserPrincipalName = $u.UserPrincipalName
            GivenName         = $u.GivenName
            Surname           = $u.Surname
            FileAs            = $u.DisplayName
            DisplayName       = $u.DisplayName
            Email             = $u.Mail
            PhoneNumber       = $phoneNumber
        }
        $contactsToImport += $contact
    }

    return $contactsToImport
}

$tenantId = Read-Host -Prompt "Tenant ID"
$clientId = Read-Host -Prompt "Application ID"
$secretId = Read-Host -AsSecureString -Prompt "Application Secret ID"
$autosyncFolderName = Read-Host -Prompt "Contacts folder name"

$MsalToken = Get-MsalToken -TenantId $TenantId -ClientId $clientId -ClientSecret $secretId

# Graph API v1.0 accepts a simple string, v2 needs a SecureString for the AccessToken parameter
$targetParameter = (Get-Command Connect-MgGraph).Parameters['AccessToken']
if ($targetParameter.ParameterType -eq [securestring]) {
    Connect-MgGraph -AccessToken ($MsalToken.AccessToken | ConvertTo-SecureString -AsPlainText -Force) -NoWelcome -ErrorAction Stop
}
else {
    Connect-MgGraph -AccessToken $MsalToken.AccessToken -ErrorAction Stop
}

# read all users
Write-Progress -Status "Read all users" -Activity "Preparing"
$allUsers = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName, GivenName, Surname, MobilePhone, BusinessPhones, Mail, UserType, OnPremisesDomainName, OnPremisesSyncEnabled, AccountEnabled  -ErrorAction Stop

# build the list of contacts
Write-Progress -Status "List contacts" -Activity "Preparing"
$contactsToImport = Get-Contacts -userArray $allUsers

# update the contacts for each username in the List contactsToImport
foreach ($user in $contactsToImport) {  
    
    if (($usersToUpdate.count -gt 0) -and ($user.UserPrincipalName -notin $usersToUpdate)) {        
        continue
    } 
    
    $userId = $user.Id
    Write-Progress -Status ("Look for a " + $autosyncFolderName + " folder") -Activity ("Processing user " + $user.UserPrincipalName)

    # check if the autosync folder exists           
    $folders = Get-MgUserContactFolder -UserId $userId -Filter "DisplayName eq '$autosyncFolderName'"

    # if the autosync folder exists, delete it    
    if ($folders.Count -gt 0) {
        Write-Progress -Status "Delete the $autosyncFolderName folder" -Activity ("Processing user " + $user.UserPrincipalName)        
        Delete-ContactFolder -userId $userId -contactFolder $folders          
    }

    # create a new autosync folder    
    Write-Progress -Status "Create the $autosyncFolderName folder" -Activity ("Processing user " + $user.UserPrincipalName)
    $newFolderId = Create-ContactFolder -userId $userId -folderName $autosyncFolderName        

    # create new contact entries in the folder we just created    
    foreach ($contact in $contactsToImport) {        
        # save the new contact
        Write-Progress -Status ("Add " + $contact.DisplayName + " to contacts") -Activity ("Processing user " + $user.UserPrincipalName)
        Create-Contact -userId $userId -folderId $newFolderId -contactData $contact
        # sleep to avoid Microsoft Graph throttling
        Start-Sleep -Milliseconds 200                            
    }    
    Write-Progress -Status "Completed" -Activity ("Processing user " + $user.UserPrincipalName)
    # sleep to avoid Microsoft Graph throttling
    Start-Sleep -Seconds 1                            
}

Disconnect-Graph | Out-Null


