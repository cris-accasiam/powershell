# Description

This script is a take on the phone contact update scripts that exist elsewhere, using Graph API commands and Entra ID as source.

**Important**

Set up a new Enterprise Application in Microsoft Entra Admin Center before trying to run the script. Then assign sufficient permissions to the application, such as User.Read.All and Contacts.ReadWrite. *Read below for an alternate method.*

As it is now, the script uses the old method of establishing a Graph API connection, based on Tenant ID, Application ID and Application Secret. 

Alternatively, you can replace everything from line 133 until 149 with a simple 
```
Connect-MgGraph -Scopes "User.Read.All", "Contacts.ReadWrite.All"
```
and sign in with an account with enough rights to read-write user's contacts. If you use this method, you don't need an Enterprise Application set up beforehand.

### How it works

The script will first build a list of contacts based on the entries in Entra ID that:
* are synced from on-prem AD
* are not guest accounts
* have an entry in either "Mobile Phone" or "Business Phone"

This will be the list of contacts to be synced.

For each user from the list above, it will create a subfolder in its Contacts folder, and create a new contact for each entry in the list. If the folder already exists, it will be deleted.

In this way, the automatically created contacts are kept separated from the ones created by the users themselves. 
Of course, if they change anything in a contact that is created by the script, those changes will be lost on the next sync.

### Parameters
The script takes a list of UserPrincipalName values as parameter. Usually these are email addresses.

If any values are given, the full list of contacts will be updated only for the users with UserPrincipalName matching the values.

If left empty, the contacts will be updated for all users with a phone number.

The script will prompt for Tenant ID, Application ID, Application Secret, and a folder name.

### Required Dependencies

Microsoft.Graph.Users
Microsoft.Graph.PersonalContacts
MSAL.PS
