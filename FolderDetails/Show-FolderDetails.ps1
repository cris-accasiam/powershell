<#
.SYNOPSIS
Shows the size and access permissions for subfolders and files

.DESCRIPTION
Read the contents of a specific folder.
By default, the script will produce a CSV with a listing of all subfolders and files, their size, date accessed/modified, and permissions.
This might produce a very large file, use the parameters to tweak the output. 

.PARAMETER Path
Path to the folder to be processed. Defaults to the current folder.

.PARAMETER OnlyFiles
If active, it will list only the files in the output.

.PARAMETER OnlyFolders
If active, it will list only the subfolders in the output.

.PARAMETER NoSize
If active, the ouput will not include the sizes of files and folders.

.PARAMETER NoDate
If active, the ouput will not include the create/modified/accessed date of files and folders.

.PARAMETER NoPermissions
If active, the ouput will not include the ACL permissions.

.PARAMETER OnlyExplicitPermissions
If active, it will list only non-inherited permissions. This parameter will be ignored if the NoPermissions switch is on.

.PARAMETER ExpandAccessGroups
If active, it will create rows for each user that is member of a group found in the ACL. This option might increase considerably the size of the output file.
This parameter will be ignored if the NoPermissions switch is on.


.EXAMPLE
Show-FolderDetails -Path C:\temp
This will produce an output.csv file in the current folder with all details about all files and subfolders in C:\temp. The file info will be duplciated for each user/group with access.

#>
[CmdletBinding(DefaultParameterSetName = 'None')]
param (
    [Parameter(Position = 0, Mandatory = $false)] [string] $Path = $pwd.Path,
    [Parameter(Position = 1, Mandatory = $false)] [string] $CSV = ($pwd.Path + '\output.csv'),
    [Parameter(Mandatory = $false)] [switch] $OnlyFiles,
    [Parameter(Mandatory = $false)] [switch] $OnlyFolders,
    [Parameter(Mandatory = $false)] [switch] $NoSize,
    [Parameter(Mandatory = $false)] [switch] $NoDate,
    [Parameter(Mandatory = $false)] [switch] $NoPermissions,
    [Parameter(Mandatory = $false)] [switch] $OnlyExplicitPermissions
#    [Parameter(Mandatory = $false)] [switch] $ExpandAccessGroups
)


function Get-AccountType {
    param($account)
    $names = $account.ToString().Split('\');
    $domain = $names[0]
    $samAccountName = $names[1]

    if ($domain -eq 'NT AUTHORITY') {
        return 'User'
    }
    if ($domain -eq 'BUILTIN') {
        return 'LocalGroup'
    }
    try {
        $adObj = Get-ADObject -Filter ('SamAccountName -eq "{0}"' -f $SamAccountName)
        if ($adObj) {
            return $adObj.ObjectClass
        } 
    }
    catch {
        Write-Debug ('Error retrieving data from AD for ' + $samAccountName)
    }

    return 'Unknown'
}


function Write-GroupMembersToLog {
    param($log, $group, $data)

    $groupName = ($group.ToString().Split('\'))[1]
    $members = Get-ADGroupMember $groupName         
    foreach ($member in $members) {
        $line = $data.linePrefix + ',"' + $member.name + '","' + $data.accountType + '","' + $data.permission + '","' + $data.type + '","' + $data.inherited + '"'
        $log.WriteLine($line)
        Write-Debug $line
    }
}


function Write-ACLsToLog {
    param (
        $log, 
        $acl,
        $linePrefix
    )

    foreach ($entry in $acl) {        
        # Skip inherited permissions, if the OnlyExplicitPermissions switch is on
        if (($OnlyExplicitPermissions -eq $true) -and ($entry.IsInherited -eq $true)) {
            continue
        }
        # Make a copy of existing output data
        $newLine = $linePrefix

        # Add the extra properties
        $account = $entry.IdentityReference.ToString()
        $accountType = Get-AccountType -account $account
        $permission = $entry.FileSystemRights.ToString()
        $type = $entry.AccessControlType.ToString()
        $newLine = $newLine + ',"' + $account + '","' + $accountType + '","' + $permission + '","' + $type + '"'
        
        if ($OnlyExplicitPermissions -eq $false) {
            if ($entry.IsInherited -eq $true) {
                $vi = 'Inherited'
            }
            else {
                $vi = 'Direct'
            }
            $newLine = $newLine + ',"' + $vi + '"'
        }
        # Write to output
        $log.WriteLine($newLine)
        Write-Debug $newLine
<#
        # TODO: Write group members to log
        if (($ExpandAccessGroups -eq $true) -and ($accountType -eq 'group')) {   
            $tempObj = [PSCustomObject]@{
                account     = $account
                accountType = $accountType
                permission  = $permission
                type        = $type
                inherited   = $vi
                linePrefix  = $lineOut
            }
            Write-GroupMembersToLog -log $log -group $account -data $tempObj                   
        }        
#>
    }
}


# ComObject as parameter
function Get-FileInfo {
    param (
        $file
    )
    $obj = [PSCustomObject]@{
        path            = $file.Path
        size            = $file.Size
        fileType        = 'File'
        accessDate      = $file.DateLastAccessed
        modifiedDate    = $file.DateLastModified
        maxAccessDate   = $file.DateLastAccessed
        maxmodifiedDate = $file.DateLastAccessed
    }
    if (($OnlyFolders -eq $true) -and ($obj.fileType -eq 'File') ) {
        # No need to output any file info.
        return $obj
    }
    
    Write-Info -obj $obj

    return $obj
}

function Add-DateToOutput {
    param($objDate)
    
    $d = ''
    if ($obj.accessDate) {            
        try {
            $d = $objDate.toString('yyyy-MM-dd')
        }
        catch {
            Write-Debug "Cannot convert date to string for: $objDate"
            $d = $objDate
        }
    }

    return $d
}

<# 
PSCustomObject as parameter, for either a file or folder
[PSCustomObject]@{
        path
        size
        fileType
        accessDate
        modifiedDate
        maxAccessDate - only for folder object
        maxModifiedDate - only for folder object
    }
#>
function Write-Info {
    param (
        [Parameter(Mandatory = $true)]$obj
    )
    
    # Build the output row
    $lineOut = '"' + $obj.path + '","' + $obj.fileType + '"'
    
    # Add size info
    if ($NoSize -eq $false) {
        $lineOut = $lineOut + ',"' + $obj.size + '"'
    }
    
    # Add date info
    if ($NoDate -eq $false) {        
        $lineOut = $lineOut + ',"' + (Add-DateToOutput -objDate $obj.accessDate) + '"'
        $lineOut = $lineOut + ',"' + (Add-DateToOutput -objDate $obj.modifiedDate) + '"'
        $lineOut = $lineOut + ',"' + (Add-DateToOutput -objDate $obj.maxAccessDate) + '"'
        $lineOut = $lineOut + ',"' + (Add-DateToOutput -objDate $obj.maxModifiedDate) + '"'                 
    }

    # Add permission info
    if ($NoPermissions) {
        # Write to output and return
        $log.WriteLine($lineOut)
        Write-Debug $lineOut
        return $true
    }

    # Read the ACL for the file    
    # The file or folder name might have non-standard characters. Use Get-ACL -LiteralPath, instead of -Path   
    try { 
        $accessList = (Get-Acl -LiteralPath $obj.path).Access
    } catch {
        Write-Debug ("Could not read ACLs for " + $obj.path)
        $log.WriteLine($lineOut)
        return $true
    }

    # Log ACLs. Each user/group will have it's own entry in output
    Write-ACLsToLog -log $log -acl $accessList -linePrefix $lineOut   

    return $true
}


function Read-FolderContents {
    param (
        $FSO,
        $fullPath,
        $log
    )
    
    $folder = $fso.GetFolder($fullPath)
    # Record the folder data in an object
    # maxAccessDate and maxModifiedDate will be set based on the values of the child items
    $folderObj = [PSCustomObject]@{
        Path            = $folder.Path
        size            = 0
        fileType        = 'Folder'
        accessDate      = $folder.DateLastAccessed
        modifiedDate    = $folder.DateLastModified
        maxAccessDate   = Get-Date $folder.DateLastAccessed
        maxModifiedDate = Get-date $folder.DateLastModified
    }
    foreach ($file in $folder.Files) {
        $fileObj = Get-FileInfo $file
        
        # Process some folder-related properties, only if the OnlyFiles switch is not present (either default or OnlyFolders on).
        if ($OnlyFiles -eq $false) {
            # Add the file size to the parent folder size.        
            $folderObj.size = $folderObj.size + $fileObj.size
        }
    }
    
    foreach ($subfolder in $folder.Subfolders) {   
        #Write-Host 'looking in ' $subfolder.Path     
        $subfolderObj = Read-FolderContents -FSO $fso -fullPath $subfolder.Path -log $log
        if ($OnlyFiles -eq $true) {
            # No need to output any folder info.
            continue    
        }
        # Add the subfolder total size to the parent
        $folderObj.size = $folderObj.size + $subfolderObj.size

        # The creation and modification dates of a parent folder do not change when the contents of a child folder are modified.
        # We need the most recent access and modified date to show up for the parent folder.     
        if ($folderObj.maxAccessDate -lt $subfolderObj.accessDate) {
            $folderObj.maxAccessDate = $subfolderObj.accessDate
        }
        if ($folderObj.maxModifiedDate -lt $subfolderObj.modifiedDate) {
            $folderObj.maxModifiedDate = $subfolderObj.modifiedDate
        }
    }
    
    if ($OnlyFiles -ne $true) {
        Write-Info -obj $folderObj
    }        
    return $folderObj
}

if ((Test-Path $Path) -ne $true) {
    Write-Host 'Invalid path: ' $Path
    exit
}

if (($OnlyFiles -eq $true) -and ($OnlyFolders -eq $true)) {
    Write-Host 'Parameters OnlyFiles and OnlyFolders are mutually exclusive. Please specify only one.'
    exit
}

$fso = New-Object -ComObject Scripting.FileSystemObject -ErrorAction Stop

$log = New-Object System.IO.StreamWriter $CSV -ErrorAction Stop

# Write the header, depending on parameter switches
$header = "Name,Type"
if ($NoSize -eq $false) {
    $header += ",Size"
}
if ($NoDate -eq $false) {
    $header += ",DateAccessed,DateModified,MaxDateAccessed,MaxDateModified"
}
if ($NoPermissions -eq $false) {
    $header += ",Account,AccountType,Permission,PermissionType"    
    if ($OnlyExplicitPermissions -eq $false) {
        $header += ",isInherited"
    }    
}
$log.WriteLine($header)

# Build a list of local group and user names
if ($NoPermissions -eq $false) {
    Set-Variable -Name 'localUsers' -Value (Get-LocalUser).Name -Scope Global
    Set-Variable -Name 'localGroups' -Value (Get-LocalGroup).Name -Scope Global
}

Read-FolderContents -FSO $fso -fullPath $Path -log $log | Out-Null

$log.Flush()
$log.Close()
