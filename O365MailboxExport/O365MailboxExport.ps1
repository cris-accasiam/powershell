<# 
    This script should be run under an account that has permissions to create an eDiscovery case, run searches and export results.
    The "Unified Export Tool" needs to be already installed.
    Steps:
    - Create a Compliance Search for the username/mailbox
    - Start the search, and wait for it to complete
    - Start an "Export" action for the search results, and wait for it to complete
    - Use the "Unified Export Tool" to download the search results as PST
    The progress report for the download is not completly accurate, as the estimated file size is not always matching the actual size on disk of the downloaded file
#>
[CmdletBinding()]
param (   
    [string]$identity,
    [string]$email,
    [string]$exportLocation #the path where the mailbox will be exported !NO TRAILING BACKSLASH!
)

while (($null -eq $identity) -or ($identity -eq '')) {
    $identity = Read-Host -Prompt 'Please enter a username'
}
while (($null -eq $mail) -or ($mail -eq '')) {
    $email = Read-Host -Prompt 'Please enter the complete email address for the user'
}
while (($null -eq $exportLocation) -or ($exportLocation -eq '') -or ((Test-Path $exportLocation) -eq $false)) {
    $exportLocation = Read-Host -Prompt 'Please specify an existing folder where the mailbox will be exported'    
}
# Remove the trailing backslash from the path
$exportLocation = $exportLocation.TrimEnd('\')

# Connect to O365
Connect-IPPSSession -Credential (Get-Credential)

Write-Host "Check if a search for email $email exists"
$previousSearch = Get-ComplianceSearch -Identity $identity -ErrorAction SilentlyContinue

if ($null -eq $previousSearch) {
    Write-Host "Creating a new search for email $email. You can find the search in the 'mailbox export' case"
    New-ComplianceSearch -case 'mailbox export' -Name $identity -ExchangeLocation $email -ErrorAction Stop
}

Write-Host "Starting the search"
Start-ComplianceSearch -Identity $identity -ErrorAction Stop | Out-Null

Write-Host "Search is running. Wait for it to complete."
do {
    Start-Sleep -Seconds 5	
    $job = Get-ComplianceSearch -Identity $identity	
} 
while ($job.Status -ne 'Completed')
Write-Host "Search completed"

Write-Host "Starting the PST export action on the search results"
New-ComplianceSearchAction -SearchName $identity -Export -Format Fxstream -ArchiveFormat PerUserPST -Scope BothIndexedAndUnindexedItems -EnableDedupe $true -ErrorAction Stop | Out-Null

Write-Host "Export is running. Wait for it to complete. It might take some time."
$taskName = $identity + '_Export'
do {
    Start-Sleep -Seconds 10
    $job = Get-ComplianceSearchAction -Identity $taskName
    Write-Host -NoNewline " ."    
}
while ($job.Status -ne 'Completed')
Write-Host "Export completed. " -NoNewline

$exportDetails = Get-ComplianceSearchAction -Identity $taskName -IncludeCredential -Details -ErrorAction Stop
$ExportDetails = $ExportDetails.Results.split(";")
$ExportContainerUrl = $ExportDetails[0].trimStart("Container url: ")
$ExportSasToken = $ExportDetails[1].trimStart(" SAS token: ")
$ExportEstSize = ($ExportDetails[18].TrimStart(" Total estimated bytes: ") -as [double])
$EstimatedSize = $ExportEstSize / 1GB
Write-Host "Estimated bytes: $EstimatedSize GB"

Write-Host "Looking for the Unified Export Tool"
$ExportExe = (Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter microsoft.office.client.discovery.unifiedexporttool.exe -Recurse).FullName | Where-Object { $_ -notmatch "_none_" } | Select-Object -First 1
if (!$ExportExe) {
    Write-Host 'Unified Export Tool not found'
    Write-Host "Go to https://compliance.microsoft.com/ and download the file"        
}
else {
    # Download the search results
    Write-Host "Initiating download. Saving export to: $exportLocation"
    $Arguments = "-name ""$identity""", "-source ""$ExportContainerUrl""", "-key ""$ExportSasToken""", "-dest ""$exportLocation""", "-trace true"
    Start-Process -FilePath "$ExportExe" -ArgumentList $Arguments -ErrorAction Stop

    # The export is now running in the background and can be found in task manager. Show a progress bar while the process is running
    while (Get-Process microsoft.office.client.discovery.unifiedexporttool -ErrorAction SilentlyContinue) {    
        $Downloaded = Get-ChildItem $exportLocation\$identity -Recurse | Measure-Object -Property Length -Sum | Select-Object -ExpandProperty Sum
        Write-Progress -Id 1 -Activity "Export in Progress" -Status ("Downloading... " + ($Downloaded / 1GB) + " / " + $EstimatedSize + "GB")        
    }
    Write-Host " Download Complete!"

    pause #wait for a key press to continue
}

Write-Host 'The script will now disconnect from Exchange Online'
Disconnect-ExchangeOnline




