<# This script searches for the latest 'Backup job summary' email on the 'Backup_Jobs'
folder, retrieves the zip file attachment inside that email, unzips its contents 
on the current working directory and deletes the zip, leaving only the file.

NOTE: THIS ONLY WORKS IF YOU HAVE A NEW BACKUPS EMAIL 
AFTER 5:00 AM FROM THE DAY THIS SCRIPT IS RUN. IT WON'T WORK OTHERWISE

It's recommended that you execute this script by opening Outlook first.

#>


# Adding asssemblies to be able to handle Outlook with a COM object later
Add-Type -assembly "Microsoft.Office.Interop.Outlook"
add-type -assembly "System.Runtime.Interopservices"


try
{           #Use Outlook if it's open already
$outlook = [Runtime.Interopservices.Marshal]::GetActiveObject('Outlook.Application')
    $outlookWasAlreadyRunning = $true
}
catch
{
    try
    {   #Activate Outlook if it's not opened
        $Outlook = New-Object -ComObject Outlook.Application
        $outlookWasAlreadyRunning = $false
    }
    catch
    {
        Write-Host "You must open Outlook first."
        exit
    }
} 


#Use the MAPI namespace, which lets us search and control mailbox elements in Outlook
$namespace = $Outlook.GetNameSpace("MAPI")
#Declare what mailbox folder we need to retrieve. This number is documented on Microsoft docs
$inboxfoldernumber = 6
# Using the namespace, save our inbox folder in a variable
$inbox = $namespace.GetDefaultFolder($inboxfoldernumber)
# Filter our inbox folders: get only the 'Backups_Jobs' folder, where our job emails are.
$job_folder = $inbox.Folders | Where-Object {$_.FolderPath -match 'Backup_Jobs'}

#Get today's date, which we will use for filters later. 
$date =  Get-Date

<# Filter the emails (items) on our jobs folder: get only those older than today's (05:00am),
 with the declared title. Only one email should come; save it in a variable. #>
$latest_job_email = $job_folder.Items | 
Where-Object {$_.Subject -match 'Backup job summary' -and 
              $_.ReceivedTime -gt $date.Date.AddHours(5)}


# Take our previous saved email and get its zip attachment. Save the attachment to a variable.              
$zip_file = $latest_job_email.Attachments | Where-Object {$_.FileName -like '*.zip'}

<# Save the current working directory (cwd) to a variable, which will help us to build folder
paths later #>
$cwd = pwd

#Using the save as file method, save the extracted zip file to the cwd
#To build the path, we use the cwd variable and the zip variable property
$zip_file.SaveAsFile((Join-Path -Path $cwd -ChildPath $zip_file.FileName)) 

# Wait a second for the file to be saved
Start-Sleep -Seconds 1

#Extract the zip's contents to the cwd. .\ is a relative path: it means the cwd
Expand-Archive -Path '.\*.zip' -DestinationPath '.\'
#Delete the zip file and leave only the content
Remove-Item -Path '.\*.zip'













