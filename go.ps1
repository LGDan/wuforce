#Windows Update Force Bypass GPO

#First, set Windows Update to ignore GOP.
if (Test-Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU) {
    #GPO is deployed to manage WU.

    $managedByWSUS = 0

    #Get Properties from obj:
    $wuGPOProperties = Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU

    #Check to see if GPO is telling the OS to point at an updates server:
    if ($wuGPOProperties.UseWuServer -ne $null) {
        #It Is!
        $managedByWSUS = $wuGPOProperties.UseWuServer
    }

    #Temporarily Disable it talking to WSUS
    Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU -Name UseWUServer -Value 0

    #And finally restart the service so it listens to the change.
    Restart-Service wuauserv
}

#Next, Let's get the relevant updates.

$updateSession = New-Object  -Com Microsoft.Update.Session
$updateSession.ClientApplicationID = "wuforce"
$updateSearcher = $updateSession.CreateUpdateSearcher()
Write-Host "Searching for updates..."
$searchResults = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

#Display them in the console...
Write-Host "List of applicable updates for this machine:"
for ($i = 0; $i -lt $searchResults.Updates.Count; $i++ ) {
    $update = $searchResults.Updates.Item($i)
    Write-Host ($i + 1) "> " $update.Title
}


if ($searchResults.Updates.Count -eq 0) {

    #If there are no updates, output to console.
    Write-Host "There are no applicable updates."

}else{
    
    #If there are updates, create a list to download them.

    Write-Host "Creating collection of updates to download:"
    $updatesToDownload = New-Object -Com Microsoft.Update.UpdateColl
    for ($i = 0; $i -lt $searchResults.Updates.Count; $i++) {
        $update = $searchResults.Updates.Item($i)
        $addThisUpdate = $false
        if ($update.InstallationBehavior.CanRequestUserInput) {
            Write-Host "Skipping " $update.Title " because it requires user input."
        }else{
            If ($update.EulaAccepted -eq $false) {
                $update.AcceptEula()   
            }
            $addThisUpdate = $true
        }
        if ($addThisUpdate) {
            Write-Host $i "> Adding " $update.Title
            $updatesToDownload.Add($update)
        }
    }



    if ($updatesToDownload.Count -eq 0) {

        #If there are no updates to download, output to console.
        Write-Host "All applicable updates were skipped."

    }else{

        #Download the updates.
        Write-Host "Downloading Updates..."
        $downloader = $updateSession.CreateUpdateDownloader()
        $downloader.Updates = $updatesToDownload
        $downloader.Download()


        #Prepare a list of updates to install.
        $updatesToInstall = New-Object -Com Microsoft.Update.UpdateColl
        $rebootMayBeRequired = $false
        Write-Host "Successfully downloaded updates:"
        For ($i = 0; $i -lt $searchResults.Updates.Count; $i++) {
            $update = $searchResults.Updates.Item($i)
            If ($update.IsDownloaded) {
                Write-Host $i "> " $update.Title
                $updatesToInstall.Add($update)
                If ($update.InstallationBehavior.RebootBehavior -gt 0) {
                    $rebootMayBeRequired = $true
                }
            }
        }

        #Did the download fail?
        If ($updatesToInstall.Count -eq 0) {

            #Yes, it failed.
            Write-Host "No updates were successfully downloaded."

        }else{

            #It did not fail.
            If ($rebootMayBeRequired) {
                Write-Host "These updates might require a reboot."
            }

            Write-Host "Installing updates..."
            $installer = $updateSession.CreateUpdateInstaller()
            $installer.Updates = $updatesToInstall
            $installationResult = $installer.Install()

            Write-Host "Installation Result: " $installationResult.ResultCode
            Write-Host "Reboot Required: " $rebootMayBeRequired
            Write-Host "List of updates installed:"

            For ($i;$i -lt $updatesToInstall.Count;$i++) {
                Write-Host ($i + 1) "> " $updatesToInstall($i).Title ":" $installationResult.GetUpdateResult($i).ResultCode
            }
        }
    }
}
