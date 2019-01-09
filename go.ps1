#Windows Update Force Bypass GPO
#Written by Dan, Jan 2019


#Intro
Write-Output "== wuforce =="
Write-Output "Checking Registry for GPO Settings..."

#First, set Windows Update to ignore GPO.
if (Test-Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU) {
    #GPO is deployed to manage WU.
    $managedByWSUS = 0

    #Get Properties from obj:
    $wuGPOProperties = Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU

    #Check to see if GPO is telling the OS to point at an updates server:
    if ($null -ne $wuGPOProperties.UseWuServer) {
        #It Is!
        Write-Output "OS Managed by WSUS."
        $managedByWSUS = $wuGPOProperties.UseWuServer
    }


    if ($managedByWSUS) {
        #Temporarily Disable it talking to WSUS
        Write-Output "Temporarily Bypassing WSUS..."
        Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU -Name UseWUServer -Value 0

        #And finally restart the service so it listens to the change.
        Write-Output "Restarting WU Service..."
        Restart-Service -Name "wuauserv" -Force
    }
}

#Next, Let's get the relevant updates.
$updateSession = New-Object -Com Microsoft.Update.Session
$updateSession.ClientApplicationID = "wuforce"
$updateSearcher = $updateSession.CreateUpdateSearcher()

#Look for the updates in prep for download.
Write-Output "Searching for updates..."
$searchResults = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")



Write-Output "List of applicable updates for this machine:"
for ($i = 0; $i -lt $searchResults.Updates.Count; $i++ ) {
    $update = $searchResults.Updates.Item($i)
    $title = $update.Title
    Write-Output "$i > $title"
}

if ($searchResults.Updates.Count -eq 0) {

    Write-Output "There are no applicable updates."
}else{

    Write-Output "Creating collection of updates to download:"
    $updatesToDownload = New-Object -Com Microsoft.Update.UpdateColl
    for ($i = 0; $i -lt $searchResults.Updates.Count; $i++) {
        $update = $searchResults.Updates.Item($i)
        $title = $update.Title
        $addThisUpdate = $false
        if ($update.InstallationBehavior.CanRequestUserInput) {
            Write-Output "Skipping $title because it requires user input."
        }else{
            If ($update.EulaAccepted -eq $false) {
                $update.AcceptEula()
            }
            $addThisUpdate = $true
        }
        if ($addThisUpdate) {
            Write-Output "$i > Adding $title"
            $updatesToDownload.Add($update)
        }
    }

    if ($updatesToDownload.Count -eq 0) {
        Write-Output "All applicable updates were skipped."
    }else{
        Write-Output "Downloading Updates..."
        $downloader = $updateSession.CreateUpdateDownloader()
        $downloader.Updates = $updatesToDownload
        $downloader.Download()

        $updatesToInstall = New-Object -Com Microsoft.Update.UpdateColl
        $rebootMayBeRequired = $false

        Write-Output "Successfully downloaded updates:"
        For ($i = 0; $i -lt $searchResults.Updates.Count; $i++) {
            $update = $searchResults.Updates.Item($i)
            $title = $update.Title
            If ($update.IsDownloaded) {
                Write-Output "$i >  $title"
                $updatesToInstall.Add($update)
                If ($update.InstallationBehavior.RebootBehavior -gt 0) {
                    $rebootMayBeRequired = $true
                }
            }
        }

        If ($updatesToInstall.Count -eq 0) {
            Write-Output "No updates were successfully downloaded."
        }else{
            If ($rebootMayBeRequired) {
                Write-Output "These updates might require a reboot."
            }

            Write-Output "Installing updates..."
            $installer = $updateSession.CreateUpdateInstaller()
            $installer.Updates = $updatesToInstall
            $installationResult = $installer.Install()

            Write-Output "Installation Result: " $installationResult.ResultCode
            Write-Output "Reboot Required: " $rebootMayBeRequired

            Write-Output "List of updates installed:"
            For ($i = 0;$i -lt $updatesToInstall.Count;$i++) {
                $title = $installer.Updates($i).Title
                $installResult = $installationResult.GetUpdateResult($i).ResultCode
                Write-Output "$i > $title : $installResult"
            }
        }
    }
}

Write-Output "Complete."
