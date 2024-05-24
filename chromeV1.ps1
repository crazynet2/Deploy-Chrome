# URL to download the latest version of Google Chrome
$chromeUrl = "https://dl.google.com/chrome/install/GoogleChromeStandaloneEnterprise64.msi"

# Local path where the Google Chrome installer will be saved
$chromeInstallerPath = "e:\Temp\GoogleChromeStandaloneEnterprise64.msi"

# Download the latest version of Google Chrome
Invoke-WebRequest -Uri $chromeUrl -OutFile $chromeInstallerPath

# Get the list of computers in the domain, excluding the local computer
$computers = Get-ADComputer -Filter * | Where-Object { $_.Name -ne $env:COMPUTERNAME } | Select-Object -ExpandProperty Name

# Loop through every computer in the domain
foreach ($computer in $computers) {
    # Check if the computer is reachable
    if (Test-Connection -ComputerName $computer -Count 1 -Quiet) {
        try {
            # Copy the Chrome installation file to the remote computer
            Copy-Item -Path $chromeInstallerPath -Destination "\\$computer\c$\" -Force
            
            # Run Chrome Setup using WinRM
            Invoke-Command -ComputerName $computer -ScriptBlock {
                Start-Process -FilePath "msiexec.exe" -ArgumentList "/i C:\GoogleChromeStandaloneEnterprise64.msi /qn /norestart" -Wait
            }
            
            Write-Host "Aggiornamento di Google Chrome su $computer eseguito."

            # Send a popup to Chrome to notify you of the update using msg.exe
            Invoke-Command -ComputerName $computer -ScriptBlock {
                $msg = "Chrome will automatically restart in 15 minutes to complete the update."
                msg * /SERVER:$env:COMPUTERNAME "$msg"
            }
        } catch {
            Write-Host "Error updating Google Chrome su ${computer}: An error occurred during the update."
        }
    } else {
        Write-Host "The computer $computer is not reachable."
    }
}
