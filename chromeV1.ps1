# URL per il download dell'ultima versione di Google Chrome
$chromeUrl = "https://dl.google.com/chrome/install/GoogleChromeStandaloneEnterprise64.msi"

# Percorso locale dove verrà salvato l'installer di Google Chrome
$chromeInstallerPath = "e:\Temp\GoogleChromeStandaloneEnterprise64.msi"

# Scarica l'ultima versione di Google Chrome
Invoke-WebRequest -Uri $chromeUrl -OutFile $chromeInstallerPath

# Ottieni l'elenco dei computer nel dominio, escludendo il computer locale
$computers = Get-ADComputer -Filter * | Where-Object { $_.Name -ne $env:COMPUTERNAME } | Select-Object -ExpandProperty Name

# Loop attraverso ogni computer nel dominio
foreach ($computer in $computers) {
    # Verifica se il computer è raggiungibile
    if (Test-Connection -ComputerName $computer -Count 1 -Quiet) {
        try {
            # Copia il file di installazione di Chrome sul computer remoto
            Copy-Item -Path $chromeInstallerPath -Destination "\\$computer\c$\" -Force
            
            # Esegui l'installazione di Chrome utilizzando WinRM
            Invoke-Command -ComputerName $computer -ScriptBlock {
                Start-Process -FilePath "msiexec.exe" -ArgumentList "/i C:\GoogleChromeStandaloneEnterprise64.msi /qn /norestart" -Wait
            }
            
            Write-Host "Aggiornamento di Google Chrome su $computer eseguito."

            # Invia un popup su Chrome per avvisare dell'aggiornamento utilizzando msg.exe
            Invoke-Command -ComputerName $computer -ScriptBlock {
                $msg = "Chrome si riavvierà automaticamente in 15 minuti per completare l'aggiornamento."
                msg * /SERVER:$env:COMPUTERNAME "$msg"
            }
        } catch {
            Write-Host "Errore durante l'aggiornamento di Google Chrome su ${computer}: Si è verificato un errore durante l'aggiornamento."
        }
    } else {
        Write-Host "Il computer $computer non è raggiungibile."
    }
}
