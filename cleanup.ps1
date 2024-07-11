# Définir les chemins des dossiers à nettoyer
$tempPaths = @(
    "$env:windir\Temp",
    "$env:USERPROFILE\AppData\Local\Temp",
    "$env:USERPROFILE\AppData\Local\Microsoft\Windows\Temporary Internet Files"
)

# Définir la date limite pour les fichiers à supprimer (ici, fichiers plus vieux que 30 jours)
$limitDate = (Get-Date).AddDays(-30)

# Fonction pour enregistrer les logs
function Write-Log {
    param (
        [string]$LogMessage
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path "$env:USERPROFILE\cleanup_log.txt" -Value "$timestamp - $LogMessage"
}

# Boucle à travers chaque chemin de dossier
foreach ($path in $tempPaths) {
    if (Test-Path $path) {
        # Récupérer tous les fichiers et dossiers
        $items = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue
        foreach ($item in $items) {
			try {
				Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop
				Write-Log "Deleted: $($item.FullName)"
			} catch {
				Write-Log "Failed to delete: $($item.FullName) - $($_.Exception.Message)"
			}
            
        }
    } else {
        Write-Log "Path not found: $path"
    }
}

# Exécuter le nettoyage de disque pour des tâches supplémentaires
function Invoke-WindowsDiskCleanup {
    Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -NoNewWindow -Wait
    Write-Log "Executed Disk Cleanup"
}

Invoke-WindowsDiskCleanup

Write-Log "Cleanup operation completed."
