# Chemin vers le fichier ospp.vbs
$osppVbsPath = "C:\Program Files\Microsoft Office\Office16\ospp.vbs"

# Fonction pour vérifier le statut de la licence
function Check-License {
    Write-Output "Vérification du statut de la licence..."
    $output = cscript $osppVbsPath /dstatus
    Write-Output "Résultat de la vérification du statut :"
    Write-Output $output
    if ($output -match "LICENSE STATUS:  ---LICENSED---") {
        Write-Output "Office est déjà licencié."
        return $true
    } else {
        Write-Output "Office n'est pas licencié."
        return $false
    }
}

# Fonction pour essayer d'activer la licence avec un serveur KMS donné
function Try-Activation {
    param (
        [string]$kmsServer
    )
    Write-Output "Tentative d'activation avec $kmsServer..."
    $sethstOutput = cscript $osppVbsPath /sethst:$kmsServer
    Write-Output "Résultat de la commande /sethst :"
    Write-Output $sethstOutput
    $actOutput = cscript $osppVbsPath /act
    Write-Output "Résultat de la commande /act :"
    Write-Output $actOutput

    # Vérifier à nouveau le statut de la licence
    Write-Output "Vérification du statut de la licence après activation avec $kmsServer..."
    $output = cscript $osppVbsPath /dstatus
    Write-Output "Résultat de la vérification du statut après activation :"
    Write-Output $output
    if ($output -match "LICENSE STATUS:  ---LICENSED---") {
        Write-Output "Activation réussie avec $kmsServer."
        return $true
    } else {
        Write-Output "Activation échouée avec $kmsServer."
        return $false
    }
}

# Script principal
Write-Output "Début du script d'activation d'Office"
if (-not (Check-License)) {
    Write-Output "Tentatives d'activation requises."
    if (Try-Activation -kmsServer "kms8.MSGuides.com") { exit }
    if (Try-Activation -kmsServer "kms.03k.org") { exit }
    if (Try-Activation -kmsServer "kms9.MSGuides.com") { exit }
    Write-Output "Échec de l'activation avec tous les serveurs KMS."
} else {
    Write-Output "Office est déjà licencié, aucune activation nécessaire."
	exit 0
}

Write-Output "Fin du script."
exit 0
