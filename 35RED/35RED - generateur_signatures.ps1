# Liste des utilisateurs FSE, disposant d'un template différent avec un logo en plus.
$fseUsers = @(
    "elise.bocquel@ml-redon.com"
)

# Définition de la variable du répertoire d'exécution du script
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = Split-Path -Path $scriptPath -Parent

# Connexion à Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$true

# Input dans une variable de l'UPN de l'utilisateur
$userUPN = Read-Host "Saisir l'UPN de l'utilisateur"
# Cible le ou les utilisateurs concernés
$user = Get-User $userUPN | Select-Object firstname,lastname,title,phone,mobilephone,userprincipalname,streetaddress,postalcode,city,office,company
		
# Modification du template pour les utilisateurs FSE si le UPN de l'utilisateur est dans la liste $fseUsers
if ($fseUsers -contains $user.UserPrincipalName) {
	$signatureHTML = Get-Content -Path "$scriptDirectory\35RED-template-signature_FSE.html" -raw
    $signatureTemplate = "FSE Redon" 
} else {	
	$signatureHTML = Get-Content -Path "$scriptDirectory\35RED-template-signature.html" -raw
	$signatureTemplate = "Standard Redon"
}

# Output si l'utilisateur a le bon Company Name, sinon sortie du script avec un message d'erreur
if ($user.Company -eq "Mission Locale du Pays de Redon et de Vilaine") {
    Write-Host "Utilisateur trouvé"
} else {
	Write-Host ("Erreur, aucune adresse ne correspond pour : {0} {1}" -f $user.firstname, $user.lastname)
	exit
} 

# Remplacement des tags dans le template par les valeurs correspondantes tirées de Azure/Entra AD
  	$signatureHTML = $signatureHTML.Replace("{First name}", $user.FirstName) 
  	$signatureHTML = $signatureHTML.Replace("{Last name}", $user.LastName) 
  	$signatureHTML = $signatureHTML.Replace("{Mail}", $user.userPrincipalName)
  	$signatureHTML = $signatureHTML.Replace("{Title}", $user.Title) 
  
# Suppression de la ligne Mobile si aucun numéro n'est renseigné
if ([string]::IsNullOrWhiteSpace($user.MobilePhone)) {
  	$signatureHTML = $signatureHTML -replace '(?m)^\s*Mobile\s*:.*<br>\s*\r?\n?', ''
} else {
	$signatureHTML = $signatureHTML.Replace("{MobilePhone}", $user.MobilePhone)
}

# Output de l'utilisateur et du template utilisé
Write-Host ("Mise en place de la signature de : {0} {1}" -f $user.firstname, $user.lastname)
Write-Host "Template utilisé : $($signatureTemplate)"

# Mise en place de la signature sur le compte
Set-MailboxMessageConfiguration -Identity $user.userPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true 

# Déconnexion d'Exchange Online sans confirmation
Disconnect-ExchangeOnline -Confirm:$false