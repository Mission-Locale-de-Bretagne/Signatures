#Liste des utilisateurs FSE, disposant d'un template différent avec un logo en plus.
$fseUsers = @(
    "elise.bocquel@ml-redon.com"
)

#Définition de la variable du répertoire d'exécution du script
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = Split-Path -Path $scriptPath -Parent

# Connexion à Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$true

# Input dans une variable de l'UPN de l'utilisateur
$userUPN = Read-Host "Saisir l'UPN de l'utilisateur"
# Cible le ou les utilisateurs concernés
$users = Get-User $userUPN | Select-Object firstname,lastname,title,phone,mobilephone,userprincipalname,streetaddress,postalcode,city,office,company


# Boucle pour chaque utilisateur
foreach ($user in $users) { 
	
	# Verification qu'il s'agit bien d'un utilisateur
	if ($user.firstname) {
		
		# Modification du template pour les utilisateurs FSE si le UPN de l'utilisateur est dans la liste $fseUsers
        if ($fseUsers -contains $user.UserPrincipalName) {
			$signatureHTML = Get-Content -Path "$scriptDirectory\35RED-template-signature_FSE.html" -raw
            $signatureTemplate = "FSE Redon" 
        } else {	
			$signatureHTML = Get-Content -Path "$scriptDirectory\35RED-template-signature.html" -raw
			$signatureTemplate = "Standard Redon"
		}

		# Output si l'utilisateur est trouvé ou non
        if ($user.Company -eq "Mission Locale du Pays de Redon et de Vilaine")
        {
            Write-Host "Utilisateur trouvé"
			
        } else {
			Write-Host ("Erreur, aucune adresse ne correspond pour : {0} {1}" -f $user.firstname, $user.lastname)
			exit
        } 
		
		# Remplacement des tags dans le template par les valeurs correspondantes
		$signatureHTML = $signatureHTML.Replace("{First name}", $user.firstname) 
		$signatureHTML = $signatureHTML.Replace("{Last name}", $user.lastname) 
		$signatureHTML = $signatureHTML.Replace("{Title}", $user.title) 
		$signatureHTML = $signatureHTML.Replace("{Address}", $address)
        $SignatureHTML = $signatureHTML.Replace("{Building}",$building)
		$signatureHTML = $signatureHTML.Replace("{Street}", $street) 
		$signatureHTML = $signatureHTML.Replace("{PostalCode}", $postalcode) 
		$signatureHTML = $signatureHTML.Replace("{City}", $city)  
		$signatureHTML = $signatureHTML.Replace("{Phone}", $phone)  
		$signatureHTML = $signatureHTML.Replace("{MobilePhone}", $user.mobilephone)
		$signatureHTML = $signatureHTML.Replace("{Mail}", $user.userPrincipalName)
		
		# Suppression de la ligne Mobile si aucun numéro n'est renseigné
		if ([string]::IsNullOrWhiteSpace($user.MobilePhone)) {
		    $signatureHTML = $signatureHTML -replace '(?m)^\s*Mobile\s*:.*<br>\s*\r?\n?', ''
		}
	} 
}

    # Output de l'utilisateur et du template utilisé
	Write-Host ("Mise en place de la signature de : {0} {1}" -f $user.firstname, $user.lastname)
	Write-Host "Template utilisé : $($signatureTemplate)"

	# Mise en place de la signature sur le compte
	Set-MailboxMessageConfiguration -Identity $users.userPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true 

Disconnect-ExchangeOnline -Confirm:$false