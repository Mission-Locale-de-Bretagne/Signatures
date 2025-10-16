﻿# Save executing directory to variable
$scriptDirectory = Get-Location

# Stockage dans une variable de l'UPN de l'utilisateur
$userUPN = Read-Host "Saisir l'UPN de l'utilisateur"

# Connexion à Exchange Online
Connect-ExchangeOnline

# Cible le ou les utilisateurs concernés
$users = Get-User $userUPN | Select-Object firstname,lastname,title,phone,mobilephone,userprincipalname,streetaddress,postalcode,city,office,company

# Chemin vers le template HTML
$templateSignatureHTML = Get-Content -Path "$scriptDirectory\35STM\35STM-template-signature.html" -raw

# Boucle pour chaque utilisateur
foreach ($user in $users) { 
	$signatureHTML = $templateSignatureHTML 
	# Vérification qu'il s'agit bien d'un utilisateur
	if ($user.firstname) { 
		# Réécriture de l'adresse pour harmonisation
        if ($user.Company -eq "Mission Locale du Pays de Saint-Malo")
        {
            $address = "Mission Locale du Pays de Saint-Malo"
            $street = "35 Avenue Comptoirs"
            $postalcode = "35400"
            $city = "Saint-Malo"
            $phone = "02 99 82 86 00 "
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
		$signatureHTML = $signatureHTML.Replace("{Street}", $user.streetaddress) 
		$signatureHTML = $signatureHTML.Replace("{PostalCode}", $user.postalcode) 
		$signatureHTML = $signatureHTML.Replace("{City}", $user.city)  
		$signatureHTML = $signatureHTML.Replace("{Phone}", $user.phone)  
		$signatureHTML = $signatureHTML.Replace("{MobilePhone}", $user.mobilephone)

	} 
}

	Write-Host ("Mise en place de la signature de : {0} {1}" -f $user.firstname, $user.lastname)

	# Mise en place de la signature sur le compte
	Set-MailboxMessageConfiguration -Identity $user.userPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true 

Disconnect-ExchangeOnline -Confirm:$false
