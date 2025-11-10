#Définition de la variable du répertoire d'exécution du script
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = Split-Path -Path $scriptPath -Parent

# Connexion à Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$true

# Cible le ou les utilisateurs concernÃ©s
$users = Get-User -Filter {UserPrincipalName -like "@armlb.bzh" -and RecipientTypeDetails -eq 'UserMailbox'} | Select-Object firstname,lastname,title,phone,mobilephone,userprincipalname,streetaddress,postalcode,city,Office,Company 

# Chemin vers le template HTML
$templateSignatureHTML = Get-Content -Path "$scriptDirectory\35ARM-template-signature.html" -raw

# Boucle pour chaque utilisateur
foreach ($user in $users) { 
	$signatureHTML = $templateSignatureHTML 
	# VÃ©rification qu'il s'agit bien d'un utilisateur
	if ($user.firstname) { 
		# RÃ©Ã©criture de l'adresse pour harmonisation
        if ($user.Company -eq "Association Régionale des Missions Locales de Bretagne")
        {

            $building = "Immeuble Colbert"
            $address = "Association Régionale des Missions Locales de Bretagne"
            $street = "31 Place du Colombier"
            $postalcode = "35000"
            $city = "Rennes"
            $phone = "" 

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