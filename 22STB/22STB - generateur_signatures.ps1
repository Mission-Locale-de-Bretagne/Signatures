#Necessite que le roaming soit desactive :
#Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater:$true 

Connect-ExchangeOnline


# Cible le ou les utilisateurs concernÃ©s
$users = Get-User -Filter {UserPrincipalName -like "vmarie@mlfougeres.onmicrosoft.com" -and RecipientTypeDetails -eq 'UserMailbox'} | Select-Object firstname,lastname,title,phone,mobilephone,userprincipalname,streetaddress,postalcode,city,Office,Company
#$CompanyName = Get-AzureADUser -Filter "UserPrincipalName eq 'vmarie@mlfougeres.onmicrosoft.com'" | Select-Object -ExpandProperty CompanyName

# Chemin vers le template HTML
$templateSignatureHTML = Get-Content -Path "C:\Users\VincentMARIE\OneDrive - ARMLB\Documents\WindowsPowerShell\Scripts\Signatures\Pack Signature HTML\22STB-template-signature.html" -raw

# Boucle pour chaque utilisateur
foreach ($user in $users) { 
	$signatureHTML = $templateSignatureHTML 
	# VÃ©rification qu'il s'agit bien d'un utilisateur
	if ($user.firstname) { 
		# RÃ©Ã©criture de l'adresse pour harmonisation
        if ($user.company -eq "Mission Locale du Pays de Saint-Brieuc")
        {
            #$building = "Immeuble .."
            $address = "Mission Locale du Pays de Saint-Brieuc"
            $street = "47 Rue du docteur Rahuel"
            $postalcode = "22000"
            $city = "Saint-Brieuc"
            $phone = "02 96 68 15 68" 

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

Disconnect-ExchangeOnline
