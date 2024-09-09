#Necessite que le roaming soit desactive :
#Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater:$true 

Connect-ExchangeOnline


# Cible le ou les utilisateurs concernÃ©s
$users = Get-User -Filter {UserPrincipalName -like "vmarie@mlfougeres.onmicrosoft.com" -and RecipientTypeDetails -eq 'UserMailbox'} | Select-Object firstname,lastname,title,phone,mobilephone,userprincipalname,streetaddress,postalcode,city,Office,Company,department
#$CompanyName = Get-AzureADUser -Filter "UserPrincipalName eq 'vmarie@mlfougeres.onmicrosoft.com'" | Select-Object -ExpandProperty CompanyName

# Chemin vers le template HTML
$templateSignatureHTML = Get-Content -Path "C:\Users\VincentMARIE\OneDrive - ARMLB\Documents\WindowsPowerShell\Scripts\Pack Signature HTML\Signatures\22DIN-PAEJ-template-signature.html" -raw

# Boucle pour chaque utilisateur
foreach ($user in $users) { 
	$signatureHTML = $templateSignatureHTML 
	# VÃ©rification qu'il s'agit bien d'un utilisateur
	if ($user.firstname) { 
		# RÃ©Ã©criture de l'adresse pour harmonisation
        if ($user.company -eq "Mission Locale du Pays de Dinan")
        ## POUR PAEJ ##
        # if ($user.department -eq "PAEJ")
        {
            #$building = "Immeuble .."
            $address = "Mission Locale du Pays de Dinan"
            $street = "5 Rue Gambetta"
            $postalcode = "22100"
            $city = "Dinan"
            $phone = "02 96 85 32 67" 

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
		$signatureHTML = $signatureHTML.Replace("{Phone}", $usephone)  
		$signatureHTML = $signatureHTML.Replace("{MobilePhone}", $user.mobilephone)

	} 
}

	Write-Host ("Mise en place de la signature de : {0} {1}" -f $user.firstname, $user.lastname)

	# Mise en place de la signature sur le compte
	Set-MailboxMessageConfiguration -Identity $user.userPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true 

Disconnect-ExchangeOnline
