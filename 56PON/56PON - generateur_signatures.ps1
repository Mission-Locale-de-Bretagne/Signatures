﻿#Necessite que le roaming soit desactive :
#Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater:$true 

Connect-ExchangeOnline


# Cible le ou les utilisateurs concernÃ©s
$users = Get-User -Filter {UserPrincipalName -like "vmarie@mlfougeres.onmicrosoft.com" -and RecipientTypeDetails -eq 'UserMailbox'} | Select-Object firstname,lastname,title,phone,mobilephone,userprincipalname,streetaddress,postalcode,city,Office,Company
#$CompanyName = Get-AzureADUser -Filter "UserPrincipalName eq 'vmarie@mlfougeres.onmicrosoft.com'" | Select-Object -ExpandProperty CompanyName

# Chemin vers le template HTML
$templateSignatureHTML = Get-Content -Path "C:\Users\VincentMARIE\OneDrive - ARMLB\Documents\WindowsPowerShell\Scripts\Signatures\56PON-template-signature.html" -raw

# Boucle pour chaque utilisateur
foreach ($user in $users) { 
	$signatureHTML = $templateSignatureHTML 
	# VÃ©rification qu'il s'agit bien d'un utilisateur
	if ($user.firstname) { 
		# RÃ©Ã©criture de l'adresse pour harmonisation
        if ($user.company -eq "Mission Locale Centre Bretagne")
        {
            #$building = "Immeuble .."
            $address = "Mission Locale Centre Bretagne"
            $street = "13 bis Rue Saint-Jory"
            $postalcode = "56300"
            $city = "Pontivy"
            $phone = "02 97 25 38 35" 
            $street2 = "1 Rue de la Chesnaie"
            $postalCode2 ="22600"
            $city2 = "Loudéac"
            $phone2 = "02 96 28 99 18"

            Write-Host "Utilisateur trouvé :"

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
		$signatureHTML = $signatureHTML.Replace("{Street2}", $street2) 
		$signatureHTML = $signatureHTML.Replace("{PostalCode2}", $postalcode2) 
		$signatureHTML = $signatureHTML.Replace("{City2}", $city2)  
		$signatureHTML = $signatureHTML.Replace("{Phone2}", $phone2) 

	} 
}

	Write-Host ("Mise en place de la signature de : {0} {1}" -f $user.firstname, $user.lastname)

	# Mise en place de la signature sur le compte
	Set-MailboxMessageConfiguration -Identity $user.userPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true 

Disconnect-ExchangeOnline
