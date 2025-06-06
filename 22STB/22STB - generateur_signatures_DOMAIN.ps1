# Nécessite que le roaming soit desactivé :
# Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater:$true 

# Nécessite PowerShell 5.1
# Nécessite le module ExchangeOnlineManagement ver. = 3.5.1

Connect-ExchangeOnline

# Cible le ou les utilisateurs concernés
$mailboxes = Get-ExoMailBox -Filter {UserPrincipalName -like "*@mlstbrieuc.fr" -and RecipientTypeDetails -eq 'UserMailbox' -and CustomAttribute15 -eq "22STB"} | Select-Object UserPrincipalName

# Chemin vers le template HTML
$templateSignatureHTML = Get-Content -Path "<CHEMIN_VERS_TEMPLATE>\22STB-template-signature.html" -raw

# Boucle pour chaque utilisateur
foreach ($mailbox in $mailboxes) { 
    $user = Get-User -Identity $mailbox.UserPrincipalName | Select-Object FirstName, LastName, Title, Phone, MobilePhone, UserPrincipalName, StreetAddress, PostalCode, City, Office, Company
    $signatureHTML = $templateSignatureHTML 
    # Vérification qu'il s'agit bien d'un utilisateur
    if ($user.FirstName) { 
        # Réécriture de l'adresse pour harmonisation
        if ($user.Company -eq "Mission Locale du Pays de Saint-Brieuc")
        {

            #$building = "Immeuble .."
            $address = "Mission Locale du Pays de Saint-Brieuc"
            $street = "47 Rue du docteur Rahuel"
            $postalcode = "22000"
            $city = "Saint-Brieuc"
            $phone = "02 96 68 15 68" 

            Write-Host "Utilisateur trouvé"
        } else {
            Write-Host ("Erreur, aucune adresse ne correspond pour : {0} {1}" -f $user.FirstName, $user.LastName)
			continue
        }
		# Remplacement des tags dans le template par les valeurs correspondantes
		$signatureHTML = $signatureHTML.Replace("{First name}", $user.firstname) 
		$signatureHTML = $signatureHTML.Replace("{Last name}", $user.lastname) 
		$signatureHTML = $signatureHTML.Replace("{Title}", $user.title) 
		$signatureHTML = $signatureHTML.Replace("{Address}", $address)
        	#$SignatureHTML = $signatureHTML.Replace("{Building}",$building)
		$signatureHTML = $signatureHTML.Replace("{Street}", $user.streetaddress) 
		$signatureHTML = $signatureHTML.Replace("{PostalCode}", $user.postalcode) 
		$signatureHTML = $signatureHTML.Replace("{City}", $user.city)  
		$signatureHTML = $signatureHTML.Replace("{Phone}", $user.phone)  
		$signatureHTML = $signatureHTML.Replace("{MobilePhone}", $user.mobilephone)

	    Write-Host ("Mise en place de la signature de : {0} {1}" -f $user.FirstName, $user.LastName)

        # Mise en place de la signature sur le compte
        Set-MailboxMessageConfiguration -Identity $user.UserPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true 
    }
}

Disconnect-ExchangeOnline -Confirm:$false
