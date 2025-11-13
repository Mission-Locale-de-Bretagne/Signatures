# Nécessite que le roaming soit désactivé :
# Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater:$true 

# Save executing directory to variable
$scriptDirectory = Get-Location

# Connexion à Exchange Online
Connect-ExchangeOnline

# Cible le ou les utilisateurs concernés
$mailboxes = Get-ExoMailBox -Filter {UserPrincipalName -like "*@mlpv.org" -and RecipientTypeDetails -eq 'UserMailbox' -and CustomAttribute15 -eq "56VAN"} | Select-Object UserPrincipalName

# Chemin vers le template HTML
$templateSignatureHTML = Get-Content -Path "$scriptDirectory\56VAN\56VAN-template-signature.html" -raw

# Boucle pour chaque utilisateur
foreach ($mailbox in $mailboxes) { 
    $user = Get-User -Identity $mailbox.UserPrincipalName | Select-Object FirstName, LastName, Title, Phone, MobilePhone, UserPrincipalName, StreetAddress, PostalCode, City, Office, Company
    $signatureHTML = $templateSignatureHTML 
    # Vérification qu'il s'agit bien d'un utilisateur
    if ($user.FirstName) { 
        # Réécriture de l'adresse pour harmonisation
           if ($user.company -eq "Mission Locale du Pays de Vannes")
        {
            #$building = "Immeuble .."
            $address = "Mission Locale du Pays de Vannes"
            $street = "1 Allée de Kérivarho"
            $postalcode = "56000"
            $city = "Vannes"
            $phone = "02 97 01 65 40" 

            Write-Host "Utilisateur trouvé"
        } else {
            Write-Host ("Erreur, aucune adresse ne correspond pour : {0} {1}" -f $user.firstname, $user.lastname)
            continue
        }
        # Remplacement des tags dans le template par les valeurs correspondantes
		$signatureHTML = $signatureHTML.Replace("{First name}", $user.firstname) 
		$signatureHTML = $signatureHTML.Replace("{Last name}", $user.lastname) 
		$signatureHTML = $signatureHTML.Replace("{Title}", $user.title) 
		$signatureHTML = $signatureHTML.Replace("{Address}", $user.Company)
        $SignatureHTML = $signatureHTML.Replace("{Building}",$building)
		$signatureHTML = $signatureHTML.Replace("{Street}", $user.StreetAddress) 
		$signatureHTML = $signatureHTML.Replace("{PostalCode}", $user.PostalCode) 
		$signatureHTML = $signatureHTML.Replace("{City}", $user.City)  
		$signatureHTML = $signatureHTML.Replace("{Phone}", $user.phone)  
		$signatureHTML = $signatureHTML.Replace("{MobilePhone}", $user.mobilephone)

        Write-Host ("Mise en place de la signature de : {0} {1}" -f $user.firstname, $user.lastname)

        # Mise en place de la signature sur le compte
        Set-MailboxMessageConfiguration -Identity $user.userPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true 
        }
    }

Disconnect-ExchangeOnline -Confirm:$false
