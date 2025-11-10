# Nécessite que le roaming soit désactivé :
# Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater:$true 

#Définition de la variable du répertoire d'exécution du script
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = Split-Path -Path $scriptPath -Parent

# Connexion à Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$true

# Cible le ou les utilisateurs concernés
$mailboxes = Get-ExoMailBox -Filter {UserPrincipalName -like "*@mllorient.org" -and RecipientTypeDetails -eq 'UserMailbox' -and CustomAttribute15 -eq "56LOR"} | Select-Object UserPrincipalName

# Chemin vers le template HTML
$templateSignatureHTML = Get-Content -Path "$scriptDirectory\56LOR-template-signature.html" -raw

# Boucle pour chaque utilisateur
foreach ($mailbox in $mailboxes) { 
    $user = Get-User -Identity $mailbox.UserPrincipalName | Select-Object FirstName, LastName, Title, Phone, MobilePhone, UserPrincipalName, StreetAddress, PostalCode, City, Office, Company
    $signatureHTML = $templateSignatureHTML 
    # Vérification qu'il s'agit bien d'un utilisateur
    if ($user.FirstName) { 
        # Réécriture de l'adresse pour harmonisation
        if ($user.company -eq "Mission Locale du Pays de Lorient")
        {
            $address = "Mission Locale du Pays de Lorient"
            $building = "Gare de Lorient"
            $street = "9bis Place François Mitterrand"
            $postalcode = "56100"
            $city = "Lorient"
            $phone = "02 97 21 42 05" 

            Write-Host "Utilisateur trouvé"
        } else {
            Write-Host ("Erreur, aucune adresse ne correspond pour : {0} {1}" -f $user.firstname, $user.lastname)
            continue
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

        Write-Host ("Mise en place de la signature de : {0} {1}" -f $user.firstname, $user.lastname)

        # Mise en place de la signature sur le compte
        Set-MailboxMessageConfiguration -Identity $user.userPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true 
        }
    }

Disconnect-ExchangeOnline -confirm:$false
