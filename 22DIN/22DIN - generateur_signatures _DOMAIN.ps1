# Nécessite que le roaming soit désactivé :
# Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater:$true 

Connect-ExchangeOnline

# Cible le ou les utilisateurs concernés
$mailboxes = Get-ExoMailBox -Filter {UserPrincipalName -like "*@mldinan.fr" -and RecipientTypeDetails -eq 'UserMailbox' -and CustomAttribute15 -eq "22DIN"} | Select-Object UserPrincipalName

# Chemin vers le template HTML
$templateSignatureHTML = Get-Content -Path "C:\Users\VincentMARIE\OneDrive - ARMLB\Documents\WindowsPowerShell\Scripts\Pack Signature HTML\Signatures\22DIN-template-signature.html" -raw

# Boucle pour chaque utilisateur
foreach ($mailbox in $mailboxes) { 
    $user = Get-User -Identity $mailbox.UserPrincipalName | Select-Object FirstName, LastName, Title, Phone, MobilePhone, UserPrincipalName, StreetAddress, PostalCode, City, Office, Company
    $signatureHTML = $templateSignatureHTML 
    # Vérification qu'il s'agit bien d'un utilisateur
    if ($user.FirstName) { 
        # Réécriture de l'adresse pour harmonisation
        if ($user.Company -eq "Mission Locale du Pays de Dinan") 
            {

            $building = "Immeuble .."
            $address = "Mission Locale du Pays de Dinan"
            $street = "5 Rue Gambetta"
            $postalcode = "22100"
            $city = "Dinan"
            $phone = "02 96 85 32 67" 

            Write-Host "Utilisateur trouvé"
        } else {
            Write-Host ("Erreur, aucune adresse ne correspond pour : {0} {1}" -f $user.FirstName, $user.LastName)
            continue
        }
        # Remplacement des tags dans le template par les valeurs correspondantes
        $signatureHTML = $signatureHTML.Replace("{First name}", $user.FirstName) 
        $signatureHTML = $signatureHTML.Replace("{Last name}", $user.LastName) 
        $signatureHTML = $signatureHTML.Replace("{Title}", $user.Title) 
        $signatureHTML = $signatureHTML.Replace("{Address}", $address)
        $signatureHTML = $signatureHTML.Replace("{Building}", $building)
        $signatureHTML = $signatureHTML.Replace("{Street}", $street) 
        $signatureHTML = $signatureHTML.Replace("{PostalCode}", $postalcode) 
        $signatureHTML = $signatureHTML.Replace("{City}", $city)  
        $signatureHTML = $signatureHTML.Replace("{Phone}", $phone)  
        $signatureHTML = $signatureHTML.Replace("{MobilePhone}", $user.MobilePhone)

        Write-Host ("Mise en place de la signature de : {0} {1}" -f $user.FirstName, $user.LastName)

        # Mise en place de la signature sur le compte
        Set-MailboxMessageConfiguration -Identity $user.UserPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true 
    }
}

Disconnect-ExchangeOnline
