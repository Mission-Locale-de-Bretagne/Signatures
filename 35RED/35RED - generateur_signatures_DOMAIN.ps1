﻿#Liste des utilisateurs FSE, disposant d'un template différent avec un logo en plus.
$fseUsers = @(
    "elise.bocquel@ml-redon.com"
)

#Définition de la variable du répertoire d'exécution du script
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = Split-Path -Path $scriptPath -Parent

# Connexion à Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$false

# Cible le ou les utilisateurs concernés
$mailboxes = Get-ExoMailBox -Filter {UserPrincipalName -like "*@ml-redon.com" -and RecipientTypeDetails -eq 'UserMailbox' -and CustomAttribute15 -eq "35RED"} | Select-Object UserPrincipalName

# Chemin vers le template HTML
$templateSignatureHTML = Get-Content -Path "$scriptDirectory\35RED-template-signature.html" -raw

# Boucle pour chaque utilisateur
foreach ($mailbox in $mailboxes) { 
    $user = Get-User -Identity $mailbox.UserPrincipalName | Select-Object FirstName, LastName, Title, Phone, MobilePhone, UserPrincipalName, StreetAddress, PostalCode, City, Office, Company
    $signatureHTML = $templateSignatureHTML 
    # Vérification qu'il s'agit bien d'un utilisateur
    if ($user.FirstName) {
        
        # Modification du template pour les utilisateurs FSE
        if ($fseUsers -contains $user.UserPrincipalName) {
            $signatureHTML = Get-Content -Path "$scriptDirectory\35RED-template-signature_FSE.html" -raw            
        }

        # Réécriture de l'adresse pour harmonisation
        if ($user.company -eq "Mission Locale du Pays de Redon et de Vilaine")
        {
            #$building = "Immeuble .."
            $address = "Mission Locale du Pays de Redon et Vilaine"
            $street = "3 rue Charles Sillard - CS 60287"
            $postalcode = "35602"
            $city = "Redon"
            $phone = "02 99 72 19 50" 

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
        $signatureHTML = $signatureHTML.Replace("{Mail}", $user.userPrincipalName)

        Write-Host ("Mise en place de la signature de : {0} {1}" -f $user.FirstName, $user.LastName)

        # Mise en place de la signature sur le compte
        Set-MailboxMessageConfiguration -Identity $user.UserPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true 
    }
}

Disconnect-ExchangeOnline -Confirm:$false