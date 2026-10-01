# Liste des utilisateurs FSE, disposant d'un template différent avec un logo en plus.
$fseUsers = @(
    "elise.bocquel@ml-redon.com"
)

# Définition de la variable du répertoire d'exécution du script
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = Split-Path -Path $scriptPath -Parent

# Connexion à Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$false

# Stockage dans la variable $mailboxes de toutes les boites mail de type UserMailbox dont le domaine est @ml-redon.com et dont l'attribut CustomAttribute15 est égal à "35RED"
$mailboxes = Get-ExoMailBox -Filter {UserPrincipalName -like "*@ml-redon.com" -and RecipientTypeDetails -eq 'UserMailbox' -and CustomAttribute15 -eq "35RED"} | Select-Object UserPrincipalName

# Boucle pour chaque utilisateur
foreach ($mailbox in $mailboxes) {
    # Récupération des informations de l'utilisateur à partir de Azure/Entra AD 
    $user = Get-User -Identity $mailbox.UserPrincipalName | Select-Object FirstName, LastName, Title, Phone, MobilePhone, UserPrincipalName, StreetAddress, PostalCode, City, Office, Company
           
    # Définition du template de signature en fonction de l'utilisateur
    # Si l'utilisateur est dans la variable $fseUsers, on utilise le template FSE, sinon on utilise le template standard.
    if ($fseUsers -contains $user.UserPrincipalName) {
        $signatureHTML = Get-Content -Path "$scriptDirectory\35RED-template-signature_FSE.html" -raw
        $signatureTemplate = "FSE Redon"            
    } else {
        $signatureHTML = Get-Content -Path "$scriptDirectory\35RED-template-signature.html" -raw
        $signatureTemplate = "Standard Redon"
    }

    # Vérification que l'utilisateur possède "Mission Locale du Pays de Redon et de Vilaine" dans la propriété "Company Name" sur Azure/Entra AD.
    if ($user.company -eq "Mission Locale du Pays de Redon et de Vilaine")
    {
        Write-Host "Utilisateur trouvé"
    } else {
        Write-Host ("Erreur, aucune adresse ne correspond pour : {0} {1}" -f $user.FirstName, $user.LastName)
        continue
    }

    # Remplacement des tags dans le template par les valeurs correspondantes tirées de Azure/Entra AD
    $signatureHTML = $signatureHTML.Replace("{First name}", $user.FirstName) 
    $signatureHTML = $signatureHTML.Replace("{Last name}", $user.LastName) 
    $signatureHTML = $signatureHTML.Replace("{Mail}", $user.userPrincipalName)
    $signatureHTML = $signatureHTML.Replace("{Title}", $user.Title) 
    
	# Suppression de la ligne Mobile si aucun numéro n'est renseigné
	if ([string]::IsNullOrWhiteSpace($user.MobilePhone)) {
        $signatureHTML = $signatureHTML -replace '(?m)^\s*Mobile\s*:.*<br>\s*\r?\n?', ''
	} else {
        $signatureHTML = $signatureHTML.Replace("{MobilePhone}", $user.MobilePhone)
    }

    # Output de l'utilisateur et du template utilisé
    Write-Host ("Mise en place de la signature de : {0} {1}" -f $user.FirstName, $user.LastName)
    Write-Host "Template utilisé : $($signatureTemplate)`n"

    # Mise en place de la signature sur le compte
    Set-MailboxMessageConfiguration -Identity $user.UserPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true 
}

# Déconnexion d'Exchange Online sans confirmation
Disconnect-ExchangeOnline -Confirm:$false