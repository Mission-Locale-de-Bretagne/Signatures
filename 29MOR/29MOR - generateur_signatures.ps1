#Necessite que le roaming soit desactive :
#Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater:$true 

#Définition de la variable du répertoire d'exécution du script
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = Split-Path -Path $scriptPath -Parent

# Connexion à Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$true

## Input dans une variable de l'UPN de l'utilisateur
$userUPN = Read-Host "Saisir l'UPN de l'utilisateur"
# Cible le ou les utilisateurs concernés
$users = Get-User $userUPN | Select-Object firstname,lastname,title,phone,mobilephone,userprincipalname,streetaddress,postalcode,city,office,company

# Chemin vers le template HTML
$templateSignatureHTML = Get-Content -Path "$scriptDirectory\29MOR-template-signature.html" -raw

# Boucle pour chaque utilisateur
foreach ($user in $users) { 
	$signatureHTML = $templateSignatureHTML 
	# VÃ©rification qu'il s'agit bien d'un utilisateur
	if ($user.firstname) { 
		# RÃ©Ã©criture de l'adresse pour harmonisation
        if ($user.company -eq "Mission Locale du Pays de Morlaix")
        {
            #$building = "Immeuble .."
            $address = "Mission Locale du Pays de Morlaix"
            $street = "Rue Jean Caerou - ZA La Boissière"
            $postalcode = "29600"
            $city = "Morlaix"
            $phone = "02 98 15 15 50" 

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
		$signatureHTML = $signatureHTML.Replace("{Street}", $street) 
		$signatureHTML = $signatureHTML.Replace("{PostalCode}", $postalcode) 
		$signatureHTML = $signatureHTML.Replace("{City}", $city)  
		$signatureHTML = $signatureHTML.Replace("{Phone}", $phone)  
		$signatureHTML = $signatureHTML.Replace("{MobilePhone}", $user.mobilephone)

	} 
}

	Write-Host ("Mise en place de la signature de : {0} {1}" -f $user.firstname, $user.lastname)

	# Mise en place de la signature sur le compte
	Set-MailboxMessageConfiguration -Identity $user.userPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true 

Disconnect-ExchangeOnline -Confirm:$false
