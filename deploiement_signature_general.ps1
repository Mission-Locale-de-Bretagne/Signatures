#Script se lançant à la fin du script de création utilisateur se trouvant : https://github.com/Mission-Locale-de-Bretagne/ARMLB-ESI/blob/main/Scripts/35ARM%20-%20user%20Creation.ps1
#
#
#
#
#Nécessite que le roaming des signatures soit désactivé :
# Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true
param(
    [string]$UserPrincipalName
)

Connect-ExchangeOnline


# Tags exclus du déploiement
$ExcludedTags = @(
    "56VAN", # Vannes
    "35RED"  # Redon
)

# Templates GitHub
$TemplateUrls = @{
    "35ARM" = "https://raw.githubusercontent.com/Mission-Locale-de-Bretagne/Signatures/main/35ARM/35ARM-template-signature.html"
    "56AUR" = "https://raw.githubusercontent.com/Mission-Locale-de-Bretagne/Signatures/main/56AUR/56AUR-template-signature.html"
    "56PLO" = "https://raw.githubusercontent.com/Mission-Locale-de-Bretagne/Signatures/main/56PLO/56PLO-template-signature.html"
    "56PON" = "https://raw.githubusercontent.com/Mission-Locale-de-Bretagne/Signatures/main/56PON/56PON-template-signature.html"
    "35VIT" = "https://raw.githubusercontent.com/Mission-Locale-de-Bretagne/Signatures/main/35VIT/35VIT-template-signature.html"
    "35STM" = "https://raw.githubusercontent.com/Mission-Locale-de-Bretagne/Signatures/main/35STM/35STM-template-signature.html"
    "22DIN" = "https://raw.githubusercontent.com/Mission-Locale-de-Bretagne/Signatures/main/22DIN/22DIN-template-signature.html"
    "29MOR" = "https://raw.githubusercontent.com/Mission-Locale-de-Bretagne/Signatures/main/29MOR/29MOR-template-signature.html"
}

# Boîtes aux lettres utilisateur
if ($UserPrincipalName)
{
    $mailboxes = Get-EXOMailbox -Identity $UserPrincipalName
}
else
{
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
}

foreach ($mailbox in $mailboxes)
{
    try
    {
        Write-Host ""
        Write-Host "Traitement de $($mailbox.UserPrincipalName)" -ForegroundColor Cyan

        $MailboxInfo = Get-Mailbox $mailbox.UserPrincipalName
        $Tag = $MailboxInfo.CustomAttribute15

        if (:IsNullOrWhiteSpace($Tag))
        {
            Write-Warning "CustomAttribute15 vide."
            continue
        }

        # Exclusions
        if ($Tag -in $ExcludedTags)
        {
            Write-Host "Signature non gérée pour $Tag." -ForegroundColor Yellow
            continue
        }

        # Vérification template
        if (-not $TemplateUrls.ContainsKey($Tag))
        {
            Write-Warning "Aucun template trouvé pour le tag $Tag."
            continue
        }

        $TemplateURL = $TemplateUrls[$Tag]

        Write-Host "Téléchargement du template $Tag..." -ForegroundColor Gray

        $TemplateSignatureHTML = Invoke-RestMethod -Uri $TemplateURL

        $User = Get-User -Identity $mailbox.UserPrincipalName |
            Select-Object FirstName,
                          LastName,
                          Title,
                          Phone,
                          MobilePhone,
                          UserPrincipalName,
                          StreetAddress,
                          PostalCode,
                          City,
                          Company

        if (-not $User.FirstName)
        {
            Write-Warning "Utilisateur introuvable."
            continue
        }

        $SignatureHTML = $TemplateSignatureHTML

        # Remplacement des variables
        $SignatureHTML = $SignatureHTML.Replace("{First name}", $User.FirstName)
        $SignatureHTML = $SignatureHTML.Replace("{Last name}", $User.LastName)
        $SignatureHTML = $SignatureHTML.Replace("{Title}", $User.Title)
        $SignatureHTML = $SignatureHTML.Replace("{Address}", $User.Company)
        $SignatureHTML = $SignatureHTML.Replace("{Street}", $User.StreetAddress)
        $SignatureHTML = $SignatureHTML.Replace("{PostalCode}", $User.PostalCode)
        $SignatureHTML = $SignatureHTML.Replace("{City}", $User.City)
        $SignatureHTML = $SignatureHTML.Replace("{Phone}", $User.Phone)
        $SignatureHTML = $SignatureHTML.Replace("{MobilePhone}", $User.MobilePhone)

        Write-Host "Application de la signature..." -ForegroundColor Green

        Set-MailboxMessageConfiguration `
            -Identity $User.UserPrincipalName `
            -SignatureHTML $SignatureHTML `
            -AutoAddSignature $true `
            -AutoAddSignatureOnReply $true

        Write-Host "Signature appliquée avec succès." -ForegroundColor Green
    }
    catch
    {
        Write-Warning "Erreur pour $($mailbox.UserPrincipalName) : $_"
    }
}

Disconnect-ExchangeOnline -Confirm:$false
