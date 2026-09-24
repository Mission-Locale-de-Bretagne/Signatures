Connect-ExchangeOnline

# Configuration GitHub
$GitHubBaseUrl = "https://raw.githubusercontent.com/Mission-Locale-de-Bretagne/Signatures/main/35VIT"

# CSV des correspondances
$CsvUrl = "$GitHubBaseUrl/occurences_vitre.csv"

# Mapping Hexa -> Template
$TemplateMap = @{
    "Cap_Jeune" = "35VIT-template-signature-Cap-jeune.html"
    "Normal" = "35VIT-template-signature.html"
    "Direction" = "35VIT-template-signature-dir.html"
}

# Téléchargement du CSV
try {
    $Occurrences = Import-Csv -Path (
        New-TemporaryFile | ForEach-Object {
            Invoke-WebRequest -Uri $CsvUrl -OutFile $_.FullName
            $_.FullName
        }
    ) -Delimiter ";"
}
catch {
    Write-Error "Impossible de télécharger le fichier des occurrences."
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Cible des boîtes
$mailboxes = Get-EXOMailbox `
    -Filter {UserPrincipalName -like "*@missionlocale-portedebretagne.fr" -and RecipientTypeDetails -eq 'UserMailbox'} |
    Select-Object UserPrincipalName

foreach ($mailbox in $mailboxes) {

    Write-Host "Traitement : $($mailbox.UserPrincipalName)" -ForegroundColor Cyan

    try {

        $user = Get-User -Identity $mailbox.UserPrincipalName |
            Select-Object FirstName,
                          LastName,
                          Title,
                          Phone,
                          MobilePhone,
                          UserPrincipalName,
                          StreetAddress,
                          PostalCode,
                          City,
                          Office,
                          Company

        if (-not $user.FirstName) {
            Write-Warning "Utilisateur non trouvé."
            continue
        }

        # Recherche du code Hexa dans le CSV
        $Occurrence = $Occurrences | Where-Object {
            $_.Mail -eq $user.UserPrincipalName
        }

        if (-not $Occurrence) {
            Write-Warning "Aucune occurrence trouvée pour $($user.UserPrincipalName)"
            continue
        }

        $Hexa = $Occurrence.Hexa.Trim()

        # Détermination du template
        $TemplateFile = $TemplateMap[$Hexa]

        if (-not $TemplateFile) {
            Write-Warning "Aucun template associé au code $Hexa"
            continue
        }

        # Téléchargement du template HTML
        $TemplateUrl = "$GitHubBaseUrl/$TemplateFile"

        try {
            $TemplateSignatureHTML = (Invoke-WebRequest -Uri $TemplateUrl).Content
        }
        catch {
            Write-Warning "Impossible de télécharger $TemplateFile"
            continue
        }

        $SignatureHTML = $TemplateSignatureHTML

        # Données ML
        switch ($user.Company) {

            "Mission Locale Porte de Bretagne" {
                $Building   = ""
                $Address    = "Mission Locale Porte de Bretagne"
                $Street     = "9 Place du Champ de foire"
                $PostalCode = "35500"
                $City       = "Vitré"
                $Phone      = "02 99 75 18 07"
            }

            default {
                Write-Warning "Société non reconnue : $($user.Company)"
                continue
            }
        }

        # Remplacements
        if ([string]::IsNullOrWhiteSpace($user.MobilePhone)){
            $MobileBlock = ""
        }
        else {
             $MobileBlock = '<span class="mobile">Mobile</span> <span class="phone">' + $user.MobilePhone + '</span>'
            }

        $SignatureHTML = $SignatureHTML.Replace("{MobileBlock}", $MobileBlock)
        $SignatureHTML = $SignatureHTML.Replace("{First name}", $user.FirstName)
        $SignatureHTML = $SignatureHTML.Replace("{Last name}", $user.LastName)
        $SignatureHTML = $SignatureHTML.Replace("{Title}", $user.Title)
        $SignatureHTML = $SignatureHTML.Replace("{Address}", $Address)
        $SignatureHTML = $SignatureHTML.Replace("{Building}", $Building)
        $SignatureHTML = $SignatureHTML.Replace("{Street}", $Street)
        $SignatureHTML = $SignatureHTML.Replace("{PostalCode}", $PostalCode)
        $SignatureHTML = $SignatureHTML.Replace("{City}", $City)
        $SignatureHTML = $SignatureHTML.Replace("{Phone}", $Phone)
        $SignatureHTML = $SignatureHTML.Replace("{Email}", $user.UserPrincipalName)

        Write-Host "Application de la signature à $($user.UserPrincipalName)" -ForegroundColor Green

        Set-MailboxMessageConfiguration `
            -Identity $user.UserPrincipalName `
            -SignatureHtml $SignatureHTML `
            -AutoAddSignature $true `
            -AutoAddSignatureOnReply $true

    }
    catch {
        Write-Warning "Erreur sur $($mailbox.UserPrincipalName) : $_"
    }
}

Disconnect-ExchangeOnline -Confirm:$false