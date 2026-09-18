Connect-ExchangeOnline

# Configuration GitHub
$GitHubBaseUrl = "https://raw.githubusercontent.com/Mission-Locale-de-Bretagne/Signatures/main/56VAN"

# CSV des correspondances
$CsvUrl = "$GitHubBaseUrl/occurences_vannes.csv"

# Mapping Hexa -> Template
$TemplateMap = @{
    "95c122" = "signature_sesame.html"
    "2d96d3" = "signature_generique.html"
    "831f81" = "signature_admin.html"
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
    -Filter {UserPrincipalName -like "*@mlpv.org" -and RecipientTypeDetails -eq 'UserMailbox'} |
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

            "Mission Locale du Pays de Vannes" {
                $Building   = ""
                $Address    = "Mission Locale du Pays de Vannes"
                $Street     = "1 Allée de Kérivarho"
                $PostalCode = "56000"
                $City       = "Vannes"
                $Phone      = "02 97 01 65 40"
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