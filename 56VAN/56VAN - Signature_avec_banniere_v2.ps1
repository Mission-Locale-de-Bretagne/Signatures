# -------------------------------
# PARAMÈTRES
# -------------------------------
$CsvPath = "C:\Signature_Vannes_Occurences.csv"

# RAW GitHub
$BaseSignatureUrl = "https://raw.githubusercontent.com/Mission-Locale-de-Bretagne/Signatures/main/56VAN/SIGNATURE/"
$BannerUrl = "https://raw.githubusercontent.com/Mission-Locale-de-Bretagne/Signatures/main/56VAN/COU_DE_JEUNES.jpg"

# -------------------------------
# CONNEXION
# -------------------------------
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# -------------------------------
# LECTURE CSV
# -------------------------------
$Users = Import-Csv -Path $CsvPath -Delimiter ","

foreach ($User in $Users) {

    $UPN = $User.UPN
    $SignatureFile = $User.ImageFile

    if ([string]::IsNullOrWhiteSpace($UPN)) {
        Write-Host "⚠️ UPN vide ignoré" -ForegroundColor Yellow
        continue
    }

    $SignatureUrl = $BaseSignatureUrl + $SignatureFile

    # -------------------------------
    # HTML SIGNATURE COMPLET
    # -------------------------------
    $SignatureHtml = @"
<div style="font-family:Calibri; font-size:11px;">
    
    <img src="$SignatureUrl" style="max-width:600px; max-height:300px;"><br>
    
    <img src="$BannerUrl" style="max-width:600px;"><br><br>

    <strong>La Mission Locale du Pays de Vannes protège vos données&nbsp;!</strong><br>
    Vous recevez ce mail en tant que personne enregistrée dans notre base de contact.<br>
    En cas d'erreur de destinataire ou pour toute demande visant à assurer l'exercice de vos droits
    sur vos données personnelles, merci de vous adresser à notre délégué(e) à la protection des données :
    <a href="mailto:rgpd@mlpv.org">rgpd@mlpv.org</a>.<br>
    Vous pouvez aussi consulter notre
    <a href="https://www.mlpvannes.org/14949-2/" target="_blank">politique de confidentialité</a>.

</div>
"@

    # -------------------------------
    # APPLICATION
    # -------------------------------
    try {
        Set-MailboxMessageConfiguration -Identity $UPN `
            -SignatureHtml $SignatureHtml `
            -AutoAddSignature $true `
            -AutoAddSignatureOnReply $true

        Write-Host "✅ Signature OK : $UPN" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Erreur pour $UPN : $_" -ForegroundColor Red
    }
}

# -------------------------------
# DECONNEXION
# -------------------------------
Disconnect-ExchangeOnline -Confirm:$false
