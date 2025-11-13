<#
.SYNOPSIS
Déploie une signature OWA (Exchange Online) contenant 3 images empilées :
- Image 1 : spécifique par utilisateur (depuis RepoSubPath, ex. 56VAN/SIGNATURE/<ImageFile>)
- Image 2 et Image 3 : communes à tous (depuis CommonImagesPath, ex. 56VAN/<fichier>)

CSV minimal (UTF-8) :
UPN,ImageFile
prenom.nom@mlpv.org,PrenomNOM.jpg

Notes :
- Les images sont chargées depuis GitHub
- Compatible PowerShell 5.1 et 7+.

Auteur : Yaël Galesne
#>

param(
    # --- Chemin CSV ---
    [string]$CsvPath = "C:\Users\YaëlGALESNE\OneDrive - ARMLB\Documents\Signatures\56VAN\Signature_Vannes_Occurences.csv",

    # --- Cible GitHub ---
    [string]$GitHubOwner  = "Mission-Locale-de-Bretagne",
    [string]$GitHubRepo   = "Signatures",
    [string]$GitBranch    = "main",

    # Dossier des images "utilisateur" (Image 1) dans le repo
    [string]$RepoSubPath  = "56VAN/SIGNATURE",

    # Dossier des images "communes" (Image 2 et 3) dans le repo
    [string]$CommonImagesPath = "56VAN",

    # Noms des images communes
    [string]$CommonImage2File = "SEMAINE_INTERIM.png",
    [string]$CommonImage3File = "PLF.jpg",

    # --- OWA ---
    [bool]$AutoAddSignature        = $true,
    [bool]$AutoAddSignatureOnReply = $true,

    # --- Contrôles & journal ---
    [bool]$FailIfUrlNotFound = $true,
    [string]$CacheBustingTag = "v=2025-11-01",
    [string]$OutputLogCsv    = ".\SignatureDeployment-OWA-GitHub-3images-$((Get-Date).ToString('yyyyMMdd-HHmmss')).csv",
    [switch]$WhatIf,
    [string]$TestUPN = "",

    # --- Rendu visuel ---
    [string]$AltTextImage1 = "Signature",
    [string]$AltTextImage2 = "Semaine intérim",
    [string]$AltTextImage3 = "PLF 2025",
    [int]$MaxWidthPx = 720
)

# ===================== Préparatifs =====================
$ErrorActionPreference = 'Stop'
Import-Module ExchangeOnlineManagement -ErrorAction Stop

function Write-Info($m){ Write-Host $m -ForegroundColor Cyan }
function Write-Ok($m){ Write-Host $m -ForegroundColor Green }
function Write-Warn2($m){ Write-Warning $m }
function Write-Err2($m){ Write-Error $m }

# CSV auto-delimiter (',' ou ';') – PS 5.1 OK
function Import-CsvAuto {
    param([string]$Path)
    $firstLine = Get-Content -Path $Path -TotalCount 1 -Encoding UTF8
    $semicolon = ($firstLine -split ';').Count
    $comma     = ($firstLine -split ',').Count
    $delim = ','; if ($semicolon -gt $comma) { $delim = ';' }
    Import-Csv -Path $Path -Delimiter $delim
}

function Encode-UrlSegment([string]$s){ [System.Uri]::EscapeDataString($s) }

function Get-GitHubRawUrl {
    param(
        [string]$Owner,[string]$Repo,[string]$Branch,
        [string]$RepoSubPath,[string]$ImageFile,[string]$CacheTag
    )
    $base = "https://raw.githubusercontent.com"
    $parts = @($Owner,$Repo,$Branch)
    if ($RepoSubPath) { $parts += $RepoSubPath.Split('/').Where({$_}) }
    if ($ImageFile)   { $parts += $ImageFile.Split('/').Where({$_}) }
    $encoded = $parts | ForEach-Object { Encode-UrlSegment $_ }
    $url = $base + "/" + ($encoded -join '/')
    if ($CacheTag) {
        if ($url -match '\?') { $url += "&$CacheTag" } else { $url += "?$CacheTag" }
    }
    $url
}

function Test-UrlOk([string]$Url){
    try {
        $r = Invoke-WebRequest -Uri $Url -Method Head -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
        return ($r.StatusCode -ge 200 -and $r.StatusCode -lt 300)
    } catch { return $false }
}

# ===================== Vérifs préalables =====================
if (-not (Test-Path -LiteralPath $CsvPath)) { Write-Err2 "CSV introuvable : $CsvPath"; return }

Write-Info "Connexion à Exchange Online..."
Connect-ExchangeOnline -ShowBanner:$false

# Optionnel : si tu veux empêcher la synchro des signatures roaming
# Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater:$true

# Chargement CSV
$rows = Import-CsvAuto -Path $CsvPath
if ($rows.Count -eq 0) { Write-Err2 "CSV vide."; Disconnect-ExchangeOnline -Confirm:$false; return }

# Colonnes minimales
$cols = $rows[0].PSObject.Properties.Name
foreach ($c in @('UPN','ImageFile')) {
    if ($c -notin $cols) {
        Write-Err2 "Colonnes requises manquantes (UPN, ImageFile). Trouvées : $($cols -join ', ')"
        Disconnect-ExchangeOnline -Confirm:$false; return
    }
}

# Filtre test unitaire
if ($TestUPN) {
    $rows = $rows | Where-Object { $_.UPN -eq $TestUPN }
    if ($rows.Count -eq 0) {
        Write-Warn2 "UPN '$TestUPN' absent du CSV. Aucun traitement."
        Disconnect-ExchangeOnline -Confirm:$false; return
    }
    Write-Info "Mode test: traitement limité à '$TestUPN' (lignes: $($rows.Count))."
}

# Journal mémoire
$results = New-Object System.Collections.Generic.List[object]

# ===================== Traitement =====================
$idx=0; $total=$rows.Count
foreach ($row in $rows) {
    $idx++
    $upn = ($row.UPN).Trim()
    $imgFile1 = ($row.ImageFile).Trim()

    if (-not $upn -or -not $imgFile1) { Write-Warn2 "[$idx/$total] Ignoré (UPN/ImageFile vide)"; continue }

    Write-Progress -Activity "Déploiement signatures OWA (GitHub - 3 images)" -Status "[$idx/$total] $upn" -PercentComplete ([double]$idx/[double]$total*100)

    # Boîte aux lettres ?
    try { $null = Get-EXOMailbox -Identity $upn -ErrorAction Stop }
    catch {
        Write-Warn2 "[$idx/$total] Boîte introuvable: $upn"
        $results.Add([pscustomobject]@{UPN=$upn;Image1=$imgFile1;Image2=$CommonImage2File;Image3=$CommonImage3File;Status='MailboxNotFound';Detail=''})
        continue
    }

    # -------- URL Image 1 (utilisateur) --------
    if ($imgFile1 -match '^https?://') {
        $imgUrl1 = $imgFile1
        if ($CacheBustingTag) {
            if ($imgUrl1 -match '\?') { $imgUrl1 += "&$CacheBustingTag" } else { $imgUrl1 += "?$CacheBustingTag" }
        }
    } else {
        $imgUrl1 = Get-GitHubRawUrl -Owner $GitHubOwner -Repo $GitHubRepo -Branch $GitBranch -RepoSubPath $RepoSubPath -ImageFile $imgFile1 -CacheTag $CacheBustingTag
    }

    $ok1 = Test-UrlOk -Url $imgUrl1
    if (-not $ok1) {
        $msg = "Image 1 non accessible: $imgUrl1"
        if ($FailIfUrlNotFound) {
            Write-Warn2 "[$idx/$total] $msg -> Ignoré."
            $results.Add([pscustomobject]@{UPN=$upn;Image1=$imgFile1;Image2=$CommonImage2File;Image3=$CommonImage3File;Status='Image1NotReachable';Detail=$msg})
            continue
        } else {
            Write-Warn2 "[$idx/$total] $msg -> On continue malgré tout."
        }
    }

    # -------- URL Images 2 et 3 (communes) --------
    $imgUrl2 = $null; $imgUrl3 = $null

    if ($CommonImage2File) {
        if ($CommonImage2File -match '^https?://') {
            $imgUrl2 = $CommonImage2File
            if ($CacheBustingTag) {
                if ($imgUrl2 -match '\?') { $imgUrl2 += "&$CacheBustingTag" } else { $imgUrl2 += "?$CacheBustingTag" }
            }
        } else {
            $imgUrl2 = Get-GitHubRawUrl -Owner $GitHubOwner -Repo $GitHubRepo -Branch $GitBranch -RepoSubPath $CommonImagesPath -ImageFile $CommonImage2File -CacheTag $CacheBustingTag
        }
    }

    if ($CommonImage3File) {
        if ($CommonImage3File -match '^https?://') {
            $imgUrl3 = $CommonImage3File
            if ($CacheBustingTag) {
                if ($imgUrl3 -match '\?') { $imgUrl3 += "&$CacheBustingTag" } else { $imgUrl3 += "?$CacheBustingTag" }
            }
        } else {
            $imgUrl3 = Get-GitHubRawUrl -Owner $GitHubOwner -Repo $GitHubRepo -Branch $GitBranch -RepoSubPath $CommonImagesPath -ImageFile $CommonImage3File -CacheTag $CacheBustingTag
        }
    }

    if ($FailIfUrlNotFound) {
        if ($imgUrl2 -and -not (Test-UrlOk -Url $imgUrl2)) { Write-Warning "Image 2 non accessible: $imgUrl2"; $imgUrl2 = $null }
        if ($imgUrl3 -and -not (Test-UrlOk -Url $imgUrl3)) { Write-Warning "Image 3 non accessible: $imgUrl3"; $imgUrl3 = $null }
    }

    # -------- HTML final : 3 images empilées --------
    $style1 = ($MaxWidthPx -gt 0) ? "max-width:${MaxWidthPx}px;height:auto;border:0;display:block;" : "height:auto;border:0;display:block;"
    $style2 = $style1
    $style3 = $style1

    $alt1 = $AltTextImage1 -replace '"','&quot;'
    $alt2 = $AltTextImage2 -replace '"','&quot;'
    $alt3 = $AltTextImage3 -replace '"','&quot;'

    $signatureHtml = "<div>"
    $signatureHtml += "<img src=""$imgUrl1"" alt=""$alt1"" style=""$style1"" />"
    if ($imgUrl2) { $signatureHtml += "<br/><img src=""$imgUrl2"" alt=""$alt2"" style=""$style2"" />" }
    if ($imgUrl3) { $signatureHtml += "<br/><img src=""$imgUrl3"" alt=""$alt3"" style=""$style3"" />" }
    $signatureHtml += "</div>"

    # (Optionnel) Debug HTML local
   # $debugHtml = Join-Path $env:TEMP "debug-signature-$($upn -replace '[^\w\.-]','_').html"
   # $signatureHtml | Out-File -FilePath $debugHtml -Encoding UTF8
   # Write-Host "HTML généré: $debugHtml"

    if ($WhatIf) {
        Write-Host "[$idx/$total] (WhatIf) $upn -> 3 images (1 spécifique + 2 communes)" -ForegroundColor Yellow
        $results.Add([pscustomobject]@{UPN=$upn;Image1=$imgFile1;Image2=$CommonImage2File;Image3=$CommonImage3File;Status='WhatIf';Detail='Aucune modification'})
        continue
    }

    # -------- Application OWA --------
    try {
        Set-MailboxMessageConfiguration -Identity $upn `
            -SignatureHtml $signatureHtml `
            -AutoAddSignature $AutoAddSignature `
            -AutoAddSignatureOnReply $AutoAddSignatureOnReply `
            -ErrorAction Stop

        Write-Ok "[$idx/$total] OK: Signature OWA appliquée à $upn (3 images)"
        $results.Add([pscustomobject]@{UPN=$upn;Image1=$imgFile1;Image2=$CommonImage2File;Image3=$CommonImage3File;Status='Success';Detail=''})
    } catch {
        Write-Err2 "[$idx/$total] Erreur $upn : $($_.Exception.Message)"
        $results.Add([pscustomobject]@{UPN=$upn;Image1=$imgFile1;Image2=$CommonImage2File;Image3=$CommonImage3File;Status='Error';Detail=$_.Exception.Message})
    }
}

# ===================== Export journal & fin =====================
try {
    $results | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $OutputLogCsv
    Write-Info "Journal exporté: $OutputLogCsv"
} catch {
    Write-Warn2 "Impossible d'écrire le journal: $($_.Exception.Message)"
}

Disconnect-ExchangeOnline -Confirm:$false
Write-Info "Terminé."