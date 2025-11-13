<#
.SYNOPSIS
Déploie une signature OWA contenant UNE image par utilisateur (hébergée sur GitHub - raw),
à partir d'un CSV (UPN,ImageFile).

CSV minimal (UTF-8) :
UPN,ImageFile
prenom.nom@mlpv.org,PrenomNOM.jpg

Notes:
- Compatible PowerShell 5.1 et 7+.

Auteur: Yaël
#>

param(
    # --- Chemin CSV ---
    [string]$CsvPath = "C:\Users\YaëlGALESNE\OneDrive - ARMLB\Documents\Signatures\56VAN\Signature_Vannes_Occurences.csv",

    # --- Cible GitHub ---
    [string]$GitHubOwner  = "Mission-Locale-de-Bretagne",
    [string]$GitHubRepo   = "Signatures",
    [string]$GitBranch    = "main",
    [string]$RepoSubPath  = "56VAN/SIGNATURE",

    # --- OWA ---
    [bool]$AutoAddSignature        = $true,
    [bool]$AutoAddSignatureOnReply = $true,

    # --- Contrôles & journal ---
    [bool]$FailIfUrlNotFound = $true,
    [string]$CacheBustingTag = "v=2025-11-01",
    [string]$OutputLogCsv    = ".\SignatureDeployment-OWA-GitHub-$((Get-Date).ToString('yyyyMMdd-HHmmss')).csv",
    [switch]$WhatIf,
    [string]$TestUPN = "",

    # --- Rendu visuel ---
    [string]$AltText = "Signature",
    [int]$MaxWidthPx = 720
)

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
    param([string]$Owner,[string]$Repo,[string]$Branch,[string]$RepoSubPath,[string]$ImageFile,[string]$CacheTag)
    $base = "https://raw.githubusercontent.com"
    $parts = @($Owner,$Repo,$Branch)
    if ($RepoSubPath) { $parts += $RepoSubPath.Split('/').Where({$_}) }
    if ($ImageFile)   { $parts += $ImageFile.Split('/').Where({$_}) }
    $encoded = $parts | ForEach-Object { Encode-UrlSegment $_ }
    $url = $base + "/" + ($encoded -join '/')
    if ($CacheTag) { $url += ( ($url -match '\?') ? "&$CacheTag" : "?$CacheTag" ) }
    $url
}

function Test-UrlOk([string]$Url){
    try {
        $r = Invoke-WebRequest -Uri $Url -Method Head -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
        return ($r.StatusCode -ge 200 -and $r.StatusCode -lt 300)
    } catch { return $false }
}

# --- Vérifs préalables ---
if (-not (Test-Path -LiteralPath $CsvPath)) { Write-Err2 "CSV introuvable : $CsvPath"; return }

Write-Info "Connexion à Exchange Online..."
Connect-ExchangeOnline -ShowBanner:$false

# Optionnel : si tu veux empêcher la synchro des signatures locales (roaming)
# Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater:$true

# Charge CSV
$rows = Import-CsvAuto -Path $CsvPath
if ($rows.Count -eq 0) { Write-Err2 "CSV vide."; Disconnect-ExchangeOnline -Confirm:$false; return }

# Colonnes minimales
$cols = $rows[0].PSObject.Properties.Name
foreach ($c in @('UPN','ImageFile')) {
    if ($c -notin $cols) { Write-Err2 "Colonnes requises manquantes (UPN, ImageFile). Trouvées: $($cols -join ', ')"; Disconnect-ExchangeOnline -Confirm:$false; return }
}

# Filtre test unitaire éventuel
if ($TestUPN) {
    $rows = $rows | Where-Object { $_.UPN -eq $TestUPN }
    if ($rows.Count -eq 0) { Write-Warn2 "UPN '$TestUPN' absent du CSV."; Disconnect-ExchangeOnline -Confirm:$false; return }
    Write-Info "Mode test: traitement limité à '$TestUPN' (lignes: $($rows.Count))."
}

# Journal mémoire
$results = New-Object System.Collections.Generic.List[object]

# Traitement
$idx=0; $total=$rows.Count
foreach ($row in $rows) {
    $idx++
    $upn = ($row.UPN).Trim()
    $imgFile = ($row.ImageFile).Trim()

    if (-not $upn -or -not $imgFile) { Write-Warn2 "[$idx/$total] Ignoré (UPN/ImageFile vide)"; continue }

    Write-Progress -Activity "Déploiement signatures OWA (GitHub)" -Status "[$idx/$total] $upn" -PercentComplete (($idx/$total)*100)

    # Boîte aux lettres existante ?
    try { $null = Get-EXOMailbox -Identity $upn -ErrorAction Stop }
    catch {
        Write-Warn2 "[$idx/$total] Boîte introuvable: $upn"
        $results.Add([pscustomobject]@{UPN=$upn;ImageFile=$imgFile;Url='';Status='MailboxNotFound';Detail=''})
        continue
    }

    # Si ImageFile est déjà une URL http(s), on l'utilise telle quelle. Sinon, on construit l'URL raw GitHub.
    if ($imgFile -match '^https?://') {
        $imgUrl = $imgFile
        if ($CacheBustingTag) { $imgUrl += ( ($imgUrl -match '\?') ? "&$CacheBustingTag" : "?$CacheBustingTag" ) }
    } else {
        $imgUrl = Get-GitHubRawUrl -Owner $GitHubOwner -Repo $GitHubRepo -Branch $GitBranch -RepoSubPath $RepoSubPath -ImageFile $imgFile -CacheTag $CacheBustingTag
    }

    # Vérifie accessibilité
    $ok = Test-UrlOk -Url $imgUrl
    if (-not $ok) {
        $msg = "Image non accessible: $imgUrl"
        if ($FailIfUrlNotFound) {
            Write-Warn2 "[$idx/$total] $msg -> Ignoré."
            $results.Add([pscustomobject]@{UPN=$upn;ImageFile=$imgFile;Url=$imgUrl;Status='ImageNotReachable';Detail=$msg})
            continue
        } else {
            Write-Warn2 "[$idx/$total] $msg -> On continue malgré tout."
        }
    }

 # HTML minimal : une image distante (balise <img> correcte)
$style = if ($MaxWidthPx -gt 0) { "max-width:${MaxWidthPx}px;height:auto;border:0;display:block;" } else { "height:auto;border:0;display:block;" }
$alt   = $AltText -replace '"','&quot;'   # protège l'attribut alt

# <<< LIGNE IMPORTANTE >>>
$signatureHtml = "<div><img src=""$imgUrl"" alt=""$alt"" style=""$style"" /></div>"

    if ($WhatIf) {
        Write-Host "[$idx/$total] (WhatIf) $upn -> $imgUrl" -ForegroundColor Yellow
        $results.Add([pscustomobject]@{UPN=$upn;ImageFile=$imgFile;Url=$imgUrl;Status='WhatIf';Detail='Aucune modification'})
        continue
    }

    try {
        Set-MailboxMessageConfiguration -Identity $upn `
            -SignatureHtml $signatureHtml `
            -AutoAddSignature $AutoAddSignature `
            -AutoAddSignatureOnReply $AutoAddSignatureOnReply `
            -ErrorAction Stop

        Write-Ok "[$idx/$total] OK: Signature OWA appliquée à $upn"
        $results.Add([pscustomobject]@{UPN=$upn;ImageFile=$imgFile;Url=$imgUrl;Status='Success';Detail=''})
    } catch {
        Write-Err2 "[$idx/$total] Erreur $upn : $($_.Exception.Message)"
        $results.Add([pscustomobject]@{UPN=$upn;ImageFile=$imgFile;Url=$imgUrl;Status='Error';Detail=$_.Exception.Message})
    }
}

try {
    $results | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $OutputLogCsv
    Write-Info "Journal exporté: $OutputLogCsv"
} catch {
    Write-Warn2 "Impossible d'écrire le journal: $($_.Exception.Message)"
}

Disconnect-ExchangeOnline -Confirm:$false
Write-Info "Terminé."