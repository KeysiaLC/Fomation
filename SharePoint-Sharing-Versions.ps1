<#
.SYNOPSIS
    Réglages SharePoint Online : partage, quota de stockage et historique des versions.

.DESCRIPTION
    Chaque réglage n'est appliqué que si son paramètre est fourni :
      Au niveau d'un site (-SiteUrl) :
        -SharingCapability   niveau de partage du site (jamais plus ouvert que celui de l'organisation)
        -StorageQuotaMB      quota de stockage (uniquement si les limites de stockage des sites sont en mode Manuel)
        -AutoVersionTrim     limite automatique de l'historique des versions, bibliothèques existantes comprises
      Au niveau de l'organisation :
        -AnonymousLinkExpirationDays   les liens « Tout le monde » expirent après N jours
        -DefaultSharingLinkType        type de lien proposé par défaut dans la fenêtre Partager
    Les valeurs actuelles sont affichées à la fin.

.PARAMETER TenantName
    Nom du locataire, tel qu'il apparaît dans https://<TenantName>-admin.sharepoint.com.
.PARAMETER SiteUrl
    Site à configurer.
.PARAMETER SharingCapability
    Disabled (interne uniquement), ExistingExternalUserSharingOnly (invités existants),
    ExternalUserSharingOnly (nouveaux invités et invités existants), ExternalUserAndGuestSharing (tout le monde).
.PARAMETER StorageQuotaMB
    Quota du site en Mo (par exemple 51200 = 50 Go).
.PARAMETER AutoVersionTrim
    Active la limite automatique de l'historique des versions pour le site.
.PARAMETER AnonymousLinkExpirationDays
    Expiration des liens « Tout le monde », en jours.
.PARAMETER DefaultSharingLinkType
    Direct (personnes spécifiques), Internal (personnes de l'organisation), AnonymousAccess (tout le monde).

.EXAMPLE
    .\SharePoint-Sharing-Versions.ps1 -TenantName contoso -SiteUrl https://contoso.sharepoint.com/sites/Projets -SharingCapability Disabled -AutoVersionTrim
    Ferme le site au partage externe et active la limite automatique des versions.
.EXAMPLE
    .\SharePoint-Sharing-Versions.ps1 -TenantName contoso -AnonymousLinkExpirationDays 30 -DefaultSharingLinkType Direct
    Liens « Tout le monde » valables 30 jours, lien « Personnes spécifiques » proposé par défaut.

.NOTES
    Auteur   : Laurent Cloux - KEYSIA
    Version  : 1.0 (2026-10)
    Licence  : MIT
    Modules  : installés automatiquement pour l'utilisateur courant s'ils manquent
    Prérequis: module Microsoft.Online.SharePoint.PowerShell (Windows PowerShell 5.1 conseillé),
               rôle Administrateur SharePoint
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$TenantName,
    [string]$SiteUrl,
    [ValidateSet('Disabled', 'ExistingExternalUserSharingOnly', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing')]
    [string]$SharingCapability,
    [int]$StorageQuotaMB,
    [switch]$AutoVersionTrim,
    [ValidateRange(1, 730)][int]$AnonymousLinkExpirationDays,
    [ValidateSet('Direct', 'Internal', 'AnonymousAccess')]
    [string]$DefaultSharingLinkType
)
$ErrorActionPreference = 'Stop'

# --- Required modules: installed for the current user if missing, then loaded ---
# --- Modules requis : installés pour l'utilisateur courant s'ils manquent, puis chargés ---
function Install-RequiredModule {
    param([string[]]$Name)
    foreach ($m in $Name) {
        if (-not (Get-Module -ListAvailable -Name $m)) {
            Write-Host "Installing module $m (current user)... / Installation du module $m..." -ForegroundColor Yellow
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
            }
            Install-Module -Name $m -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
        }
        if ($m -eq 'Microsoft.Online.SharePoint.PowerShell' -and $PSVersionTable.PSEdition -eq 'Core') {
            Import-Module $m -UseWindowsPowerShell -DisableNameChecking -WarningAction SilentlyContinue
        } else {
            Import-Module $m -DisableNameChecking
        }
    }
}
Install-RequiredModule -Name 'Microsoft.Online.SharePoint.PowerShell'
# --- Sign in to the tenant (skipped if already connected) / Connexion au locataire ---
$adminUrl = "https://$TenantName-admin.sharepoint.com"
try { Get-SPOTenant | Out-Null } catch { Connect-SPOService -Url $adminUrl }
Write-Host "Connected to $adminUrl" -ForegroundColor Green

# --- Site level ---
if ($SiteUrl) {
    if ($SharingCapability) {
        Set-SPOSite -Identity $SiteUrl -SharingCapability $SharingCapability
        Write-Host "Sharing set to $SharingCapability"
    }
    if ($StorageQuotaMB) {
        # Works only when Settings > Site storage limits is set to Manual in the SharePoint admin center
        Set-SPOSite -Identity $SiteUrl -StorageQuota $StorageQuotaMB
        Write-Host "Storage quota set to $StorageQuotaMB MB"
    }
    if ($AutoVersionTrim) {
        # Starts a background job for the existing document libraries
        Set-SPOSite -Identity $SiteUrl -EnableAutoExpirationVersionTrim $true -ApplyToExistingDocumentLibraries -Confirm:$false
        Write-Host 'Automatic version history limits turned on (existing libraries: background job).'
    }
    Get-SPOSite -Identity $SiteUrl | Format-List Url, SharingCapability, StorageQuota, StorageUsageCurrent, EnableAutoExpirationVersionTrim
} elseif ($SharingCapability -or $StorageQuotaMB -or $AutoVersionTrim) {
    Write-Warning '-SharingCapability, -StorageQuotaMB and -AutoVersionTrim need -SiteUrl.'
}

# --- Organization level ---
if ($AnonymousLinkExpirationDays) {
    Set-SPOTenant -RequireAnonymousLinksExpireInDays $AnonymousLinkExpirationDays
    Write-Host "'Anyone' links now expire after $AnonymousLinkExpirationDays days"
}
if ($DefaultSharingLinkType) {
    Set-SPOTenant -DefaultSharingLinkType $DefaultSharingLinkType
    Write-Host "Default sharing link: $DefaultSharingLinkType"
}
Get-SPOTenant | Format-List SharingCapability, RequireAnonymousLinksExpireInDays, DefaultSharingLinkType
