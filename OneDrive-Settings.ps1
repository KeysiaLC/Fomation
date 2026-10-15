<#
.SYNOPSIS
    Réglages OneDrive : conservation des OneDrive supprimés, partage, stockage, accès de l'administrateur au OneDrive d'un utilisateur.

.DESCRIPTION
    Chaque réglage n'est appliqué que si son paramètre est fourni :
      Organisation :
        -RetentionDays     durée de conservation du OneDrive d'un utilisateur supprimé (30 à 3650 jours)
        -OneDriveSharing   niveau de partage OneDrive (jamais plus ouvert que SharePoint)
        -StorageQuotaMB    stockage OneDrive par défaut (par exemple 1048576 = 1 To)
      Un utilisateur (-UserPrincipalName) :
        -DisableUserSharing           désactive le partage pour le OneDrive de cet utilisateur
        -GrantAdmin / -RevokeAdmin    donne ou retire à -AdminUpn les droits d'administrateur
                                      sur le OneDrive de l'utilisateur (par exemple pour récupérer
                                      des fichiers après un départ)

.PARAMETER TenantName
    Nom du locataire, tel qu'il apparaît dans https://<TenantName>-admin.sharepoint.com.
.PARAMETER RetentionDays
    Durée de conservation du OneDrive d'un utilisateur supprimé, en jours.
.PARAMETER OneDriveSharing
    Disabled, ExistingExternalUserSharingOnly, ExternalUserSharingOnly ou ExternalUserAndGuestSharing.
.PARAMETER StorageQuotaMB
    Stockage OneDrive par défaut, en Mo.
.PARAMETER UserPrincipalName
    Utilisateur dont le OneDrive est visé.
.PARAMETER DisableUserSharing
    Désactive le partage de ce OneDrive.
.PARAMETER AdminUpn
    Compte administrateur pour -GrantAdmin / -RevokeAdmin.
.PARAMETER GrantAdmin
    Donne à -AdminUpn l'accès au OneDrive de l'utilisateur.
.PARAMETER RevokeAdmin
    Retire cet accès.

.EXAMPLE
    .\OneDrive-Settings.ps1 -TenantName contoso -RetentionDays 90 -OneDriveSharing ExistingExternalUserSharingOnly
    OneDrive supprimés conservés 90 jours, partage limité aux invités existants.
.EXAMPLE
    .\OneDrive-Settings.ps1 -TenantName contoso -UserPrincipalName adelev@contoso.com -AdminUpn holly@contoso.com -GrantAdmin
    Donne à Holly l'accès au OneDrive d'Adele.

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
    [ValidateRange(30, 3650)][int]$RetentionDays,
    [ValidateSet('Disabled', 'ExistingExternalUserSharingOnly', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing')]
    [string]$OneDriveSharing,
    [int]$StorageQuotaMB,
    [string]$UserPrincipalName,
    [switch]$DisableUserSharing,
    [string]$AdminUpn,
    [switch]$GrantAdmin,
    [switch]$RevokeAdmin
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

# --- Organization ---
if ($RetentionDays)   { Set-SPOTenant -OrphanedPersonalSitesRetentionPeriod $RetentionDays }
if ($OneDriveSharing) { Set-SPOTenant -OneDriveSharingCapability $OneDriveSharing }
if ($StorageQuotaMB)  { Set-SPOTenant -OneDriveStorageQuota $StorageQuotaMB }

# --- One user ---
if ($UserPrincipalName) {
    # OneDrive URL: https://<tenant>-my.sharepoint.com/personal/<upn with . and @ replaced by _>
    $oneDriveUrl = "https://$TenantName-my.sharepoint.com/personal/" + ($UserPrincipalName.ToLower() -replace '[\.@]', '_')
    Write-Host "OneDrive: $oneDriveUrl"

    if ($DisableUserSharing) {
        Set-SPOSite -Identity $oneDriveUrl -SharingCapability Disabled
        Write-Host 'Sharing turned off for this OneDrive.'
    }
    if ($GrantAdmin -or $RevokeAdmin) {
        if (-not $AdminUpn) { throw '-GrantAdmin / -RevokeAdmin need -AdminUpn.' }
        Set-SPOUser -Site $oneDriveUrl -LoginName $AdminUpn -IsSiteCollectionAdmin ([bool]$GrantAdmin) | Out-Null
        Write-Host ("Admin access for $AdminUpn : " + $(if ($GrantAdmin) { 'granted' } else { 'removed' }))
    }
    Get-SPOSite -Identity $oneDriveUrl | Format-List Url, Owner, SharingCapability, StorageUsageCurrent
}

Get-SPOTenant | Format-List OrphanedPersonalSitesRetentionPeriod, OneDriveSharingCapability, OneDriveStorageQuota
