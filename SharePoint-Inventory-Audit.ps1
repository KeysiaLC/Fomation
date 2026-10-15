<#
.SYNOPSIS
    Inventaire et audit SharePoint Online : sites, niveau de partage, stockage, sites supprimés.

.DESCRIPTION
    Affiche :
      - tous les sites avec leur URL, leur titre, le stockage utilisé (Mo), le niveau de partage
        et la date de dernière modification du contenu ;
      - les sites encore ouverts au partage externe (SharingCapability différent de Disabled) ;
      - le détail d'un site (propriétaire, quota, partage) avec -SiteUrl ;
      - les sites supprimés encore dans la corbeille, avec le nombre de jours restants.
    -ExportCsv enregistre la liste des sites dans un fichier CSV.

.PARAMETER TenantName
    Nom du locataire, tel qu'il apparaît dans https://<TenantName>-admin.sharepoint.com.
.PARAMETER SiteUrl
    Facultatif : URL d'un site à afficher en détail.
.PARAMETER ExportCsv
    Facultatif : chemin d'un fichier CSV pour la liste des sites.

.EXAMPLE
    .\SharePoint-Inventory-Audit.ps1 -TenantName contoso
.EXAMPLE
    .\SharePoint-Inventory-Audit.ps1 -TenantName contoso -ExportCsv .\sites.csv
    Inventaire, avec export de la liste dans sites.csv.

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
    [string]$ExportCsv
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

Write-Host "`n=== Sites ===" -ForegroundColor Cyan
$sites = Get-SPOSite -Limit All |
    Select-Object Url, Title, @{ n = 'StorageMB'; e = { $_.StorageUsageCurrent } }, SharingCapability, LastContentModifiedDate
$sites | Sort-Object StorageMB -Descending | Format-Table -AutoSize
if ($ExportCsv) {
    $sites | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "Exported to $ExportCsv"
}

Write-Host "=== Sites open to external sharing / Sites ouverts au partage externe ===" -ForegroundColor Cyan
$sites | Where-Object SharingCapability -ne 'Disabled' | Format-Table Url, SharingCapability -AutoSize

if ($SiteUrl) {
    Write-Host "=== $SiteUrl ===" -ForegroundColor Cyan
    Get-SPOSite -Identity $SiteUrl | Format-List Url, Owner, StorageQuota, StorageUsageCurrent, SharingCapability
}

Write-Host "=== Deleted sites / Sites supprimés ===" -ForegroundColor Cyan
Get-SPODeletedSite | Format-Table Url, DeletionTime, DaysRemaining -AutoSize
