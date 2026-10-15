<#
.SYNOPSIS
    Réserve la création des groupes Microsoft 365 (et des équipes Teams) aux membres d'un seul groupe de sécurité.

.DESCRIPTION
    Le script règle le paramètre d'annuaire « Group.Unified » de Microsoft Entra ID avec
    Microsoft Graph PowerShell (Beta) :
      - EnableGroupCreation         = False
      - GroupCreationAllowedGroupId = identifiant du groupe autorisé
    Si le paramètre « Group.Unified » n'existe pas encore dans le locataire, il est d'abord créé.
    -AllowEveryone permet de revenir en arrière : tout le monde peut de nouveau créer des groupes.

.PARAMETER GroupName
    Nom du groupe de sécurité dont les membres pourront créer des groupes.
.PARAMETER AllowEveryone
    Annule la restriction : tout le monde peut de nouveau créer des groupes.
.PARAMETER TenantId
    Locataire cible (GUID ou domaine, par exemple contoso.onmicrosoft.com). Force une nouvelle connexion.
.PARAMETER UseDeviceCode
    Force la connexion par code d'appareil. Utilisé automatiquement dans PowerShell ISE,
    où la fenêtre de connexion Windows (WAM) ne peut pas s'ouvrir.

.EXAMPLE
    .\Restrict-M365GroupCreation.ps1 -GroupName "GRP-Createurs-Equipes"
    Seuls les membres de GRP-Createurs-Equipes peuvent créer des groupes Microsoft 365 et des équipes Teams.
.EXAMPLE
    .\Restrict-M365GroupCreation.ps1 -AllowEveryone
    Supprime la restriction.

.NOTES
    Auteur   : Laurent Cloux - KEYSIA
    Version  : 1.0 (2026-10)
    Licence  : MIT
    Modules  : installés automatiquement pour l'utilisateur courant s'ils manquent
    Prérequis: modules Microsoft.Graph.Beta.Identity.DirectoryManagement et Microsoft.Graph.Beta.Groups ;
               rôle Administrateur général (ou Administrateur des groupes) et consentement à Directory.ReadWrite.All ;
               licence Microsoft Entra ID P1 pour les membres du groupe autorisé
    Source   : https://learn.microsoft.com/fr-fr/microsoft-365/solutions/manage-creation-of-groups
#>

[CmdletBinding(DefaultParameterSetName = 'Restrict')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Restrict')]
    [string]$GroupName,

    [Parameter(Mandatory = $true, ParameterSetName = 'Rollback')]
    [switch]$AllowEveryone,

    [string]$TenantId,

    [switch]$UseDeviceCode
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
Install-RequiredModule -Name 'Microsoft.Graph.Beta.Identity.DirectoryManagement', 'Microsoft.Graph.Beta.Groups'
$templateId = '62375ab9-6b52-47ed-826b-58e47e0e304b'   # "Group.Unified" settings template
$scopes     = @('Directory.ReadWrite.All', 'Group.Read.All')


# --- Sign in to the tenant (skipped if already connected with the required scopes) ---
# --- Connexion au locataire (ignorée si déjà connecté avec les bonnes permissions) ---
$context = Get-MgContext
$missing = if ($context) { $scopes | Where-Object { $context.Scopes -notcontains $_ } } else { $scopes }
if ($TenantId -and $context) { Disconnect-MgGraph | Out-Null; $missing = $scopes }
if ($missing) {
    $c = @{ Scopes = $scopes; NoWelcome = $true }
    if ($TenantId) { $c.TenantId = $TenantId }
    # PowerShell ISE cannot display the Windows sign-in window (WAM): use a device code there.
    if ($psISE -or $UseDeviceCode) { $c.UseDeviceCode = $true }
    Connect-MgGraph @c
}
$context = Get-MgContext
Write-Host "Connected to tenant $($context.TenantId) as $($context.Account)" -ForegroundColor Green

# --- Get the "Group.Unified" setting, or create it ---
$settingId = (Get-MgBetaDirectorySetting | Where-Object TemplateId -eq $templateId).Id
if (-not $settingId) {
    Write-Host 'Group.Unified setting not found: creating it.'
    $params = @{
        templateId = $templateId
        values     = @(@{ name = 'EnableMSStandardBlockedWords'; value = 'true' })
    }
    $settingId = (New-MgBetaDirectorySetting -BodyParameter $params).Id
}

# --- Build the values to apply ---
if ($AllowEveryone) {
    $values = @(
        @{ name = 'EnableGroupCreation';         value = 'True' },
        @{ name = 'GroupCreationAllowedGroupId'; value = '' }
    )
} else {
    $safeName = $GroupName.Replace("'", "''")
    $group = Get-MgBetaGroup -Filter "displayName eq '$safeName'"
    if (-not $group) {
        Write-Warning "Group '$GroupName' not found / groupe introuvable. Nothing was changed."
        return
    }
    if (@($group).Count -gt 1) {
        Write-Warning "Several groups are named '$GroupName' / plusieurs groupes portent ce nom. Nothing was changed."
        return
    }
    $values = @(
        @{ name = 'EnableGroupCreation';         value = 'False' },
        @{ name = 'GroupCreationAllowedGroupId'; value = $group.Id }
    )
}

# --- Apply and display the result ---
Update-MgBetaDirectorySetting -DirectorySettingId $settingId -BodyParameter @{ templateId = $templateId; values = $values }
(Get-MgBetaDirectorySetting -DirectorySettingId $settingId).Values |
    Where-Object Name -in 'EnableGroupCreation', 'GroupCreationAllowedGroupId' |
    Format-Table Name, Value -AutoSize
