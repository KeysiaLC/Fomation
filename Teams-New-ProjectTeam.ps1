<#
.SYNOPSIS
    Crée une équipe projet normalisée dans Microsoft Teams : canaux standards, un canal privé et des membres.

.DESCRIPTION
    Reprend les étapes du lab « Teams » du cours d'administration Microsoft 365 :
      1. Équipe privée (New-Team)
      2. Canaux standards (par défaut : Pilotage, Livrables)
      3. Un canal privé avec son propriétaire (par défaut : Direction)
      4. Membres
    Remarque : si la création des groupes est restreinte (voir Restrict-M365GroupCreation.ps1),
    le compte utilisé doit être autorisé à créer des groupes ou avoir un rôle d'administrateur.

.PARAMETER DisplayName
    Nom de l'équipe. Par défaut : « Projet Hermès ».
.PARAMETER MailNickName
    Alias du groupe Microsoft 365 associé (sans espace ni accent). Par défaut : « ProjetHermes ».
.PARAMETER Description
    Description de l'équipe.
.PARAMETER Channels
    Canaux standards à créer. Par défaut : « Pilotage », « Livrables ».
.PARAMETER PrivateChannel
    Nom du canal privé (vide = aucun). Par défaut : « Direction ».
.PARAMETER PrivateChannelOwner
    UPN du propriétaire du canal privé (obligatoire avec -PrivateChannel).
.PARAMETER Members
    UPN des membres à ajouter.
.PARAMETER TenantId
    Locataire cible (GUID ou domaine). Force une nouvelle connexion.
.PARAMETER UseDeviceAuthentication
    Connexion par code d'appareil (automatique dans PowerShell ISE).

.EXAMPLE
    .\Teams-New-ProjectTeam.ps1 -PrivateChannelOwner holly@contoso.com -Members alexw@contoso.com, adelev@contoso.com

.NOTES
    Auteur   : Laurent Cloux - KEYSIA
    Version  : 1.0 (2026-10)
    Licence  : MIT
    Modules  : installés automatiquement pour l'utilisateur courant s'ils manquent
    Prérequis: module MicrosoftTeams, rôle Administrateur Teams (ou utilisateur autorisé à créer des équipes)
#>
[CmdletBinding()]
param(
    [string]  $DisplayName  = 'Projet Hermès',
    [string]  $MailNickName = 'ProjetHermes',
    [string]  $Description  = 'Projet Hermès - équipe normalisée',
    [string[]]$Channels     = @('Pilotage', 'Livrables'),
    [string]  $PrivateChannel = 'Direction',
    [string]  $PrivateChannelOwner,
    [string[]]$Members,
    [string]  $TenantId,
    [switch]  $UseDeviceAuthentication
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
Install-RequiredModule -Name 'MicrosoftTeams'

if ($PrivateChannel -and -not $PrivateChannelOwner) { throw '-PrivateChannelOwner is required to create the private channel (or use -PrivateChannel "").' }

# --- Sign in to the tenant (skipped if already connected) / Connexion au locataire ---
$connected = $false
try { Get-CsTenant | Out-Null; $connected = $true } catch { }
if ($TenantId -and $connected) { Disconnect-MicrosoftTeams | Out-Null; $connected = $false }
if (-not $connected) {
    $c = @{}
    if ($TenantId) { $c.TenantId = $TenantId }
    if ($UseDeviceAuthentication -or $psISE) { $c.UseDeviceAuthentication = $true }
    Connect-MicrosoftTeams @c | Out-Null
}
Write-Host "Connected to tenant: $((Get-CsTenant).DisplayName)" -ForegroundColor Green

# --- 1. Team ---
$team = New-Team -DisplayName $DisplayName -Description $Description -Visibility Private -MailNickName $MailNickName
Write-Host "Team created: $DisplayName ($($team.GroupId))"

# --- 2. Standard channels ---
foreach ($c in $Channels) { New-TeamChannel -GroupId $team.GroupId -DisplayName $c | Out-Null }

# --- 3. Private channel (its owner is also added to the team if needed) ---
if ($PrivateChannel) {
    Add-TeamUser -GroupId $team.GroupId -User $PrivateChannelOwner -Role Member
    New-TeamChannel -GroupId $team.GroupId -DisplayName $PrivateChannel -MembershipType Private -Owner $PrivateChannelOwner | Out-Null
}

# --- 4. Members ---
foreach ($m in $Members) { Add-TeamUser -GroupId $team.GroupId -User $m -Role Member }

Get-TeamChannel -GroupId $team.GroupId | Format-Table DisplayName, MembershipType -AutoSize
Get-TeamUser -GroupId $team.GroupId | Format-Table User, Role -AutoSize
