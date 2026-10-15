<#
.SYNOPSIS
    Permet aux utilisateurs de voir le calendrier d'une salle, d'un équipement ou d'une boîte partagée.

.DESCRIPTION
    Accorde, modifie ou retire un droit sur le dossier Calendrier d'une boîte aux lettres.
    Le nom réel du dossier est détecté automatiquement : « Calendrier » quand la langue
    de la boîte est le français, « Calendar » en anglais.
    Pour une salle, -KeepSubject conserve l'objet réel des réunions au lieu du nom de
    l'organisateur (s'applique aux nouvelles réservations seulement).

.PARAMETER Mailbox
    Salle, équipement ou boîte partagée.
.PARAMETER User
    Utilisateur ou groupe de sécurité à extension messagerie. « Default » = toute l'organisation. Par défaut : « Default ».
.PARAMETER AccessRights
    AvailabilityOnly (libre/occupé), LimitedDetails (libre/occupé, objet et lieu),
    Reviewer (lecture de tout le contenu), Editor (lecture et modification). Par défaut : LimitedDetails.
.PARAMETER KeepSubject
    Salles uniquement : conserve l'objet réel des réunions.
.PARAMETER Remove
    Retire le droit de -User.
.PARAMETER UserPrincipalName
    Compte administrateur utilisé pour la connexion ; il détermine le locataire.
.PARAMETER Device
    Connexion par code d'appareil (automatique dans PowerShell ISE).

.EXAMPLE
    .\Exchange-CalendarPermissions.ps1 -Mailbox salle-reunion@contoso.com -KeepSubject
    Toute l'organisation voit les créneaux, l'objet et le lieu des réservations de la salle.
.EXAMPLE
    .\Exchange-CalendarPermissions.ps1 -Mailbox compta@contoso.com -User "GRP-Ventes" -AccessRights Reviewer
    Le groupe GRP-Ventes lit tout le calendrier de la boîte partagée.
.EXAMPLE
    .\Exchange-CalendarPermissions.ps1 -Mailbox compta@contoso.com -User "GRP-Ventes" -Remove
    Retire ce droit.

.NOTES
    Auteur   : Laurent Cloux - KEYSIA
    Version  : 1.0 (2026-10)
    Licence  : MIT
    Modules  : installés automatiquement pour l'utilisateur courant s'ils manquent
    Prérequis: module ExchangeOnlineManagement, rôle Administrateur Exchange
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Mailbox,
    [string]$User = 'Default',
    [ValidateSet('AvailabilityOnly', 'LimitedDetails', 'Reviewer', 'Editor')]
    [string]$AccessRights = 'LimitedDetails',
    [switch]$KeepSubject,
    [switch]$Remove,
    [string]$UserPrincipalName,
    [switch]$Device
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
Install-RequiredModule -Name 'ExchangeOnlineManagement'

# --- Sign in to the tenant (skipped if already connected) / Connexion au locataire ---
# The tenant is the one of the -UserPrincipalName account. Le locataire est celui du compte -UserPrincipalName.
$session = Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1
if ($session -and $UserPrincipalName -and $session.UserPrincipalName -ne $UserPrincipalName) {
    Disconnect-ExchangeOnline -Confirm:$false; $session = $null
}
if (-not $session) {
    $c = @{ ShowBanner = $false }
    if ($UserPrincipalName) { $c.UserPrincipalName = $UserPrincipalName }
    if ($Device -or $psISE) { $c.Device = $true }
    Connect-ExchangeOnline @c
    $session = Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1
}
Write-Host "Connected to Exchange Online as $($session.UserPrincipalName) (tenant $($session.TenantID))" -ForegroundColor Green

# --- Real name of the Calendar folder (Calendar / Calendrier...) ---
$folderName = (Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope Calendar |
               Where-Object FolderType -eq 'Calendar' | Select-Object -First 1).Name
$folder = "$($Mailbox):\$folderName"
Write-Host "Calendar folder: $folder"

# --- Add, update or remove the permission ---
$existing = try { Get-MailboxFolderPermission -Identity $folder -User $User -ErrorAction Stop } catch { $null }
if ($Remove) {
    if ($existing -and $User -ne 'Default') {
        Remove-MailboxFolderPermission -Identity $folder -User $User -Confirm:$false
        Write-Host "Permission removed for $User"
    } else {
        Write-Warning 'Nothing to remove ("Default" cannot be removed: set it to AvailabilityOnly instead).'
    }
} elseif ($existing) {
    Set-MailboxFolderPermission -Identity $folder -User $User -AccessRights $AccessRights | Out-Null
} else {
    Add-MailboxFolderPermission -Identity $folder -User $User -AccessRights $AccessRights | Out-Null
}

# --- Rooms: keep the real subject instead of the organizer's name ---
if ($KeepSubject) {
    Set-CalendarProcessing -Identity $Mailbox -DeleteSubject $false -AddOrganizerToSubject $false
    Write-Host 'Room: the real meeting subject is now kept (new bookings only).'
}

Get-MailboxFolderPermission -Identity $folder | Format-Table User, AccessRights -AutoSize
