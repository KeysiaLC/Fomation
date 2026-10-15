<#
.SYNOPSIS
    Crée une boîte partagée avec ses délégations, une salle avec ses règles de réservation et une liste de salles (Exchange Online).

.DESCRIPTION
    Reprend les étapes du lab « Messagerie » du cours d'administration Microsoft 365 :
      1. Boîte partagée, avec Accès total (et mappage automatique dans Outlook) et Envoyer en tant que pour un délégué
      2. En option : conversion d'une boîte utilisateur existante en boîte partagée
      3. Boîte de salle et règles de réservation (pas de réunions périodiques, réservation à 90 jours au plus, 4 h maximum)
      4. Liste de salles contenant la salle
    Les objets qui existent déjà ne sont pas recréés.

.PARAMETER SharedMailboxName
    Nom de la boîte partagée. Par défaut : « Support Adatum ».
.PARAMETER SharedMailboxAlias
    Alias de la boîte partagée. Par défaut : « support ».
.PARAMETER Delegate
    Utilisateur (UPN ou alias) qui reçoit l'Accès total et le droit Envoyer en tant que sur la boîte partagée.
.PARAMETER ConvertMailbox
    Facultatif : boîte utilisateur à convertir en boîte partagée.
.PARAMETER RoomName
    Nom de la boîte de salle. Par défaut : « Salle Lyon Part-Dieu ».
.PARAMETER RoomCapacity
    Capacité de la salle. Par défaut : 12.
.PARAMETER RoomListName
    Nom de la liste de salles. Par défaut : « Salles Lyon ».
.PARAMETER UserPrincipalName
    Compte administrateur utilisé pour la connexion ; il détermine le locataire.
.PARAMETER Device
    Connexion par code d'appareil (automatique dans PowerShell ISE).

.EXAMPLE
    .\Exchange-SharedMailbox-Rooms.ps1 -Delegate AdeleV -UserPrincipalName holly@contoso.onmicrosoft.com

.NOTES
    Auteur   : Laurent Cloux - KEYSIA
    Version  : 1.0 (2026-10)
    Licence  : MIT
    Modules  : installés automatiquement pour l'utilisateur courant s'ils manquent
    Prérequis: module ExchangeOnlineManagement, rôle Administrateur Exchange
#>
[CmdletBinding()]
param(
    [string]$SharedMailboxName  = 'Support Adatum',
    [string]$SharedMailboxAlias = 'support',
    [Parameter(Mandatory = $true)][string]$Delegate,
    [string]$ConvertMailbox,
    [string]$RoomName     = 'Salle Lyon Part-Dieu',
    [int]   $RoomCapacity = 12,
    [string]$RoomListName = 'Salles Lyon',
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

# Returns $true if the command finds the object, $false otherwise (works whatever the error mode)
function Test-Exists([scriptblock]$Command) { try { [bool](& $Command) } catch { $false } }

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

# --- 1. Shared mailbox and delegations ---
if (-not (Test-Exists { Get-Mailbox -Identity $SharedMailboxAlias -ErrorAction Stop })) {
    New-Mailbox -Shared -Name $SharedMailboxName -DisplayName $SharedMailboxName -Alias $SharedMailboxAlias | Out-Null
    Write-Host "Shared mailbox created: $SharedMailboxName"
}
# Full Access: the mailbox appears automatically in the delegate's Outlook (automapping)
Add-MailboxPermission -Identity $SharedMailboxAlias -User $Delegate -AccessRights FullAccess -InheritanceType All -AutoMapping $true -WarningAction SilentlyContinue | Out-Null
# Send As: the delegate sends messages that appear to come from the shared mailbox
Add-RecipientPermission -Identity $SharedMailboxAlias -Trustee $Delegate -AccessRights SendAs -Confirm:$false -WarningAction SilentlyContinue | Out-Null
Write-Host "Full Access and Send As granted to $Delegate"

# --- 2. Optional: convert a user mailbox to a shared mailbox (its license can then be removed) ---
if ($ConvertMailbox) {
    Set-Mailbox -Identity $ConvertMailbox -Type Shared
    Write-Host "Mailbox converted to shared: $ConvertMailbox"
}

# --- 3. Room mailbox and booking rules ---
if (-not (Test-Exists { Get-Mailbox -Identity $RoomName -ErrorAction Stop })) {
    New-Mailbox -Room -Name $RoomName -ResourceCapacity $RoomCapacity | Out-Null
    Write-Host "Room created: $RoomName"
}
Set-CalendarProcessing -Identity $RoomName -AllowRecurringMeetings $false -BookingWindowInDays 90 -MaximumDurationInMinutes 240

# --- 4. Room list (shown in the Outlook room finder) ---
if (-not (Test-Exists { Get-DistributionGroup -Identity $RoomListName -ErrorAction Stop })) {
    New-DistributionGroup -Name $RoomListName -RoomList | Out-Null
    Write-Host "Room list created: $RoomListName"
}
try { Add-DistributionGroupMember -Identity $RoomListName -Member $RoomName -ErrorAction Stop } catch { } # already a member

# --- Summary ---
Get-MailboxPermission -Identity $SharedMailboxAlias | Where-Object { -not $_.IsInherited -and $_.User -notlike 'NT AUTHORITY*' } | Format-Table User, AccessRights -AutoSize
Get-CalendarProcessing -Identity $RoomName | Format-List AutomateProcessing, AllowRecurringMeetings, BookingWindowInDays, MaximumDurationInMinutes
Get-DistributionGroupMember -Identity $RoomListName | Format-Table DisplayName, RecipientTypeDetails -AutoSize
