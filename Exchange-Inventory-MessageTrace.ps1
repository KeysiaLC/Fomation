<#
.SYNOPSIS
    Inventaire Exchange Online (boîtes partagées, règles de transport, domaines acceptés) et suivi des messages.

.DESCRIPTION
    Sans -SenderAddress ni -RecipientAddress : affiche l'inventaire.
    Avec l'un des deux : lance un suivi des messages (Get-MessageTraceV2) et, avec -Detail,
    affiche le détail des événements de chaque message (Get-MessageTraceDetailV2).
    Get-MessageTraceV2 remonte jusqu'à 90 jours en arrière, 10 jours au plus par requête.

.PARAMETER SenderAddress
    Expéditeur à suivre.
.PARAMETER RecipientAddress
    Destinataire à suivre.
.PARAMETER Subject
    Facultatif : seulement les messages dont l'objet contient ce texte.
.PARAMETER Days
    Nombre de jours en arrière (1 à 10). Par défaut : 1.
.PARAMETER Detail
    Affiche aussi le détail des événements (réception, remise, échec...).
.PARAMETER UserPrincipalName
    Compte administrateur utilisé pour la connexion ; il détermine le locataire.
.PARAMETER Device
    Connexion par code d'appareil (automatique dans PowerShell ISE).

.EXAMPLE
    .\Exchange-Inventory-MessageTrace.ps1
    Inventaire seul.
.EXAMPLE
    .\Exchange-Inventory-MessageTrace.ps1 -SenderAddress adelev@contoso.com -Subject "Test suivi" -Detail
    Suivi des messages d'Adele dont l'objet contient « Test suivi », avec le détail des événements.

.NOTES
    Auteur   : Laurent Cloux - KEYSIA
    Version  : 1.0 (2026-10)
    Licence  : MIT
    Modules  : installés automatiquement pour l'utilisateur courant s'ils manquent
    Prérequis: module ExchangeOnlineManagement 3.7 ou ultérieur (pour Get-MessageTraceV2)
#>
[CmdletBinding()]
param(
    [string]$SenderAddress,
    [string]$RecipientAddress,
    [string]$Subject,
    [ValidateRange(1, 10)][int]$Days = 1,
    [switch]$Detail,
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

if (-not $SenderAddress -and -not $RecipientAddress) {
    Write-Host "`n=== Shared mailboxes / Boîtes partagées ===" -ForegroundColor Cyan
    Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Format-Table DisplayName, PrimarySmtpAddress -AutoSize

    Write-Host "=== Transport rules / Règles de transport ===" -ForegroundColor Cyan
    Get-TransportRule | Format-Table Name, Priority, State, Mode -AutoSize

    Write-Host "=== Accepted domains / Domaines acceptés ===" -ForegroundColor Cyan
    Get-AcceptedDomain | Format-Table DomainName, DomainType, Default -AutoSize
    return
}

# --- Message trace ---
$t = @{ StartDate = (Get-Date).AddDays(-$Days); EndDate = Get-Date }
if ($SenderAddress)    { $t.SenderAddress    = $SenderAddress }
if ($RecipientAddress) { $t.RecipientAddress = $RecipientAddress }
if ($Subject)          { $t.Subject = $Subject; $t.SubjectFilterType = 'Contains' }

$messages = Get-MessageTraceV2 @t
if (-not $messages) { Write-Warning 'No message found / aucun message trouvé.'; return }
$messages | Format-Table Received, SenderAddress, RecipientAddress, Subject, Status -AutoSize

if ($Detail) {
    foreach ($m in $messages) {
        Write-Host "`n--- $($m.Subject) -> $($m.RecipientAddress) ---" -ForegroundColor Cyan
        Get-MessageTraceDetailV2 -MessageTraceId $m.MessageTraceId -RecipientAddress $m.RecipientAddress |
            Format-Table Date, Event, Detail -AutoSize -Wrap
    }
}
