<#
.SYNOPSIS
    Ajoute un message HTML personnalisé à la réponse automatique d'une salle.

.DESCRIPTION
    Règle -AddAdditionalResponse et -AdditionalResponse (Set-CalendarProcessing).
    Trois façons de fournir le message :
      1. Modèle intégré (par défaut) : une fiche HTML compatible Outlook, remplie avec le nom
         et la capacité de la salle lus dans Exchange, et avec -Location, -Equipment,
         -WifiNetwork, -ContactPhone et -ContactEmail.
      2. -MessageFile : un fichier HTML écrit par vous ou exporté avec -ExportTemplate.
      3. -Message : une chaîne HTML courte.
    -ExportTemplate enregistre le modèle rempli dans un fichier SANS l'appliquer : on le modifie
    dans le Bloc-notes, puis on l'applique avec -MessageFile.

    Le texte est inséré dans le corps HTML de la réponse : balises simples et styles écrits dans
    les balises acceptés, pas de script ni de bloc <style>, images uniquement par lien https public.
    Ce fonctionnement HTML est constaté mais non documenté par Microsoft : testez le rendu dans
    Outlook classique et dans Outlook sur le web. La salle doit être en AutoAccept, sinon aucune
    réponse n'est envoyée.

    IMPORTANT : lancez ce script en tant que FICHIER .ps1. Ne collez pas son contenu dans la
    console : un message HTML de plusieurs lignes y serait coupé après la première.

.PARAMETER Room
    Boîte de la salle.
.PARAMETER Location
    Modèle : accès à la salle.
.PARAMETER Equipment
    Modèle : équipement de la salle.
.PARAMETER WifiNetwork
    Modèle : réseau Wi-Fi invités (vide = ligne masquée).
.PARAMETER ContactPhone
    Modèle : téléphone de l'accueil.
.PARAMETER ContactEmail
    Modèle : adresse e-mail de l'accueil.
.PARAMETER ExportTemplate
    Enregistre le modèle rempli dans ce fichier, sans l'appliquer.
.PARAMETER MessageFile
    Fichier HTML à appliquer (encodé en UTF-8).
.PARAMETER Message
    Chaîne HTML courte à appliquer.
.PARAMETER Disable
    Désactive le message (le texte reste enregistré).
.PARAMETER UserPrincipalName
    Compte administrateur utilisé pour la connexion ; il détermine le locataire.
.PARAMETER Device
    Connexion par code d'appareil (automatique dans PowerShell ISE).

.EXAMPLE
    .\Exchange-RoomAdditionalResponse.ps1 -Room salle-lyon@contoso.com -Location "3e étage, aile B - badge requis" -WifiNetwork "Contoso-Guest" -ContactPhone "poste 1234" -ContactEmail accueil@contoso.com
    Applique le modèle intégré.
.EXAMPLE
    .\Exchange-RoomAdditionalResponse.ps1 -Room salle-lyon@contoso.com -ExportTemplate .\message-salle.html
    .\Exchange-RoomAdditionalResponse.ps1 -Room salle-lyon@contoso.com -MessageFile .\message-salle.html
    Exporte le modèle, le modifie dans le Bloc-notes, puis l'applique.
.EXAMPLE
    .\Exchange-RoomAdditionalResponse.ps1 -Room salle-lyon@contoso.com -Disable
    Désactive le message.

.NOTES
    Auteur   : Laurent Cloux - KEYSIA
    Version  : 1.1 (2026-10)
    Licence  : MIT
    Modules  : installés automatiquement pour l'utilisateur courant s'ils manquent
    Prérequis: module ExchangeOnlineManagement, rôle Administrateur Exchange
#>
[CmdletBinding(DefaultParameterSetName = 'Template')]
param(
    [Parameter(Mandatory = $true)][string]$Room,

    [Parameter(ParameterSetName = 'Template')]
    [Parameter(ParameterSetName = 'Export')][string]$Location = 'Voir le plan à l''accueil',
    [Parameter(ParameterSetName = 'Template')]
    [Parameter(ParameterSetName = 'Export')][string]$Equipment = 'Écran Teams Rooms, caméra, tableau blanc',
    [Parameter(ParameterSetName = 'Template')]
    [Parameter(ParameterSetName = 'Export')][string]$WifiNetwork,
    [Parameter(ParameterSetName = 'Template')]
    [Parameter(ParameterSetName = 'Export')][string]$ContactPhone,
    [Parameter(ParameterSetName = 'Template')]
    [Parameter(ParameterSetName = 'Export')][string]$ContactEmail,

    [Parameter(Mandatory = $true, ParameterSetName = 'Export')][string]$ExportTemplate,
    [Parameter(Mandatory = $true, ParameterSetName = 'File')][string]$MessageFile,
    [Parameter(Mandatory = $true, ParameterSetName = 'Text')][string]$Message,
    [Parameter(Mandatory = $true, ParameterSetName = 'Disable')][switch]$Disable,

    [string]$UserPrincipalName,
    [switch]$Device
)

# --- Guard: the script must run from a .ps1 file, not be pasted into the console ---
if (-not $PSCommandPath) {
    Write-Warning 'Save this script as a .ps1 file and run it as a file (.\Exchange-RoomAdditionalResponse.ps1). Pasted into the console, a multi-line message would be cut. / Enregistrez ce script en .ps1 et lancez-le comme un fichier.'
    return
}

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

# --- Built-in HTML template: Outlook-friendly (tables, inline styles, no image) ---
# --- Modèle HTML intégré : compatible Outlook (tableaux, styles en ligne, sans image) ---
$template = @'
<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#242424;">
  <tr>
    <td style="background-color:#1F4E79;color:#FFFFFF;padding:12px 16px;font-size:16px;font-weight:bold;">
      {{ROOM}} &ndash; informations pratiques
    </td>
  </tr>
  <tr>
    <td style="padding:12px 16px;border-left:1px solid #D0D7DE;border-right:1px solid #D0D7DE;">
      <p style="margin:0 0 10px 0;">Votre réservation est confirmée. Merci de prendre connaissance des points suivants&nbsp;:</p>
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;font-size:14px;">
        {{ROWS}}
      </table>
      <p style="margin:12px 0 4px 0;font-weight:bold;">Avant de partir</p>
      <ul style="margin:0 0 10px 20px;padding:0;">
        <li>Remettez la salle en ordre et effacez le tableau.</li>
        <li>Si la réunion est annulée, <b>supprimez la réservation</b> pour libérer la salle.</li>
      </ul>
    </td>
  </tr>
  {{FOOTER}}
</table>
'@

function Get-HtmlSafeText([string]$Text) { [System.Net.WebUtility]::HtmlEncode($Text) }

function New-RoomMessage {
    $mbx = Get-Mailbox -Identity $Room
    $rows = @()
    $info = [ordered]@{ 'Accès' = $Location; 'Capacité' = $(if ($mbx.ResourceCapacity) { "$($mbx.ResourceCapacity) personnes" }); 'Équipement' = $Equipment; 'Wi-Fi invités' = $WifiNetwork }
    foreach ($k in $info.Keys) {
        if ($info[$k]) {
            $rows += "<tr><td style=""padding:3px 10px 3px 0;font-weight:bold;"">$k</td><td style=""padding:3px 0;"">$(Get-HtmlSafeText $info[$k])</td></tr>"
        }
    }
    $contact = @()
    if ($ContactPhone) { $contact += "au $(Get-HtmlSafeText $ContactPhone)" }
    if ($ContactEmail) { $contact += "<a href=""mailto:$ContactEmail"" style=""color:#2E74B5;"">$(Get-HtmlSafeText $ContactEmail)</a>" }
    $footer = if ($contact) {
        "<tr><td style=""background-color:#F3F6F9;padding:10px 16px;font-size:12px;color:#595959;border:1px solid #D0D7DE;border-top:none;"">Un problème avec la salle&nbsp;? Contactez l'accueil $($contact -join ' ou ').</td></tr>"
    } else {
        "<tr><td style=""border-bottom:1px solid #D0D7DE;font-size:1px;line-height:1px;"">&nbsp;</td></tr>"
    }
    $template.Replace('{{ROOM}}', (Get-HtmlSafeText $mbx.DisplayName)).Replace('{{ROWS}}', ($rows -join "`r`n        ")).Replace('{{FOOTER}}', $footer)
}

# --- Disable ---
if ($Disable) {
    Set-CalendarProcessing -Identity $Room -AddAdditionalResponse $false
    Write-Host 'Additional response disabled (text kept). / Message désactivé (texte conservé).'
    Get-CalendarProcessing -Identity $Room | Format-List AutomateProcessing, AddAdditionalResponse
    return
}

# --- Build the message ---
switch ($PSCmdlet.ParameterSetName) {
    'File'     { $Message = Get-Content -Path $MessageFile -Raw -Encoding UTF8 }
    'Template' { $Message = New-RoomMessage }
    'Export'   {
        New-RoomMessage | Set-Content -Path $ExportTemplate -Encoding UTF8
        Write-Host "Template saved to $ExportTemplate. Edit it, then run the script with -MessageFile. / Modèle enregistré : modifiez-le puis relancez avec -MessageFile." -ForegroundColor Green
        return
    }
}

# --- Basic checks on the HTML ---
if ($Message -match '<table' -and $Message -notmatch '</table>\s*$') {
    Write-Warning 'The HTML seems cut: it contains <table> but does not end with </table>. Nothing was changed. / Le HTML semble tronqué : rien n''a été modifié.'
    return
}
if ($Message -match '<script|<style') {
    Write-Warning 'Script and <style> blocks are removed by mail clients: use inline styles only. / Les blocs script et style sont supprimés par les clients de messagerie.'
}

# --- Apply ---
$processing = Get-CalendarProcessing -Identity $Room
if ($processing.AutomateProcessing -ne 'AutoAccept') {
    Write-Warning "AutomateProcessing is '$($processing.AutomateProcessing)': no automatic reply will be sent until it is AutoAccept."
}
Set-CalendarProcessing -Identity $Room -AddAdditionalResponse $true -AdditionalResponse $Message
Write-Host 'Additional response set. Book the room with a test account to check the rendering. / Message appliqué : réservez la salle avec un compte de test pour vérifier le rendu.' -ForegroundColor Green

$result = Get-CalendarProcessing -Identity $Room
$result | Format-List AutomateProcessing, AddAdditionalResponse
Write-Host "AdditionalResponse ($($result.AdditionalResponse.Length) characters):"
$result.AdditionalResponse
