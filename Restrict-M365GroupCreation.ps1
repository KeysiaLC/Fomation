<#
.SYNOPSIS
    Restricts Microsoft 365 group (and Teams team) creation to the members of a single security group.
    Réserve la création des groupes Microsoft 365 (et des équipes Teams) aux membres d'un groupe de sécurité.

.DESCRIPTION
    The script configures the "Group.Unified" directory setting in Microsoft Entra ID through
    Microsoft Graph PowerShell (Beta):
      - EnableGroupCreation         = False
      - GroupCreationAllowedGroupId = Id of the allowed group
    If the "Group.Unified" setting does not exist yet in the tenant, it is created first.
    Use -AllowEveryone to roll back (everyone can create groups again).

    Le script règle le paramètre d'annuaire « Group.Unified » de Microsoft Entra ID via
    Microsoft Graph PowerShell (Beta). Le paramètre est créé s'il n'existe pas encore.
    -AllowEveryone permet de revenir en arrière (tout le monde peut de nouveau créer des groupes).

.PARAMETER GroupName
    Display name of the security group whose members are allowed to create groups.
    Nom du groupe de sécurité dont les membres pourront créer des groupes.

.PARAMETER AllowEveryone
    Rolls back the restriction: everyone can create groups again.
    Annule la restriction : tout le monde peut de nouveau créer des groupes.

.PARAMETER UseDeviceCode
    Forces sign-in with a device code. Used automatically in Windows PowerShell ISE,
    where the default Windows sign-in window (WAM) cannot open.
    Force la connexion par code d'appareil. Utilisé automatiquement dans PowerShell ISE.

.EXAMPLE
    .\Restrict-M365GroupCreation.ps1 -GroupName "GRP-Createurs-Equipes"
    Only members of GRP-Createurs-Equipes can create Microsoft 365 groups and Teams teams.

.EXAMPLE
    .\Restrict-M365GroupCreation.ps1 -AllowEveryone
    Removes the restriction.

.NOTES
    Author  : Laurent Cloux - KEYSIA
    Version : 1.0 (2026-10)
    License : MIT
    Requires: Microsoft.Graph.Beta.Identity.DirectoryManagement, Microsoft.Graph.Beta.Groups
              Global Administrator or Groups Administrator + Directory.ReadWrite.All consent
              Microsoft Entra ID P1 for the members of the allowed group
    Based on: https://learn.microsoft.com/microsoft-365/solutions/manage-creation-of-groups
#>

[CmdletBinding(DefaultParameterSetName = 'Restrict')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Restrict')]
    [string]$GroupName,

    [Parameter(Mandatory = $true, ParameterSetName = 'Rollback')]
    [switch]$AllowEveryone,

    [switch]$UseDeviceCode
)

$ErrorActionPreference = 'Stop'
$templateId = '62375ab9-6b52-47ed-826b-58e47e0e304b'   # "Group.Unified" settings template
$scopes     = @('Directory.ReadWrite.All', 'Group.Read.All')

Import-Module Microsoft.Graph.Beta.Identity.DirectoryManagement
Import-Module Microsoft.Graph.Beta.Groups

# --- Sign in (skipped if already connected with the required scopes) ---
$context = Get-MgContext
$missing = if ($context) { $scopes | Where-Object { $context.Scopes -notcontains $_ } } else { $scopes }
if ($missing) {
    # PowerShell ISE cannot display the Windows sign-in window (WAM): use a device code there.
    if ($psISE -or $UseDeviceCode) {
        Connect-MgGraph -Scopes $scopes -UseDeviceCode -NoWelcome
    } else {
        Connect-MgGraph -Scopes $scopes -NoWelcome
    }
}

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
