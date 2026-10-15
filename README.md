# Restreindre la création des groupes Microsoft 365 et des équipes Teams

🇫🇷 Français · [🇬🇧 English](#english)

Script PowerShell (Microsoft Graph) qui réserve la création des **groupes Microsoft 365**, et donc des **équipes Teams**, des plans Planner ou des sites SharePoint de groupe, aux seuls membres d'un groupe de sécurité.

Il automatise la procédure décrite par Microsoft : [Gérer qui peut créer des groupes Microsoft 365](https://learn.microsoft.com/fr-fr/microsoft-365/solutions/manage-creation-of-groups).

## Prérequis

- Windows PowerShell 5.1 ou PowerShell 7.
- Les modules Microsoft Graph **Beta** suivants :

  ```powershell
  Install-Module Microsoft.Graph.Beta.Identity.DirectoryManagement -Scope CurrentUser
  Install-Module Microsoft.Graph.Beta.Groups -Scope CurrentUser
  ```

- Un compte **Administrateur général** (ou Administrateur des groupes) et le consentement aux permissions `Directory.ReadWrite.All` et `Group.Read.All`.
- Un **groupe de sécurité** existant qui contient les personnes autorisées à créer des groupes, par exemple `GRP-Createurs-Equipes`.
- Une licence **Microsoft Entra ID P1** pour les membres de ce groupe.

## Utilisation

Restreindre la création au groupe `GRP-Createurs-Equipes` :

```powershell
.\Restrict-M365GroupCreation.ps1 -GroupName "GRP-Createurs-Equipes"
```

Revenir en arrière (tout le monde peut de nouveau créer des groupes) :

```powershell
.\Restrict-M365GroupCreation.ps1 -AllowEveryone
```

Aide complète :

```powershell
Get-Help .\Restrict-M365GroupCreation.ps1 -Full
```

Résultat attendu : `EnableGroupCreation` vaut `False` et `GroupCreationAllowedGroupId` contient l'identifiant du groupe autorisé.

## Bon à savoir

- **Connexion** : dans une console PowerShell, la fenêtre de connexion Windows s'ouvre normalement. Dans **PowerShell ISE**, qui ne peut pas l'afficher, le script passe automatiquement à la **connexion par code d'appareil** : ouvrez https://microsoft.com/devicelogin et saisissez le code affiché. Le paramètre `-UseDeviceCode` force ce mode ailleurs.
- **Délai** : comptez jusqu'à 30 minutes ou plus avant que la restriction ne s'applique partout, notamment dans Teams.
- **Les administrateurs** (Administrateur général, Administrateur Teams, etc.) peuvent toujours créer des groupes, même s'ils ne sont pas membres du groupe autorisé.
- **Erreur « module not found »** : installez les deux modules Beta ci-dessus. Ils ne font pas partie du module `Microsoft.Graph` standard.
- **Ne mélangez pas** Graph et Exchange Online dans la même fenêtre PowerShell : les deux modules embarquent des versions différentes de la bibliothèque d'authentification et peuvent entrer en conflit.

## Auteur

Laurent Cloux, KEYSIA, Microsoft Certified Trainer. Script partagé dans le cadre de formations Microsoft 365.

Licence [MIT](LICENSE) : réutilisation libre, sans garantie. Testez d'abord sur un locataire de test.

---

<a id="english"></a>

# Restrict Microsoft 365 group and Teams team creation

[🇫🇷 Français](#restreindre-la-création-des-groupes-microsoft-365-et-des-équipes-teams) · 🇬🇧 English

A PowerShell script (Microsoft Graph) that restricts the creation of **Microsoft 365 groups**, and therefore **Teams teams**, Planner plans and group-connected SharePoint sites, to the members of a single security group.

It automates the procedure documented by Microsoft: [Manage who can create Microsoft 365 Groups](https://learn.microsoft.com/microsoft-365/solutions/manage-creation-of-groups).

## Requirements

- Windows PowerShell 5.1 or PowerShell 7.
- The following Microsoft Graph **Beta** modules:

  ```powershell
  Install-Module Microsoft.Graph.Beta.Identity.DirectoryManagement -Scope CurrentUser
  Install-Module Microsoft.Graph.Beta.Groups -Scope CurrentUser
  ```

- A **Global Administrator** (or Groups Administrator) account, with consent to `Directory.ReadWrite.All` and `Group.Read.All`.
- An existing **security group** containing the people allowed to create groups, for example `GRP-Createurs-Equipes`.
- A **Microsoft Entra ID P1** license for the members of that group.

## Usage

Restrict creation to the `GRP-Createurs-Equipes` group:

```powershell
.\Restrict-M365GroupCreation.ps1 -GroupName "GRP-Createurs-Equipes"
```

Roll back (everyone can create groups again):

```powershell
.\Restrict-M365GroupCreation.ps1 -AllowEveryone
```

Full help:

```powershell
Get-Help .\Restrict-M365GroupCreation.ps1 -Full
```

Expected result: `EnableGroupCreation` is `False` and `GroupCreationAllowedGroupId` contains the Id of the allowed group.

## Good to know

- **Sign-in**: in a PowerShell console, the Windows sign-in window opens as usual. In **PowerShell ISE**, which cannot display it, the script automatically switches to **device code sign-in**: open https://microsoft.com/devicelogin and enter the code shown. Use `-UseDeviceCode` to force this mode elsewhere.
- **Delay**: allow 30 minutes or more for the restriction to apply everywhere, especially in Teams.
- **Administrators** (Global Administrator, Teams Administrator, etc.) can still create groups even if they are not members of the allowed group.
- **"Module not found" error**: install the two Beta modules above. They are not part of the standard `Microsoft.Graph` module.
- **Don't mix** Graph and Exchange Online in the same PowerShell window: both modules ship different versions of the authentication library and may conflict.

## Author

Laurent Cloux, KEYSIA, Microsoft Certified Trainer. Script shared as part of Microsoft 365 training courses.

[MIT](LICENSE) license: free to reuse, no warranty. Test on a test tenant first.
