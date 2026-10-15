# Scripts PowerShell – Formation Microsoft 365

🇫🇷 Français · [🇬🇧 English](#english)

Scripts PowerShell présentés dans les exercices de la formation **Microsoft 365 – Administrer la suite complète**. Ils reprennent les manipulations des labs, avec des paramètres à la place des valeurs propres au locataire de lab.

## Les scripts

| Script | Service | Ce qu'il fait |
|---|---|---|
| [Exchange-SharedMailbox-Rooms.ps1](Exchange-SharedMailbox-Rooms.ps1) | Exchange Online | Boîte partagée avec Accès total et Envoyer en tant que, conversion d'une boîte utilisateur, salle avec règles de réservation, liste de salles |
| [Exchange-Inventory-MessageTrace.ps1](Exchange-Inventory-MessageTrace.ps1) | Exchange Online | Inventaire (boîtes partagées, règles de transport, domaines acceptés) et suivi des messages avec `Get-MessageTraceV2` |
| [Exchange-CalendarPermissions.ps1](Exchange-CalendarPermissions.ps1) | Exchange Online | Droits sur le calendrier d'une salle ou d'une boîte partagée, détection du dossier Calendar / Calendrier, objet réel des réunions conservé |
| [Exchange-RoomAdditionalResponse.ps1](Exchange-RoomAdditionalResponse.ps1) | Exchange Online | Message HTML ajouté à la réponse automatique d'une salle : modèle intégré compatible Outlook, rempli avec le nom et la capacité de la salle, ou fichier HTML exportable et modifiable |
| [SharePoint-Inventory-Audit.ps1](SharePoint-Inventory-Audit.ps1) | SharePoint Online | Inventaire des sites, sites ouverts au partage externe, sites supprimés, export CSV |
| [SharePoint-Sharing-Versions.ps1](SharePoint-Sharing-Versions.ps1) | SharePoint Online | Partage et quota d'un site, limite automatique des versions, expiration des liens « Tout le monde », type de lien par défaut |
| [OneDrive-Settings.ps1](OneDrive-Settings.ps1) | OneDrive | Conservation des OneDrive supprimés, partage, stockage, accès de l'administrateur au OneDrive d'un utilisateur |
| [Teams-New-ProjectTeam.ps1](Teams-New-ProjectTeam.ps1) | Teams | Équipe projet normalisée : canaux standards, canal privé, membres |
| [Restrict-M365GroupCreation.ps1](Restrict-M365GroupCreation.ps1) | Entra ID / Teams | Réserve la création des groupes Microsoft 365 et des équipes Teams aux membres d'un groupe ; `-AllowEveryone` pour revenir en arrière |

Chaque script contient une aide complète en français :

```powershell
Get-Help .\Exchange-CalendarPermissions.ps1 -Full
```

## Télécharger un script

- **Un fichier seul** : cliquez sur le script, puis sur l'icône **Download raw file**, en haut à droite du code.
- **Tous les scripts** : bouton vert **Code**, puis **Download ZIP**.

> ⚠️ **Débloquer les fichiers après téléchargement.** Windows marque les fichiers téléchargés comme provenant d'Internet. Avec la stratégie d'exécution `RemoteSigned`, PowerShell refuse alors de lancer le script (« … is not digitally signed »). Pour lever ce marquage, dans le dossier des scripts :
>
> ```powershell
> Get-ChildItem *.ps1 | Unblock-File
> ```
>
> Pour un seul fichier : clic droit, **Propriétés**, cocher **Débloquer**, puis **OK**.

## Prérequis

Chaque script **installe lui-même les modules qui manquent**, pour l'utilisateur courant (sans droits d'administrateur sur le poste), puis **se connecte au locataire**. Le tableau indique ce qui sera installé et le rôle nécessaire.

| Scripts | Module installé automatiquement | Rôle minimal |
|---|---|---|
| `Exchange-*` | `Install-Module ExchangeOnlineManagement -Scope CurrentUser` | Administrateur Exchange |
| `SharePoint-*`, `OneDrive-*` | `Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser` | Administrateur SharePoint |
| `Teams-*` | `Install-Module MicrosoftTeams -Scope CurrentUser` | Administrateur Teams |
| `Restrict-M365GroupCreation` | `Install-Module Microsoft.Graph.Beta.Identity.DirectoryManagement, Microsoft.Graph.Beta.Groups -Scope CurrentUser` | Administrateur général, licence Entra ID P1 pour les membres du groupe autorisé |

## Bon à savoir

- **Connexion au locataire** : chaque script ne demande la connexion que si la session n'est pas déjà connectée, puis affiche le compte et le locataire utilisés. Pour viser un locataire précis : `-UserPrincipalName` (scripts Exchange), `-TenantName` (SharePoint, OneDrive), `-TenantId` (Teams, Restrict-M365GroupCreation).
- **PowerShell ISE** : ISE ne sait pas afficher la fenêtre de connexion Windows moderne. Les scripts Exchange, Teams et Graph passent alors automatiquement à la **connexion par code d'appareil** : ouvrez https://microsoft.com/devicelogin et saisissez le code affiché.
- **SharePoint et OneDrive** : utilisez de préférence **Windows PowerShell 5.1**. Sous PowerShell 7, le script charge le module SharePoint Online en mode de compatibilité Windows PowerShell.
- **Ne mélangez pas** Graph et Exchange Online dans la même fenêtre PowerShell : les deux modules embarquent des versions différentes de la bibliothèque d'authentification et peuvent entrer en conflit. Ouvrez une fenêtre par service.
- **Lancez les scripts en tant que fichiers** (`.\Nom-Du-Script.ps1`) : ne collez pas leur contenu dans la console, un texte de plusieurs lignes y serait coupé.
- **Testez** d'abord sur un locataire de test.

## Auteur

Laurent Cloux, KEYSIA, Microsoft Certified Trainer.

Licence [MIT](LICENSE) : réutilisation libre, sans garantie.

---

<a id="english"></a>

# PowerShell scripts – Microsoft 365 training

[🇫🇷 Français](#scripts-powershell--formation-microsoft-365) · 🇬🇧 English

PowerShell scripts shown in the exercises of the **Microsoft 365 administration** training course. They reproduce the lab steps, with parameters instead of values specific to the lab tenant.

## The scripts

| Script | Service | What it does |
|---|---|---|
| [Exchange-SharedMailbox-Rooms.ps1](Exchange-SharedMailbox-Rooms.ps1) | Exchange Online | Shared mailbox with Full Access and Send As, conversion of a user mailbox, room mailbox with booking rules, room list |
| [Exchange-Inventory-MessageTrace.ps1](Exchange-Inventory-MessageTrace.ps1) | Exchange Online | Inventory (shared mailboxes, transport rules, accepted domains) and message trace with `Get-MessageTraceV2` |
| [Exchange-CalendarPermissions.ps1](Exchange-CalendarPermissions.ps1) | Exchange Online | Calendar permissions on a room or shared mailbox, Calendar / Calendrier folder detection, real meeting subject kept |
| [Exchange-RoomAdditionalResponse.ps1](Exchange-RoomAdditionalResponse.ps1) | Exchange Online | HTML message added to the automatic reply of a room: built-in Outlook-friendly template filled with the room name and capacity, or an HTML file you can export and edit |
| [SharePoint-Inventory-Audit.ps1](SharePoint-Inventory-Audit.ps1) | SharePoint Online | Site inventory, sites open to external sharing, deleted sites, CSV export |
| [SharePoint-Sharing-Versions.ps1](SharePoint-Sharing-Versions.ps1) | SharePoint Online | Site sharing and quota, automatic version history limits, "Anyone" link expiration, default link type |
| [OneDrive-Settings.ps1](OneDrive-Settings.ps1) | OneDrive | Retention of deleted users' OneDrive, sharing, storage, admin access to a user's OneDrive |
| [Teams-New-ProjectTeam.ps1](Teams-New-ProjectTeam.ps1) | Teams | Standardized project team: standard channels, private channel, members |
| [Restrict-M365GroupCreation.ps1](Restrict-M365GroupCreation.ps1) | Entra ID / Teams | Restricts Microsoft 365 group and Teams team creation to the members of one group; `-AllowEveryone` to roll back |

Each script includes full help (in French):

```powershell
Get-Help .\Exchange-CalendarPermissions.ps1 -Full
```

## Download a script

- **One file**: click the script, then the **Download raw file** icon at the top right of the code.
- **All scripts**: green **Code** button, then **Download ZIP**.

> ⚠️ **Unblock the files after downloading.** Windows marks downloaded files as coming from the Internet. With the `RemoteSigned` execution policy, PowerShell then refuses to run the script ("… is not digitally signed"). To remove this mark, in the scripts folder:
>
> ```powershell
> Get-ChildItem *.ps1 | Unblock-File
> ```
>
> For a single file: right-click, **Properties**, tick **Unblock**, then **OK**.

## Requirements

Each script **installs the missing modules itself**, for the current user (no local admin rights needed), then **signs in to the tenant**. The table shows what gets installed and the role required.

| Scripts | Module installed automatically | Minimum role |
|---|---|---|
| `Exchange-*` | `Install-Module ExchangeOnlineManagement -Scope CurrentUser` | Exchange Administrator |
| `SharePoint-*`, `OneDrive-*` | `Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser` | SharePoint Administrator |
| `Teams-*` | `Install-Module MicrosoftTeams -Scope CurrentUser` | Teams Administrator |
| `Restrict-M365GroupCreation` | `Install-Module Microsoft.Graph.Beta.Identity.DirectoryManagement, Microsoft.Graph.Beta.Groups -Scope CurrentUser` | Global Administrator, Entra ID P1 license for the members of the allowed group |

## Good to know

- **Tenant sign-in**: each script only signs in if the session is not already connected, then shows the account and tenant used. To target a specific tenant: `-UserPrincipalName` (Exchange scripts), `-TenantName` (SharePoint, OneDrive), `-TenantId` (Teams, Restrict-M365GroupCreation).
- **PowerShell ISE**: ISE cannot display the modern Windows sign-in window. The Exchange, Teams and Graph scripts then automatically switch to **device code sign-in**: open https://microsoft.com/devicelogin and enter the code shown.
- **SharePoint and OneDrive**: preferably use **Windows PowerShell 5.1**. In PowerShell 7, the script loads the SharePoint Online module in Windows PowerShell compatibility mode.
- **Don't mix** Graph and Exchange Online in the same PowerShell window: both modules ship different versions of the authentication library and may conflict. Use one window per service.
- **Run the scripts as files** (`.\Script-Name.ps1`): do not paste their content into the console, where multi-line text would be cut.
- **Test** on a test tenant first.

## Author

Laurent Cloux, KEYSIA, Microsoft Certified Trainer.

[MIT](LICENSE) license: free to reuse, no warranty.
