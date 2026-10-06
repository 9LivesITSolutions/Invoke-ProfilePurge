# Invoke-ProfilePurge

> **Script PowerShell de purge automatisée des profils locaux sur un parc de serveurs Windows.**

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)](https://github.com/PowerShell/PowerShell)
[![PS7 Parallel](https://img.shields.io/badge/PS7-Mode%20parall%C3%A8le-blueviolet?logo=powershell)](https://github.com/PowerShell/PowerShell)
[![Licence: MIT](https://img.shields.io/badge/Licence-MIT-green.svg)](LICENSE)

[English version](README.md)

---

## Présentation

`Invoke-ProfilePurge` est un script PowerShell de production pour les administrateurs Windows qui doivent nettoyer régulièrement les profils utilisateurs obsolètes sur des serveurs (RDS, serveurs de fichiers, VDI, contrôleurs de domaine). Il combine trois phases de nettoyage, un système de journalisation structuré, des rapports HTML, une exécution parallèle et une intégration au journal Windows.

---

## Fonctionnalités

| Fonctionnalité | Détail |
|---|---|
| **Purge des profils** | Suppression via `Win32_UserProfile` CIM (dossier + clé registre) |
| **Critère de date** | `Max(ntuser.dat LastWriteTime, CIM LastUseTime)` — résistant à WSearch |
| **Nettoyage clés .bak** | Clés `ProfileList\*.bak` orphelines avec résolution SID→utilisateur |
| **Nettoyage dossiers BACKUP** | Dossiers `*BACKUP*` avec garde-fou session active |
| **Réparation doublons domaine** | Paires `username` / `username.DOMAINE` repontées vers le profil local |
| **Arrêt/démarrage WSearch** | Libère les verrous `ntuser.dat.LOG` ; dates calculées AVANT l'arrêt |
| **Exécution parallèle** | `ForEach-Object -Parallel` PS7 avec mutex pour logs thread-safe |
| **Journalisation structurée** | `Write-Log` 6 niveaux, console colorée + fichier `.log` UTF-8 BOM |
| **Rapport HTML** | Thème clair, topbar sticky, KPI strip, tables filtrées sur les actions |
| **Journal Windows** | EventID 4100 (synthèse) et 4199 (erreur critique) |
| **Mode WhatIf** | Simulation complète, aucune modification |
| **PS5.1 + PS7** | Compatibilité double ; parallèle PS7 uniquement |

---

## Fonctionnement du critère de date

Le critère d'inactivité est `Max(ntuser.dat LastWriteTime, CIM LastUseTime)` :

| Source | Mis à jour par | Fiable ? |
|---|---|---|
| `ntuser.dat` LastWriteTime | Windows au logoff (commit du hive registre) | ✅ Oui |
| `Win32_UserProfile.LastUseTime` | Windows au logoff, stocké dans `ProfileList\<SID>` | ✅ Oui (peut être null) |
| Dossier profil LastWriteTime | Tout process écrivant dans l'arborescence profil | ❌ Contaminé |

`ntuser.dat` est ouvert en lecture seule par Windows Search pour l'indexation — il n'est jamais modifié par WSearch, les antivirus, la copie shadow ou les agents de sauvegarde. Il reflète la vraie dernière session utilisateur.

L'architecture deux passes garantit que WSearch est arrêté **après** le calcul de toutes les dates, évitant toute contamination pendant le run.

---

## Prérequis

- PowerShell **5.1+** (mode parallèle : **7.0+**)
- **Droits admin élevés** sur chaque serveur cible (`SeRestorePrivilege` requis pour `Win32_UserProfile.Delete()`)
- **WinRM** activé sur les cibles distantes (`Enable-PSRemoting`)
- Création de la source Event Log : droits admin à la **première exécution uniquement**

---

## Installation

```powershell
git clone https://github.com/9LivesITSolutions/Invoke-ProfilePurge.git
cd Invoke-ProfilePurge
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Aucune dépendance externe. Aucun module à installer.

---

## Démarrage rapide

```powershell
# Simulation en premier — toujours
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -WhatIf

# Exécution réelle — serveur local
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -StopWSearch -LogPath C:\Logs

# Purge complète — toutes les phases
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 `
    -PurgeProfileListBak -PurgeBackupFolders `
    -RepairDomainDuplicates -StopWSearch `
    -WriteEventLog -LogPath C:\Logs
```

---

## Paramètres

### Cibles

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-ComputerName` | `string[]` | — | Serveurs cibles inline |
| `-ComputerList` | `string` | — | Fichier texte (1 serveur/ligne) |
| `-Credential` | `PSCredential` | — | Credential pour sessions WinRM |

### Purge des profils

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-DaysInactive` | `int` | `90` | Seuil d'inactivité en jours |
| `-ExcludeUsers` | `string[]` | — | Exclusions inline (wildcards : `svc_*`) |
| `-ExcludeFile` | `string` | — | Fichier whitelist (1/ligne) |
| `-DeleteUnknownDate` | `switch` | off | Supprimer les profils sans date (⚠ comptes service) |
| `-StopWSearch` | `switch` | off | Arrêter Windows Search avant la purge |
| `-RepairDomainDuplicates` | `switch` | off | Réparer les paires `user` / `user.DOMAINE` |

### Phases supplémentaires

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-PurgeProfileListBak` | `switch` | off | Supprimer les clés `.bak` orphelines |
| `-PurgeBackupFolders` | `switch` | off | Supprimer les dossiers `*BACKUP*` |
| `-UsersPath` | `string` | `C:\Users` | Chemin racine pour la recherche BACKUP |

### Exécution

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-Parallel` | `switch` | off | Traitement parallèle (PS7+) |
| `-ThrottleLimit` | `int` | `5` | Serveurs simultanés maximum |
| `-WhatIf` | `switch` | off | Simulation — aucune modification |
| `-PassThru` | `switch` | off | Émettre les objets résultat dans le pipeline |

### Sortie

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-LogPath` | `string` | Dossier script | Destination logs + rapport HTML |
| `-LogRetentionDays` | `int` | `30` | Rétention des logs en jours |
| `-ReportPath` | `string` | Auto | Chemin personnalisé du rapport HTML |
| `-WriteEventLog` | `switch` | off | Écrire dans le journal Windows |
| `-EventSource` | `string` | `ProfilePurge` | Nom de la source d'événement |
| `-EventLogName` | `string` | `Application` | Journal cible |

---

## Exemples

```powershell
# Simulation locale
.\Invoke-ProfilePurge.ps1 -DaysInactive 60 -WhatIf

# Purge multi-serveurs en production
.\Invoke-ProfilePurge.ps1 `
    -ComputerList .\servers.txt -DaysInactive 90 `
    -ExcludeFile .\whitelist.txt `
    -PurgeProfileListBak -PurgeBackupFolders `
    -StopWSearch -WriteEventLog -LogPath C:\Logs\ProfilePurge

# Mode parallèle PS7
.\Invoke-ProfilePurge.ps1 `
    -ComputerList .\servers.txt -DaysInactive 90 `
    -Parallel -ThrottleLimit 8 `
    -WriteEventLog -LogPath C:\Logs

# Réparation des doublons domaine
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -RepairDomainDuplicates -WhatIf

# Export CSV
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -PassThru |
    Export-Csv -Path C:\Logs\purge.csv -NoTypeInformation -Encoding UTF8
```

---

## Codes de sortie

| Code | Signification |
|---|---|
| `0` | Succès |
| `1` | Erreur critique non gérée |
| `2` | Partiel — erreurs WinRM ou par profil |

---

## Journal Windows (Event Log)

| EventID | Type | Déclencheur |
|---|---|---|
| `4100` | Information / Warning / Error | Synthèse de fin d'exécution |
| `4199` | Error | Exception critique + stack trace |

```powershell
# Créer la source une fois (admin requis, première exécution uniquement)
New-EventLog -LogName Application -Source ProfilePurge
```

---

## Tâche planifiée

```
Programme  : powershell.exe
Arguments  : -NonInteractive -NoProfile -ExecutionPolicy Bypass
             -File "C:\Scripts\Invoke-ProfilePurge.ps1"
             -ComputerList "C:\Scripts\servers.txt"
             -DaysInactive 90 -PurgeProfileListBak -StopWSearch
             -WriteEventLog -LogPath "C:\Logs\ProfilePurge"
Exécuter   : SYSTEM  (ou admin élevé — SeRestorePrivilege requis)
```

---

## Limites

- Les profils sans `ntuser.dat` et sans `LastUseTime` sont ignorés, sauf avec `-DeleteUnknownDate`.
- `-Parallel` nécessite PowerShell 7+.
- Chaque cible nécessite WinRM et une session administrateur élevée (`SeRestorePrivilege`).
- La source du journal d'événements doit être créée une fois, avec les droits administrateur.

---

## Structure du projet

```
Invoke-ProfilePurge/
├── Invoke-ProfilePurge.ps1   # Script principal (EN)
├── README.md                 # Documentation (EN)
├── README.fr.md              # Documentation française
├── CHANGELOG.md              # Historique des versions
├── LICENSE                   # Licence MIT
├── .gitignore
├── servers.txt               # (non suivi) liste de serveurs
└── whitelist.txt             # (non suivi) liste d'exclusions
```

---

## Contribuer

1. Forker le dépôt
2. Créer une branche (`git checkout -b feature/ma-fonctionnalite`)
3. Commiter (`git commit -m 'feat: add ma-fonctionnalite'`)
4. Pousser la branche (`git push origin feature/ma-fonctionnalite`)
5. Ouvrir une Pull Request

Merci de suivre les [Conventional Commits](https://www.conventionalcommits.org/) pour les messages de commit.

---

## Licence

Ce projet est distribué sous licence MIT. Voir le fichier [LICENSE](LICENSE).

---

Maintenu par **9 Lives IT Solutions** — Informatique de santé & automatisation d'infrastructure.
