# Invoke-ProfilePurge

> **Script PowerShell de purge automatisée des profils locaux sur un parc de serveurs Windows.**

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)](https://github.com/PowerShell/PowerShell)
[![PS7 Parallel](https://img.shields.io/badge/PS7-Mode%20parall%C3%A8le-blueviolet?logo=powershell)](https://github.com/PowerShell/PowerShell)
[![Licence: MIT](https://img.shields.io/badge/Licence-MIT-green.svg)](LICENSE)

📖 [Read in English](README.md)

---

## Présentation

`Invoke-ProfilePurge` est un script PowerShell de production destiné aux administrateurs Windows qui doivent nettoyer régulièrement les profils utilisateurs obsolètes sur des serveurs (RDS, serveurs de fichiers, VDI, contrôleurs de domaine). Il combine trois phases de nettoyage, un système de journalisation structuré, des rapports HTML et une exécution parallèle optionnelle.

---

## Fonctionnalités

| Fonctionnalité | Détail |
|---|---|
| **Purge des profils** | Profils inactifs supprimés via `Win32_UserProfile` CIM (dossier + clé registre) |
| **Nettoyage des clés .bak** | Clés orphelines `ProfileList\*.bak` supprimées avec résolution SID → nom |
| **Nettoyage des dossiers BACKUP** | Dossiers `*BACKUP*` supprimés avec vérification de session active |
| **Réparation des doublons domaine** | Paires `username` / `username.DOMAINE` détectées et repontées |
| **Arrêt/démarrage WSearch** | Libère les verrous `ntuser.dat.LOG` avant la purge |
| **Exécution parallèle** | `ForEach-Object -Parallel` PS7 pour les parcs multi-serveurs |
| **Journalisation structurée** | `Write-Log` 6 niveaux, console colorée + fichier `.log` UTF-8 BOM |
| **Rapport HTML** | Rapport moderne thème clair avec KPI et tableaux filtrés |
| **Journal Windows** | EventID 4100 (synthèse) et 4199 (erreur critique) |
| **Mode WhatIf** | Simulation complète, aucune modification effectuée |
| **PS5.1 + PS7** | Compatibilité double, parallèle PS7 uniquement |

---

## Prérequis

- PowerShell **5.1+** (mode parallèle : **7.0+**)
- Droits **administrateur local** sur chaque serveur cible
- **WinRM** activé sur les cibles distantes (`Enable-PSRemoting`)
- Pour la création de la source Event Log : droits admin à la **première exécution uniquement**

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
# 1. Simulation d abord — toujours
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -WhatIf

# 2. Examiner le rapport HTML, puis exécution réelle
.\Invoke-ProfilePurge.ps1 -DaysInactive 90 -StopWSearch -LogPath C:\Logs

# 3. Purge complète — toutes les phases
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
| `-ComputerList` | `string` | — | Fichier texte de serveurs (1/ligne) |
| `-Credential` | `PSCredential` | — | Credential pour sessions WinRM |

### Purge des profils

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-DaysInactive` | `int` | `90` | Seuil d inactivité en jours |
| `-ExcludeUsers` | `string[]` | — | Exclusions inline (wildcards : `svc_*`) |
| `-ExcludeFile` | `string` | — | Fichier whitelist (1/ligne, wildcards OK) |
| `-DeleteUnknownDate` | `switch` | off | Supprimer les profils sans date (⚠ comptes service) |
| `-StopWSearch` | `switch` | off | Arrêter Windows Search avant la purge |
| `-RepairDomainDuplicates` | `switch` | off | Réparer les paires `user` / `user.DOMAINE` |

### Phases supplémentaires

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-PurgeProfileListBak` | `switch` | off | Supprimer les clés registre `.bak` orphelines |
| `-PurgeBackupFolders` | `switch` | off | Supprimer les dossiers `*BACKUP*` |
| `-UsersPath` | `string` | `C:\Users` | Chemin racine pour la recherche BACKUP |

### Exécution

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-Parallel` | `switch` | off | Traitement parallèle des serveurs (PS7+) |
| `-ThrottleLimit` | `int` | `5` | Serveurs simultanés maximum |
| `-WhatIf` | `switch` | off | Simulation — aucune modification |
| `-PassThru` | `switch` | off | Émettre les objets résultat dans le pipeline |

### Sortie

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-LogPath` | `string` | Dossier script | Destination log + rapport HTML |
| `-LogRetentionDays` | `int` | `30` | Rétention des logs en jours |
| `-ReportPath` | `string` | Auto | Chemin personnalisé du rapport HTML |
| `-WriteEventLog` | `switch` | off | Écrire dans le journal Windows |
| `-EventSource` | `string` | `ProfilePurge` | Nom de la source d événement |
| `-EventLogName` | `string` | `Application` | Journal cible |

---

## Codes de sortie

| Code | Signification |
|---|---|
| `0` | Succès — toutes les opérations terminées |
| `1` | Erreur critique — exception non gérée |
| `2` | Partiel — échecs WinRM ou erreurs par profil |

---

## Journal Windows (Event Log)

| EventID | Type | Déclencheur |
|---|---|---|
| `4100` | Information / Warning / Error | Synthèse de fin d exécution |
| `4199` | Error | Exception critique non gérée |

```powershell
# Première exécution : créer la source (une seule fois, en tant qu administrateur)
New-EventLog -LogName Application -Source ProfilePurge
```

---

## Tâche planifiée

```
Programme : powershell.exe  (ou pwsh.exe pour PS7 parallèle)
Arguments :
  -NonInteractive -NoProfile -ExecutionPolicy Bypass
  -File "C:\Scripts\Invoke-ProfilePurge.ps1"
  -ComputerList "C:\Scripts\servers.txt"
  -DaysInactive 90 -PurgeProfileListBak -StopWSearch
  -WriteEventLog -LogPath "C:\Logs\ProfilePurge"
Exécuter en tant que : SYSTEM ou compte de service avec admin local sur les cibles
```

---

## Licence

MIT — voir [LICENSE](LICENSE).

---

*9 Lives IT Solutions — Informatique de santé & Automatisation d infrastructure*
