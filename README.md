# Outil de Test PING

## Description

Ce fichier Excel au format XLSM est un outil permettant de réaliser des tests PING sur une liste d'hôtes (hostname ou adresses IP) et d'afficher les résultats directement dans le classeur. Les tests sont effectués via le protocole ICMP en utilisant WMI (Windows Management Instrumentation) pour une meilleure intégration et performance.

## Fonctionnalités

- **Test PING automatisé** : Lance des tests PING sur plusieurs hôtes en une seule opération
- **Utilisation de WMI** : Tests ICMP via Windows Management Instrumentation pour plus de fiabilité
- **Liste personnalisable** : Ajoutez facilement vos propres hostname ou adresses IP
- **Résultats en temps réel** : Les résultats s'affichent directement dans Excel
- **Statut de disponibilité** : Identifie rapidement les hôtes accessibles ou inaccessibles
- **Temps de réponse** : Affiche le temps de latence pour chaque hôte
- **Performance optimisée** : Utilisation de requêtes WMI natives Windows

## Prérequis

- **Microsoft Excel** avec support des macros (version 2010 ou ultérieure recommandée)
- **Macros activées** : Les macros doivent être autorisées pour que l'outil fonctionne
- **Système d'exploitation** : Windows (utilise WMI - Windows Management Instrumentation)
- **Droits d'accès** : Permissions réseau pour effectuer des requêtes ICMP via WMI
- **WMI actif** : Service WMI doit être en cours d'exécution sur le système

## Installation

1. Téléchargez le fichier `PingTool.xlsm`
2. Ouvrez le fichier avec Microsoft Excel
3. Activez les macros lorsque Excel vous le demande :
   - Cliquez sur **"Activer le contenu"** dans la barre d'avertissement de sécurité

## Utilisation

### Configuration de base

1. **Saisir les hôtes** :
   - Ouvrez la feuille principale
   - Entrez vos hostname ou adresses IP dans la colonne dédiée
   - Exemples : `google.com`, `192.168.1.1`, `server01.local`

2. **Lancer les tests** :
   - Cliquez sur le bouton "Calculer" ou exécutez la macro principale
   - Patientez pendant l'exécution des tests

3. **Consulter les résultats** :
   - Les résultats apparaissent dans les colonnes adjacentes
   - Codes couleur possibles pour une lecture rapide (vert = accessible, rouge = inaccessible)

## Structure du fichier

- **Colonne A** : Hostname ou adresse IP
- **Colonne B** : Statut (Accessible / Inaccessible / En cours...)

## Fonctionnement technique

### WMI et ICMP

L'outil utilise la classe WMI `Win32_PingStatus` qui :
- Envoie des paquets ICMP Echo Request (protocole ICMP)
- Récupère les réponses ICMP Echo Reply
- Fournit des informations détaillées (temps de réponse, TTL, code de statut)

## Dépannage

### Les macros ne fonctionnent pas
- Vérifiez que les macros sont activées dans les paramètres de sécurité Excel
- Accédez à : Fichier > Options > Centre de gestion de la confidentialité > Paramètres du Centre de gestion > Paramètres des macros

### Erreur "Accès WMI refusé"
- Vérifiez que le service WMI est démarré (services.msc > "Windows Management Instrumentation")
- Assurez-vous d'avoir les droits suffisants pour utiliser WMI
- Exécutez Excel en tant qu'administrateur si nécessaire

### Erreurs "Délai d'attente dépassé" (Code 11010)
- L'hôte peut être hors ligne ou inaccessible
- Vérifiez votre connexion réseau
- Certains pare-feu peuvent bloquer les requêtes ICMP
- L'hôte peut être configuré pour ignorer les requêtes ICMP

### Performances lentes
- Réduisez le nombre d'hôtes testés simultanément
- Certains hôtes peuvent avoir des temps de réponse élevés
- Ajustez le timeout WMI dans le code VBA si nécessaire

### WMI ne fonctionne pas
- Vérifiez que le service "Windows Management Instrumentation" est démarré
- Commande : `net start winmgmt` (en tant qu'administrateur)
- Vérifiez l'intégrité WMI : `winmgmt /verifyrepository`

## Sécurité

⚠️ **Important** :
- Ce fichier contient des macros VBA qui utilisent WMI pour effectuer des requêtes réseau
- N'ouvrez ce fichier que si vous en connaissez la provenance
- Les macros effectuent des requêtes ICMP via WMI (plus sécurisé que l'exécution de commandes système)
- Aucune commande système externe (cmd.exe) n'est exécutée

## Personnalisation

Le code VBA peut être modifié pour :
- Ajuster le timeout des requêtes WMI (par défaut 1000 ms)
- Modifier la taille des paquets ICMP envoyés
- Personnaliser le format d'affichage des résultats
- Ajouter des fonctionnalités supplémentaires (export, logs, historique, graphiques)
- Implémenter des tests récurrents automatiques

Pour accéder au code : `Alt + F11` dans Excel

### Exemple de requête WMI utilisée

```vba
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colPings = objWMIService.ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address = '" & hostname & "'")
```

## Limitations

- Nécessite une connexion réseau active
- Fonctionne uniquement sur Windows (utilise WMI spécifique à Windows)
- Les tests peuvent prendre du temps avec de nombreux hôtes
- Certains serveurs peuvent bloquer les requêtes ICMP
- WMI doit être actif et accessible sur le système
- Nécessite des droits suffisants pour utiliser WMI

## Avantages de WMI vs Commande PING

✅ **Avantages** :
- Pas besoin d'analyser la sortie texte de cmd.exe
- Résultats structurés et fiables
- Meilleure performance pour plusieurs tests
- Accès aux informations détaillées (TTL, taille buffer, etc.)
- Plus sécurisé (pas d'exécution de commandes système)

## Support et Contributions

Pour toute question, amélioration ou signalement de bug :
- Ouvrez une issue sur le dépôt du projet
- Soumettez vos pull requests
- Contactez le mainteneur du projet

Les contributions sont les bienvenues ! Merci de respecter la licence GPL lors de vos modifications.

## Licence

Ce projet est distribué sous licence **GNU General Public License v3.0 (GPL-3.0)**.

### En résumé :
- ✅ Utilisation libre (personnelle et commerciale)
- ✅ Modification du code autorisée
- ✅ Distribution autorisée
- ⚠️ Les modifications doivent également être distribuées sous GPL-3.0
- ⚠️ Le code source doit être disponible
- ⚠️ Les modifications doivent être documentées

Pour plus de détails, consultez le fichier LICENSE fourni avec ce projet ou visitez : https://www.gnu.org/licenses/gpl-3.0.html

---

**Version** : 1.1  
**Dernière mise à jour** : Octobre 2025  
**Auteur** : [micro-one.com](https://micro-one.com)
**Licence** : GPL-3.0  
**Technologie** : VBA + WMI (Windows Management Instrumentation) + ICMP

---

### Clause de non-responsabilité

Ce logiciel est fourni "tel quel", sans garantie d'aucune sorte. L'auteur ne peut être tenu responsable des dommages résultant de l'utilisation de cet outil.
