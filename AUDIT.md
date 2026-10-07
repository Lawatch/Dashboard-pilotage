# Audit du dashboard de pilotage — et ce qui change dans la version 8

Fichiers étudiés : `Prestations_CDamien_7.xlsm` (prototype) et `Abonnements_iGclean071026.xlsx` (abonnements).
Nouvelle version : `Prestations_CDamien_8.xlsm`.

---

## 1. En bref

Le prototype contient déjà ce qu'il faut : intégration des extracts jour par jour, rattachement aux abonnements, carte, contrôles.
Mais **plusieurs chiffres affichés sont faux**, l'information « qui fait quoi » n'existe pas, et la manipulation demande trop de gestes
(nommer les onglets à la main, cliquer « Mettre à jour », naviguer entre 6 onglets).

La version 8 garde la logique et la carte, corrige les calculs, et se réduit à **4 onglets + les onglets sources** :

| Onglet | Rôle |
|---|---|
| **DASHBOARD** | Indicateurs de la période, tableau par agent, tableau par jour. Clic sur un agent → **sa fiche** : ses chantiers et abonnements (fréquence, CA, sa part), ses prestations jour par jour **dans l'ordre des clics**. Bouton « Ouvrir la carte ». |
| **CONTROLES** | Tout ce qui est à corriger, avec un résumé chiffré (voir §4). |
| **PARAMETRES** | Mode d'emploi (5 lignes), adresse de l'agence, agents comptés ou non, domiciles, passages agence matin / soir, rattachements manuels. |
| **ABONNEMENTS** | On colle l'export des abonnements en A1, tel quel (comme l'onglet Source du fichier Abonnements). |
| **Onglets jours** | Un onglet par jour, renommé **automatiquement** à la date lue dans les données (ex. `061026`). |

À votre question « et c'est tout a priori ? » : oui, plus **CONTROLES**. Les corrections ont besoin d'un endroit à elles :
mélangées au DASHBOARD, elles le rendraient illisible. Le DASHBOARD affiche seulement une ligne d'alerte qui renvoie vers CONTROLES.

---

## 2. Ce que les données révèlent (chiffres réels, du 21/09 au 06/10)

| Constat | Détail | Conséquence |
|---|---|---|
| **Le « 07/10 » est un doublon du 06/10** | L'onglet `071026` contient les prestations datées du 06/10 (ré-export avec 2 réaffectations : Nabil KHATER → Salim ROUABAH et Alexandre QUATELA). Le prototype prend la date dans le **nom** de l'onglet. | Le 07/10 affiché (619 prestations) est faux et le 06/10 est compté deux fois sur les périodes « 7 j / 30 j ». La v8 lit la date dans les lignes, garde la version la plus récente et met l'autre de côté (`061026 ancien`). |
| **Heures pointées gonflées par les conteneurs** | Gestion des conteneurs : 1er clic = sortie le matin, 2e clic = rentrée le soir. Sur la période, 1 465 h des 2 203 h « pointées » viennent des conteneurs. Le 06/10 : 770 h pointées affichées pour 362 h prévues. | Indicateur inutilisable. La v8 compte la durée prévue pour un conteneur pointé (253 h pointées le 06/10). |
| **149 prestations planifiées sur des agents fictifs** | « A SUPPRIMER », « Vincent CALABRO A SUPPRIMER », « Tournée remplacement n°2 » : 72 h sur 6 jours. | Personne n'y est affecté. Listées dans CONTROLES §3. |
| **Le n° de projet ne suffit pas pour recoller** | Seules 46 % des prestations ont un projet présent dans les abonnements : un chantier a souvent un projet par prestation (conteneurs, parties communes…) et un seul côté facturation. En ajoutant « même nom à moins de 500 m » et « nom = Client Intervention / N.Ref. », on atteint 95 %. | 58 chantiers restent sans abonnement (127 h sur la période), avec une **suggestion** d'abonnement pour la plupart. |
| **31 abonnements sans aucune prestation** | 10 142 € HT / mois. Ex. CITY LODGE (928 €), L'EMINENCE (793 €), LES JARDINS DE LA VENCE A / B / C (1 200 €) — leurs prestations existent sous d'autres n° de projet (à recoller). | À vérifier un par un (oubli de planning, projet différent, résiliation). |
| **Doublons de facturation possibles** | LE LIBERTE (20260188 et 20260618, même projet, 257,82 €) ; RÉSIDENCE PARC AVENIR (20260260, 20260623, 20260606 : 3 × 455,40 €) ; L'ENTRE 2 PARCS (20200002 et 20260016, 475,99 €). ACCRETECH (20260610) à 0 €. | Listés dans CONTROLES §5. |
| **Coordonnées GPS fausses** | LE CARRÉ D'OR placé à Paris, LES BALCONS DU CIEL à Bordeaux, 6 à 16 RUE GALILÉE à La Rochelle. | Absents de la carte et des km : à corriger dans la fiche projet de l'outil. |
| **6 jours ouvrés sans extract** | Du 25/09 au 02/10. | Listés dans CONTROLES §6. |
| **Clics « à la chaîne »** | 45 prestations (13 agents) pointées début **et** fin aux mêmes minutes que d'autres prestations du même agent — ex. Sébastien MEURIN (15), Abdelkrim HAOUA le 06/10 : 3 chantiers différents tous pointés 04:06 → 05:07. Impossible sur place. | Statut « Clics simultanés », listé dans CONTROLES §4 et visible dans la fiche de l'agent. |
| **Le taux de pointage progresse** | 18 % le 21/09 → 58 % le 06/10. Au total 206 pointages à vérifier (clics simultanés, début sans fin, durée moins du quart ou plus du triple du prévu). | Visible jour par jour sur le DASHBOARD, détail dans CONTROLES §4. |
| **Pas de géolocalisation des clics** | Les colonnes « GSM In / Out » sont vides dans tous les extracts : seules les coordonnées des chantiers existent. | On ne peut pas vérifier qu'un agent était sur place. Si l'outil peut exporter ces colonnes, on ajoutera le contrôle « clic loin du chantier ». |

Les montants d'abonnement au-delà de 999 € arrivent en texte (« 1 029,63 ») dans l'export : c'est géré.

---

## 3. Audit du prototype, élément par élément

| Élément du prototype | Verdict | Pourquoi / ce que fait la v8 |
|---|---|---|
| Onglets jours `JJMMAA` nommés à la main | **Gardé, automatisé** | On colle l'extract dans n'importe quel nouvel onglet : il est renommé d'après la date de ses lignes. Doublon identique → mis de côté ; nouvelle version d'un jour → remplace l'ancienne. |
| Onglet `aboJJMMAA` | **Remplacé** | Un onglet fixe **ABONNEMENTS** : on colle l'export en A1, comme dans le fichier Abonnements. |
| Bouton « Mettre à jour » | **Facultatif** | Tout se met à jour au retour sur DASHBOARD ou CONTROLES (et à l'ouverture) dès qu'une source a changé. Le bouton reste en secours. |
| Tuiles du DASHBOARD | **Gardées, corrigées** | Prestations, % pointées, heures prévues, heures pointées (conteneurs corrigés), chantiers servis, **CA mensuel**, km pointés. |
| Colonnes « Inclure / Agence matin / Agence soir » du tableau par agent | **Déplacées** | Ce sont des réglages, pas du pilotage : elles sont dans PARAMETRES (listes Oui / Non). |
| **Km planning** | **Supprimés** | Le planning n'est pas dans l'ordre réel : ce chiffre n'a pas de sens. |
| Km pointés | **Gardés** | Calculés dans l'ordre des clics (OpenStreetMap), départ domicile / agence selon PARAMETRES. |
| Graphiques « 30 derniers jours » | **Remplacés** | Par un tableau « Par jour » : prestations, % pointées (en couleur), heures, agents. Plus lisible, sans macro de recadrage. |
| « Qui fait quoi » | **Nouveau** | Fiche agent : ses chantiers et abonnements (n°, client, fréquence, CA mensuel, **sa part**, CA attribué, passages, dernier passage) et ses prestations jour par jour **dans l'ordre des clics** (heure de début / fin, durée pointée vs prévue, statut). |
| Onglet CA | **Fusionné** | CA par agent dans le tableau du DASHBOARD, CA par chantier dans la fiche agent. L'ancien libellé « CA HT mensuel… à facturer sur le mois » prêtait à confusion (c'est un CA mensualisé). |
| Onglet CONTROLE | **Enrichi** | 8 contrôles au lieu de 2, résumé en tête, suggestion de rattachement cliquable, la saisie d'un n° rattache immédiatement. |
| Onglet TOURNEES | **Masqué** | Technique (km par agent et par jour). |
| Onglet LISEZ-MOI | **Remplacé** | 5 lignes de mode d'emploi en haut de PARAMETRES. |
| Carte | **Conservée telle quelle** | Seule la mention « km planning » disparaît, puisqu'ils ne sont plus calculés. Les choix faits sur la carte (agence matin / soir, maison) reviennent dans PARAMETRES. |

---

## 4. CONTROLES : ce qu'on y trouve

Période de contrôle : les 28 derniers jours de données (réglable dans PARAMETRES).

1. **Prestations sans abonnement** — par projet / chantier, avec heures, agents, dernière date et une suggestion d'abonnement. On clique sur la suggestion, ou on saisit le n° en colonne J : le rattachement est enregistré dans PARAMETRES et appliqué tout de suite.
2. **Abonnements sans prestation** — triés par CA mensuel, avec la dernière prestation connue.
3. **Prestations non affectées** — planifiées sur un agent fictif.
4. **Pointages à vérifier** — clics simultanés, début sans fin, durée anormale (conteneurs exclus).
5. **Abonnements en double ou à 0 €**.
6. **Onglets jours** — onglets mis de côté (à supprimer), jours en double, jours ouvrés sans extract, onglets illisibles.
7. **Coordonnées GPS des chantiers** — absentes ou à plus de 60 km de l'agence.
8. **Rattachements manuels invalides** — n° saisi inconnu ou abonnement terminé.

---

## 5. Règles de calcul

- **Prestation pointée** : début et fin cliqués.
- **Heures pointées** : durée réelle ; pour la gestion des conteneurs (2 clics = sortie / rentrée), durée prévue.
- **Abonnement actif** : date de fin postérieure au début de la période de contrôle, date de début passée.
- **CA mensuel** : MOIS = montant, 4X ANS = montant ÷ 3, 2X ANS = montant ÷ 6.
- **CA d'un agent** : pour chaque chantier, CA mensuel × (ses heures prévues ÷ heures prévues du chantier), sur les 28 derniers jours. Les agents fictifs ne prennent pas de part.
- **Rattachement prestation → abonnement**, dans l'ordre : rattachement manuel, n° de projet, même nom de chantier à moins de 500 m d'une prestation déjà rattachée, nom = Client Intervention ou N.Ref. d'un seul abonnement.
- **Agents fictifs** : nom contenant SUPPRIMER, REMPLACEMENT, A DEFINIR ou NOUVEAU. Ils ne sont pas comptés dans les indicateurs. Les autres agents à exclure (encadrement…) se règlent dans PARAMETRES (« Compté dans les indicateurs » = Non) ; vos réglages du prototype ont été repris.

---

## 6. Utilisation au quotidien

1. **Abonnements** (quand ils changent) : copier l'export, le coller dans ABONNEMENTS en A1.
2. **Chaque jour** : ajouter un onglet (bouton +), coller l'extract des validations en A1.
3. **Revenir sur DASHBOARD** : tout se met à jour. Choisir un agent dans la liste (ou cliquer son nom) pour voir sa fiche ; « Dernier jour / 7 jours / 30 jours / Tout » pour la période.

---

## 7. Limites et suites possibles

- **À tester dans Excel chez vous.** Les macros ont été exécutées de bout en bout sur vos données dans LibreOffice ; Excel n'était pas disponible ici. Un fichier téléchargé depuis GitHub a ses macros bloquées : clic droit sur le fichier → Propriétés → cocher « Débloquer ».
- **Ne pas supprimer les anciens onglets jours** : l'historique est reconstruit à partir d'eux. Si leur nombre devient gênant, on ajoutera un archivage.
- **Km** : calculés par la route via OpenStreetMap (connexion internet nécessaire pour les nouveaux trajets ; les distances déjà connues sont mémorisées).
- **Géolocalisation des clics** : à ajouter si l'outil peut exporter les colonnes GSM.
- La projection du CA et la comparaison Cegid restent dans le fichier Abonnements.
