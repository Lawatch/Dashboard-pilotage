# Pilotage des prestations de propreté

| Fichier | Contenu |
|---|---|
| `Prestations_CDamien_8.xlsm` | **Le dashboard à utiliser** : DASHBOARD, CONTROLES, PARAMETRES, ABONNEMENTS + un onglet par jour. |
| `AUDIT.md` | Audit du prototype, erreurs trouvées dans les données, ce qui a changé et pourquoi. |
| `Prestations_CDamien_7.xlsm` | Prototype (conservé pour comparaison). |
| `Abonnements_iGclean071026.xlsx` | Fichier abonnements (projection du CA, comparaison Cegid). |
| `vba/` | Code des macros de la v8, lisible et versionné. |
| `outils/` | Scripts qui reconstruisent le classeur à partir du prototype, de l'export d'abonnements et de `vba/`. |

## Au quotidien

1. **Abonnements** (quand ils changent) : coller l'export dans l'onglet ABONNEMENTS, cellule A1.
2. **Chaque jour** : ajouter un onglet, y coller l'extract des validations. Il est renommé tout seul à la date (ex. `071026`).
3. **Revenir sur DASHBOARD** : tout se met à jour. Ce qui est à corriger est dans CONTROLES.

Après un téléchargement depuis GitHub, Windows bloque les macros : clic droit sur le fichier → Propriétés → cocher « Débloquer ».

## Reconstruire le classeur (après une modification de `vba/`)

Prérequis : Python 3 avec `openpyxl`, `olefile`, `oletools` ; Node.js (`npm install` dans `outils/vbaproject`) ; LibreOffice.

```bash
./outils/fabriquer.sh
```
