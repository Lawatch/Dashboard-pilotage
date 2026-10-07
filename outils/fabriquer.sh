#!/bin/bash
# Reconstruit Prestations_CDamien_8.xlsm :
#   1. classeur + macros (onglets calculés vides)
#   2. exécution de « Mettre à jour » dans LibreOffice
#   3. classeur final, pré-rempli avec le résultat des macros
set -e
cd "$(dirname "$0")/.."
TMP=$(mktemp -d)
python3 outils/construire_classeur.py "$TMP/vide.xlsm"
python3 outils/executer_macros.py "$TMP/vide.xlsm" "$TMP/rempli.xlsx" Pipeline.MettreAJour
python3 outils/construire_classeur.py Prestations_CDamien_8.xlsm --remplir="$TMP/rempli.xlsx"
