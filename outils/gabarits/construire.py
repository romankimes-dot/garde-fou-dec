"""Construit les deux gabarits Word du jury dans gabarits/ (source unique, décision du 05/11/2026).

Usage : python3 outils/gabarits/construire.py        (python-docx requis : pip install python-docx)
Les gabarits ne contiennent que des titres imposés par la Note du jury et des repères entre crochets.
Après régénération : vérifier dans Word (onglet Gabarit) puis refaire les documents d'essai
(python3 tests/fabriquer-fixtures.py) et les consignes (python3 outils/extraire-consignes.py gabarits).
"""
from pathlib import Path

from gabarits import gabarit_demande_agrement, gabarit_memoire

RACINE = Path(__file__).resolve().parents[2]
GABARITS = RACINE / "gabarits"

if __name__ == "__main__":
    GABARITS.mkdir(parents=True, exist_ok=True)
    for fabrique, chemin in (
        (gabarit_memoire, GABARITS / "gabarit-memoire.docx"),
        (gabarit_demande_agrement, GABARITS / "gabarit-demande-agrement.docx"),
    ):
        fabrique(chemin)
        print(chemin.relative_to(RACINE))
