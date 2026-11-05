"""Recopie les consignes des gabarits Word (style « Consigne ») dans consignes.json.

Word en ligne ne laisse pas toujours voir le nom du style : le complément reconnaît ces paragraphes à leur texte.
Usage : python3 outils/extraire-consignes.py DOSSIER_DES_GABARITS (par défaut tests/fixtures/).
"""
import json
import sys
from pathlib import Path

from docx import Document

RACINE = Path(__file__).resolve().parents[1]
GABARITS = Path(sys.argv[1]) if len(sys.argv) > 1 else RACINE / "tests" / "fixtures"
SORTIE = RACINE / "consignes.json"

consignes = []
for g in sorted(GABARITS.glob("gabarit-*.docx")):
    for p in Document(g).paragraphs:
        if p.style.name == "Consigne" and p.text.strip() and p.text not in consignes:
            consignes.append(p.text)
SORTIE.write_text(json.dumps(consignes, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
print(f"{len(consignes)} consignes -> {SORTIE.relative_to(RACINE)}")
