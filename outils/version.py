"""Fixe le numéro de version partout : volet (taskpane.js), manifeste, et paramètre ?v= des fichiers chargés par le volet.

Le paramètre ?v= oblige Word à recharger ensemble tous les fichiers d'une même version (GitHub Pages les garde
en cache jusqu'à 10 minutes, chacun de son côté : sans lui, une page neuve peut tourner avec un ancien script).
Usage : python3 outils/version.py 0.6.2
"""
import re
import sys
from pathlib import Path

RACINE = Path(__file__).resolve().parents[1]
MODULES = ["sources/docx.js", "sources/office.js", "sources/correction.js", "sources/composition.js", "sources/zip.js", "sources/xml.js",
           "moteur/moteur.js", "moteur/controles.js", "moteur/structure.js", "moteur/texte.js"]

v = sys.argv[1]
assert re.fullmatch(r"\d+\.\d+\.\d+", v), "version attendue : 1.2.3"

js = RACINE / "taskpane.js"
js.write_text(re.sub(r"const VERSION = '[^']+';", f"const VERSION = '{v}';", js.read_text(encoding="utf-8")), encoding="utf-8")

man = RACINE / "manifest.xml"
man.write_text(re.sub(r"<Version>[^<]+</Version>", f"<Version>{v}.0</Version>", man.read_text(encoding="utf-8")), encoding="utf-8")

html = RACINE / "taskpane.html"
h = html.read_text(encoding="utf-8")
carte = "\n".join(f'      "./{m}": "./{m}?v={v}",' for m in MODULES).rstrip(",")
bloc = f'<script type="importmap">\n  {{\n    "imports": {{\n{carte}\n    }}\n  }}\n  </script>'
h = re.sub(r'<script type="importmap">.*?</script>', bloc, h, flags=re.S)
h = re.sub(r'href="taskpane\.css(\?v=[^"]*)?"', f'href="taskpane.css?v={v}"', h)
h = re.sub(r'src="taskpane\.js(\?v=[^"]*)?"', f'src="taskpane.js?v={v}"', h)
html.write_text(h, encoding="utf-8")
print("version", v)
