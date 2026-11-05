# Garde-fou DEC (complément Word)

Complément Word « Garde-fou DEC » : contrôles de forme de la demande d'agrément (notice) et du mémoire du diplôme d'expertise comptable, d'après la Note du jury aux candidats du 6 juillet 2026, la charte d'usage de l'IA et le document annexe de la demande d'agrément. Sans IA : le texte du document ne quitte pas Word ; le document n'est modifié que sur un clic « Corriger », avec aperçu ou confirmation, annulable par Ctrl+Z.

**Ce dépôt est la source du complément.** L'ancienne copie de `dec-memoire` (`scripts/garde-fou/`, recopiée par `publier.sh`) n'est plus à jour et ne doit plus être publiée ici.

## Installation dans Word

Accueil, Compléments, « Plus de compléments », « Mes compléments », « Charger mon complément », choisir `manifest.xml`. Le volet est servi par GitHub Pages depuis la branche `main`.

## Organisation

- `regles.json` : règles (base dans la Note, niveau par mode, « comment corriger »), paramètres, liste des exigences de la Note (contrôlées ou à vérifier soi-même).
- `consignes.json` : consignes des gabarits, reconnues et exclues des contrôles.
- `moteur/` : repérage de la structure et contrôles, sans dépendance au navigateur.
- `sources/` : lecture du document ouvert (Office.js) ou d'un .docx, et corrections dans Word.
- `taskpane.*` : volet.
