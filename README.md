# Garde-fou DEC (complément Word)

Volet de Word (en ligne ou sur ordinateur) qui contrôle la demande d'agrément (notice) ou le mémoire du diplôme d'expertise comptable au regard de la Note du jury aux candidats du 6 juillet 2026, de la charte d'usage de l'IA et de son annexe n° 1, et du document annexe de la demande d'agrément. Aucune IA, aucun serveur : le texte reste dans Word et dans le navigateur. Le document n'est modifié que sur un clic « Corriger », avec aperçu ou confirmation, annulable par Ctrl+Z.

**Ce dépôt est la seule source du complément.** L'ancienne copie de `dec-memoire` (`scripts/garde-fou/`, `publier.sh`) est retirée.

## Installer dans Word (une fois)

Accueil, Compléments, « Plus de compléments », « Mes compléments », « Charger mon complément », choisir `manifest.xml`. Le volet est servi par GitHub Pages depuis la branche `main` (https://romankimes-dot.github.io/garde-fou-dec/). Après une mise à jour : clic droit dans le volet, « Actualiser » (cache de GitHub Pages : jusqu'à 10 minutes). Changer d'adresse : `python3 outils/construire-manifeste.py https://nouvelle/adresse`, puis recharger `manifest.xml`.

## Utiliser

- **Document** : Notice ou Mémoire ; **Mode** : Rédaction ou Dépôt (les exigences deviennent bloquantes, les restes de travail sont contrôlés).
- **Contrôler ce document** : lit le document ouvert, mise en page comprise. Un clic sur un extrait sélectionne le passage. **Contrôler un fichier .docx** : même contrôle sur une copie.
- Chaque règle : sa base (Note du jury, page citée, ou préférence), « Comment corriger », et dans Word les boutons **Corriger** (aperçu avant/après pour les changements de texte) et **Sélectionner pour trier** (tri de la bibliographie par Accueil, Trier).
- **Exigences de la Note** : toutes les exigences, contrôlées ou à vérifier soi-même, dépôt compris, et les codes du sujet.
- **Copier les alertes** : règles et messages, passages masqués, pour le calibrage.
- **Réglages** : termes du cabinet, des clients et des personnes (conservés dans ce navigateur seulement) ; occurrences ignorées ; diagnostic de lecture.

Le PDF exporté reste la référence pour la pagination exacte (`scripts/mise-en-forme/controle.py` dans `dec-memoire`).

## Contrôles

| Id | Contrôle |
| --- | --- |
| A01 à A05 | Volumes (notice ≤ 10 p., corps 95 à 110 p., annexes ≤ 100 p., note de synthèse 2 à 3 p.) ; équilibre (préférence) |
| A06 à A08 | Nomenclature Partie/Chapitre/Section/Paragraphe (I, 1.1, 1.1.1, A, a) proscrits) ; 2 ou 3 parties de 2 ou 3 chapitres, numérotation reprise ou continue sans mélange ; niveau isolé |
| A09 à A13 | Notice : sept sections dans l'ordre (titre approchant proposé), tableaux des sections 3 et 5, problématique en une question, plan détaillé (paragraphes, pages à chaque niveau, outils), plan synthétique |
| A14 à A17 | Pièces et ordre ; pages de garde (notice et mémoire) ; termes du cabinet ; annexe n° 1 de la charte IA (tableau des usages, exemple laissé, date, nom, signature) |
| A18 | Bibliographie : six catégories dans l'ordre, ordre alphabétique, présentation de chaque référence selon les exemples de la Note |
| A19, A20 | Mise en page (taille 12, interligne 1,5, marges, page de garde non numérotée, numéros en bas à droite) ; notes de bas de page |
| A21 | Tableau des changements (après 4.2) et note liminaire (deuxième présentation) |
| B01 à B10 | Sommaire et table des matières ; titres ; annonces de plan ; conclusions partielles ; renvois aux annexes ; sigles ; numéros des titres ; titres et origine des tableaux et figures |
| F01 à F08 | Restes de travail ; commentaires et révisions ; appariement ; typographie (insécables, guillemets, espaces, majuscules accentuées, nombres, chiffre en début de phrase) |

## Limites connues

- Lecture du document ouvert et corrections testées sur une imitation de Word (`tests/`), à confirmer dans Word (Réglages, Diagnostic, ligne `ooxml`).
- Non contrôlables : pagination réelle (le PDF fait foi), signature manuscrite, contenu des images, fond (contenu des sections, transitions, plagiat) ; le panneau « Exigences de la Note » les rappelle.

## Organisation

| Chemin | Rôle |
| --- | --- |
| `manifest.xml`, `taskpane.*`, `icones/` | Complément et volet |
| `regles.json` | Règles (base, niveau par mode, « comment corriger »), paramètres, exigences de la Note, codes du sujet |
| `consignes.json` | Consignes des gabarits, exclues des contrôles (`python3 outils/extraire-consignes.py DOSSIER_DES_GABARITS`) |
| `moteur/` | Structure (pièces, plan, annexes) et contrôles |
| `sources/` | Lecture du document ouvert (Office.js) ou d'un .docx ; corrections dans Word |
| `tests/` | Gabarits et documents fictifs (`fabriquer-fixtures.py`), modèle officiel de l'annexe n° 1, tests |

Tests : `npm test` (Node 22). Aperçu en console : `node tests/apercu.js FICHIER.docx notice depot`.
