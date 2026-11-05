# Garde-fou DEC (complément Word)

Volet de Word (en ligne ou sur ordinateur) qui contrôle la demande d'agrément (notice) ou le mémoire du diplôme d'expertise comptable au regard de la Note du jury aux candidats du 6 juillet 2026, de la charte d'usage de l'IA et de son annexe n° 1, et du document annexe de la demande d'agrément. Aucune IA, aucun serveur : le texte reste dans Word et dans le navigateur. Le document n'est modifié que sur un clic « Corriger », avec aperçu ou confirmation, annulable par Ctrl+Z.

**Ce dépôt est la seule source du complément.** L'ancienne copie de `dec-memoire` (`scripts/garde-fou/`, `publier.sh`) est retirée.

## Installer dans Word (une fois)

Accueil, Compléments, « Plus de compléments », « Mes compléments », « Charger mon complément », choisir `manifest.xml`. Le volet est servi par GitHub Pages depuis la branche `main` (https://romankimes-dot.github.io/garde-fou-dec/). Après une mise à jour : Word en ligne, recharger la page du navigateur (F5) puis rouvrir le complément ; Word sur ordinateur, clic droit dans le volet, « Actualiser » (cache de GitHub Pages : jusqu'à 10 minutes ; la version s'affiche en haut du volet). Changer d'adresse : `python3 outils/construire-manifeste.py https://nouvelle/adresse`, puis recharger `manifest.xml`.

## Utiliser

- **Document** : demande d'agrément (notice) ou mémoire. Un seul contrôle : exigences de la Note bloquantes, préférences et typographie à vérifier, restes de travail (consignes, repères, TODO, commentaires, révisions) comptés à part, à nettoyer avant l'export PDF.
- **Contrôler ce document** : lit le document ouvert, mise en page comprise. Un clic sur un extrait sélectionne le passage. **Contrôler un fichier .docx** : même contrôle sur une copie.
- Chaque règle : sa base (Note du jury, page citée, ou préférence), « Comment corriger », et dans Word les boutons **Corriger** (aperçu avant/après pour les changements de texte) et **Sélectionner pour trier** (tri de la bibliographie par Accueil, Trier).
- **Gabarit** (mise en forme) : le texte collé sans mise en forme dans le document ouvert est versé dans un **nouveau document** créé depuis le gabarit du jury ; le document d'origine n'est pas modifié et l'outil n'écrit aucun texte. Les styles voyagent avec le document (rien à créer dans Word) :

  | Élément reconnu | Style appliqué |
  | --- | --- |
  | Page de garde | Page de garde (champs de l'onglet) |
  | Titres du sommaire et de la table des matières | Titre hors table (absent de la table) |
  | Notice explicative, plans, bibliographie, liste des annexes, introduction, conclusion… | Titre hors plan (niveau 1) |
  | Section de la notice « 1. … » à « 7. … » | Titre notice (niveau 2) |
  | Sous-titre de la notice (numéro « 1.1 » retiré) | Sous-titre notice (niveau 3) |
  | Partie, Chapitre, Section, § du mémoire | Titre 1 à 4, numéros automatiques « Partie 1 – » |
  | Catégorie de la bibliographie | Titre catégorie |
  | Annexe | Titre annexe, numéro automatique « Annexe 1 – » |
  | Plan synthétique et plan détaillé de la notice (lignes « Partie n – … (n pages) ») | Tableau « intitulé / pages » : introduction, conclusion et parties en gras sur gris foncé, chapitres en gras sur gris clair, sections en retrait, § en italique, outil souligné sous son § ; le nombre de pages tapé passe dans la colonne de droite |
  | Tableau tapé en texte (cellules séparées par « \| ») | Vrai tableau Word, première ligne en en-tête grisé répété, colonnes au prorata du texte |
  | Légende « Tableau n – … », « Figure n – … » | Légende (liée au tableau) |
  | « Source : … », « Réalisé par l'auteur » | Source (juste sous le tableau, liée à sa dernière ligne) |
  | Ligne commençant par un tiret ou une puce | Liste tirets (retrait suspendu, tiret conservé) |
  | Référence de la bibliographie | Référence bibliographique (retrait suspendu) |
  | Texte | Normal (taille 12, interligne 1,5) |

  Couleurs : titres de pièce, sections de la notice, annexes, parties et chapitres en bleu foncé `#1F3864` ; sous-titres, catégories de la bibliographie, sections et § du mémoire en bleu moyen `#2E5C8A` ; fonds des tableaux `#DCE6F2` (en-têtes, parties) et `#F2F2F2` (chapitres) ; texte courant en noir. Jamais de rouge, réservé aux modifications après une décision 4.2 (Note, p. 13).

  Repères « [Tableau n à insérer ici] » (style Consigne, signalés par F01 tant qu'ils restent) : légende « Tableau n – … » sans tableau à sa suite, section 3 ou 5 de la notice sans tableau, annexe n° 1 de la charte IA absente. L'outil n'écrit que le repère, jamais le contenu.

  Tables pré-remplies (entrées et pages estimées, annexes comprises) ; dans le nouveau document, Garde-fou propose « Corriger : mise à jour des tables » pour les pages exactes. Liens hypertextes et images non repris (signalés avant création). Gabarits vierges téléchargeables depuis le même onglet.
- **Exigences de la Note** : toutes les exigences, contrôlées ou à vérifier soi-même, dépôt compris, et les codes du sujet. Chaque alerte est rattachée à une seule exigence ; une exigence à vérifier soi-même se coche « Marquer vérifié » (mémorisé pour ce document, dans ce navigateur).
- **Copier les alertes** : règles et messages, passages masqués, pour le calibrage.
- **Réglages** : termes du cabinet, des clients et des personnes (conservés dans ce navigateur seulement) ; occurrences ignorées ; diagnostic de lecture.

Le PDF exporté reste la référence pour la pagination exacte (`scripts/mise-en-forme/controle.py` dans `dec-memoire`).

## Contrôles

| Id | Contrôle |
| --- | --- |
| A01 à A05 | Volumes (notice ≤ 10 p., corps 95 à 110 p., annexes ≤ 100 p., note de synthèse 2 à 3 p.) ; équilibre (préférence) |
| A06 à A08 | Nomenclature Partie/Chapitre/Section/Paragraphe (I, 1.1, 1.1.1, A, a) proscrits) ; 2 ou 3 parties de 2 ou 3 chapitres, numérotation reprise ou continue sans mélange ; niveau isolé |
| A09 à A13 | Notice : sept sections dans l'ordre (titre approchant proposé), tableaux des sections 3 et 5 (tableau tapé en « \| » signalé), problématique en une question, plan détaillé (paragraphes, pages à chaque niveau, outils), plan synthétique ; plans lus en paragraphes ou en tableau « intitulé / pages » |
| A14 à A17 | Pièces et ordre ; pages de garde (notice et mémoire) ; termes du cabinet ; annexe n° 1 de la charte IA (tableau des usages, exemple laissé, date, nom, signature) |
| A18 | Bibliographie : six catégories dans l'ordre, ordre alphabétique, présentation de chaque référence selon les exemples de la Note |
| A19, A20 | Mise en page (taille 12, interligne 1,5, marges, page de garde non numérotée, numéros en bas à droite) ; notes de bas de page |
| A21 | Tableau des changements (après 4.2) et note liminaire (deuxième présentation) |
| B01 à B10 | Sommaire et table des matières ; titres ; annonces de plan ; conclusions partielles ; renvois aux annexes ; sigles ; numéros des titres ; titres et origine des tableaux et figures |
| F01 à F08 | Restes de travail ; commentaires et révisions ; appariement ; typographie (insécables, guillemets, espaces, majuscules accentuées, nombres, chiffre en début de phrase) |

## Limites connues

- Lecture du document ouvert et corrections testées sur une imitation de Word (`tests/`), à confirmer dans Word (Réglages, Diagnostic, ligne `ooxml`).
- Mise en forme au gabarit : liens hypertextes et images non repris (signalés avant création) ; pages des tables estimées jusqu'à leur mise à jour par Word.
- Non contrôlables : pagination réelle (le PDF fait foi), signature manuscrite, contenu des images, fond (contenu des sections, transitions, plagiat) ; le panneau « Exigences de la Note » les rappelle.

## Organisation

| Chemin | Rôle |
| --- | --- |
| `manifest.xml`, `taskpane.*`, `icones/` | Complément et volet |
| `regles.json` | Règles (base, niveau, « comment corriger »), paramètres, exigences de la Note, codes du sujet |
| `consignes.json` | Consignes des gabarits, exclues des contrôles (`python3 outils/extraire-consignes.py DOSSIER_DES_GABARITS`) |
| `moteur/` | Structure (pièces, plan, annexes) et contrôles |
| `sources/` | Lecture du document ouvert (Office.js) ou d'un .docx ; corrections dans Word |
| `gabarits/` | Gabarits Word au format du jury : **source unique** (décision du 05/11/2026), régénérés par `python3 outils/gabarits/construire.py` (python-docx) ; dec-memoire n'en garde plus de copie |
| `tests/` | Gabarits et documents fictifs (`fabriquer-fixtures.py`), modèle officiel de l'annexe n° 1, tests |

Nouvelle version : `python3 outils/version.py 1.2.3` (volet, manifeste et paramètre `?v=` de chaque fichier chargé, pour que Word ne mélange jamais deux versions en cache).

Vérifier la mise en forme au gabarit sans Word : `node outils/verifier-mise-en-forme.mjs SOURCE.docx notice|memoire [RESULTAT.docx] [--pdf]` (texte conservé dans l'ordre, styles, tableaux, page de garde, consignes restées, alertes de Garde-fou apparues, rendu LibreOffice). Boucle de correction avec Claude Code : commande `/verifier-gabarit` (`.claude/commands/verifier-gabarit.md`) ; documents réels dans `tests/local/`, ignoré par Git.

Tests : `npm test` (Node 22). Aperçu en console : `node tests/apercu.js FICHIER.docx notice depot`.
