---
description: Boucle de vérification de la mise en forme au gabarit (onglet Gabarit) : mesurer, corriger le code, retester jusqu'à zéro défaut
argument-hint: SOURCE.docx notice|memoire [RESULTAT.docx]
---
Vérifier et améliorer la mise en forme au gabarit. Arguments : $ARGUMENTS (texte d'origine en .docx, profil, et si possible le document créé par le complément dans Word : Fichier, Enregistrer sous, Télécharger une copie).

Règles :
- Les .docx de Roman (nom, adresse, notice réelle) restent hors du dépôt : `tests/local/` (ignoré par Git) ou le dossier de la session. Jamais de commit d'un document réel.
- L'outil range et met en forme ; il n'écrit aucun texte. Une correction ne doit jamais ajouter, reformuler ou supprimer un mot du candidat (seuls les numéros proscrits par la Note et les numéros remplacés par la numérotation automatique du gabarit sont retirés).

Boucle (au plus 5 tours) :
1. Mesurer : `node outils/verifier-mise-en-forme.mjs SOURCE notice|memoire [RESULTAT] --pdf --sortie /tmp/vmf`. Sans RESULTAT, le document est composé par le même code que le volet ; avec RESULTAT, comparer aussi les deux (écarts propres à Word en ligne).
2. Lire le rapport : « Verdict » (défauts objectifs : texte perdu ou déplacé, style manquant, tableaux, page de garde, consignes restées) et « À examiner » (alertes Garde-fou plus nombreuses après mise en forme). Pour chaque point à examiner, ouvrir les paragraphes cités (`node -e` avec `lireDocx`) et trancher : défaut de `sources/composition.js`, lecture trompeuse d'une règle (`moteur/`), ou effet attendu (champs à mettre à jour dans Word, lien hypertexte à recréer).
3. Rendu : convertir le PDF en images (`pdftoppm -r 50 -png`) et regarder la page de garde, la table, une page de texte, un tableau.
4. Corriger le code (`sources/composition.js`, `moteur/`), ajouter un cas fictif reproduisant le défaut (`tests/fixtures/` ou test en ligne dans `tests/composition.test.js`), `npm test`.
5. Relancer l'étape 1. S'arrêter quand le verdict est « aucun défaut détecté » et que chaque point « à examiner » est expliqué.

Fin : `python3 outils/version.py X.Y.Z`, commit, push sur la branche de travail ; compte rendu à Roman en cinq lignes au plus : défauts trouvés, corrigés, restants (avec ce qu'il doit faire dans Word), et ce que la boucle ne peut pas voir (rendu exact de Word en ligne, images, liens).
