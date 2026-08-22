# NOVACAB — V6 Social & imports

## Corrections
- Social & paie : bouton « Importer Excel/CSV » réellement fonctionnel.
- Social & paie : bouton « Nouvelle action » ouvre maintenant la création d'une tâche/action avec client, responsable, priorité et échéance.
- Import social : aperçu préalable, format attendu, contrôle des lignes et confirmation obligatoire avant écriture.
- Accès organismes sociaux : import du classeur avec aperçu préalable, règles de colonnes, lignes valides/à corriger et confirmation obligatoire.
- Accès organismes sociaux : rubrique visible par tous ; modification réservée aux Admin, Expert, Chef de mission et Gestionnaire de paie.
- Signature mail : option « Insérer automatiquement ma signature », éditable et mémorisée par utilisateur.
- Vue d'ensemble : nouveau graphique « Par catégorie fiscale » (BIC, BNC, BA, EI, IS…), avec filtre vers le registre.
- Fiche client : champ « Catégorie fiscale » avec auto-détection ou valeur explicite.
- Cotisations sociales : détection BTP prioritaire par code NAF 41/42/43, avec contrôles PRO BTP et CIBTP ajoutés aux vérifications proposées.

## Validation
Le projet est livré sous forme de source Vite/React. La commande `npm ci` n'a pas pu terminer dans l'environnement de génération (délai d'exécution), donc aucun build Vite complet n'est déclaré comme validé ici.
