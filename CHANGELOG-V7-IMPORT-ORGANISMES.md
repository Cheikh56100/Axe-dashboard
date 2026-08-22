# NOVACAB V7 — Import Accès organismes sociaux

- Ajout du bouton **Télécharger le modèle Excel**.
- Le modèle contient deux feuilles : `Accès organismes` et `Instructions`.
- Ajout du bouton **Format attendu** avant tout import.
- Prévisualisation détaillée avant écriture en base.
- Rapprochement client par ordre : nom exact, SIREN, SIRET.
- Les colonnes obligatoires sont : `Organisme` + `Client/Dossier` ou `SIREN/SIRET`.
- Les colonnes `Libellé`, `Identifiant`, `Mot de passe`, `URL`, `Note` sont optionnelles.
- Les lignes non reconnues sont bloquées mais les lignes valides peuvent être importées.
- Les organismes non référencés sont conservés comme libellés personnalisés avec avertissement.
- Aucun accès n'est créé avant confirmation explicite.
