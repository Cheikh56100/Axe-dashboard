# NOVACAB V8 — Mode sombre professionnel

## Modifications

- Ajout d'un vrai thème **Sombre** depuis `Compte > Apparence` et `Compte > Préférences`.
- Le thème clair reste inchangé comme thème par défaut.
- Persistance du choix dans `localStorage` sous `novacab-theme`.
- Palette sombre alignée sur la maquette NOVACAB fournie : bleu nuit profond, cartes bleu ardoise, bordures bleutées discrètes, texte blanc cassé et accent bleu NOVACAB.
- Sidebar, barre supérieure, cartes, champs, listes, tableaux, badges, états actifs et composants de compte utilisent désormais les mêmes variables de thème.
- Suppression des principaux fonds blancs codés en dur afin d'éviter les ruptures visuelles en sombre.
- Contrôles `input/select/textarea` adaptés au mode sombre.
- Ombres et contrastes ajustés pour rester lisibles sans produire d'effet gris/blanc agressif.
- Le mode sombre ne modifie pas la logique métier ni les permissions.
