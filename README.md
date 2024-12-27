# salem_ed_roll
Lancer les dés Earthdawn à partir de la feuille Excel de Salem

## Logiciel simple qui vous permet de :
- ouvrir un fichier Excel
- trouver des compétences dans les feuilles du fichier Excel
  - feuilles parcourues:
    - D1
    - D2
    - D3
    - CH
    - Passions+autres
- ajouter des boutons rapides pour lancer les compétences essentielles (celles avec une valeur dans la colonne "Classification")
- ajouter un champ de saisie semi-automatique avec toutes les compétences trouvées dans le fichier Excel
- lancer des dés avec le mécanisme "explosion"
- vous demander si vous souhaitez ajouter un dé de karma (pour les compétences qui le permettent)
- afficher l'historique des 50 derniers lancers

## Utilisation :
```
python roll.py
```

Windows : 
- télécharger dist/roll.exe
- `roll.exe`

## Documentation :
- Dans le fichier Excel (le fiche de personnage), le chiffre indiqué dans la colonne "Classification" correspond à la colonne de boutons rapides créés automatiquement dans l'application
- Les Compétences ne sont pas lues, seulement les talents
- Si la fiche indique plusieurs fois le même talent, le talent de discipline est utilisé de préférence. S'il y a plusieurs talents possibles amlgré tout, celui qui a le plus haut niveau total est utilisé 
- N'oubliez pas d'indiquer le dé de karma dans la liste déroulante

## Limitations actuelles:
- pas de jet de dommage
- pas de compétences

## Roadmap :
- lire les compétences
- lire le dé de karma dans la fiche
- coloriser les jets de dé (vert en cas d'explosition et rouge en cas de valeur minimale)
- possibilité de définir et enregistrer des boutons personnalisés de jet de dé (typiquement pour les jets de dommages, impossible de gérer toutes les possibilités envisageables avec les règles)
- compteurs essentiels (points de karma, de points de vie...)
