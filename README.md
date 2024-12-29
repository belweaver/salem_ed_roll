# salem_ed_roll
Lancer les dés Earthdawn à partir de la feuille Excel de Salem

## Logiciel simple qui vous permet de :
- ouvrir un fichier Excel
- trouver des talents et compétences dans les feuilles du fichier Excel
  - feuilles parcourues:
    - D1
    - D2
    - D3
    - CH
    - Comp
    - Passions+autres
- ajoute les talents qui se trouvent dans le fichier dice_roller_config.json
- ajouter des boutons rapides pour lancer les talents essentiels (ceux avec une valeur dans la colonne "Classification")
- ajouter un champ de saisie semi-automatique avec tous les talents trouvés dans le fichier Excel
- lancer des dés avec le mécanisme "d'explosion"
- demander si vous souhaitez ajouter un dé de karma (pour les talents qui le permettent) par une fenêtre popup
- afficher l'historique des 50 derniers lancers

## Utilisation :
Si vous voulez ajouter des lancés de dommages, placez un fichier dice_roller_config.json dans le même répertoire que roll.py
```
python roll.py
```

Windows : 
- télécharger dist/roll.exe
- Si vous voulez ajouter des lancés de dommages, placez un fichier dice_roller_config.json dans le même répertoire que roll.exe
- `roll.exe`

## Documentation :
- Dans le fichier Excel (le fiche de personnage), le chiffre indiqué dans la colonne "Classification" correspond à la colonne de boutons rapides créés automatiquement dans l'application
- Si la fiche indique plusieurs fois le même talent, le talent de discipline est utilisé de préférence. S'il y a plusieurs talents possibles amlgré tout, celui qui a le plus haut niveau total est utilisé 
- N'oubliez pas d'indiquer le dé de karma dans la liste déroulante
- Vous pouvez configurer certaines choses dans le fichier dice_roller_config.json, y compris ajouter des jets de dommages.

## Roadmap :
- lire le dé de karma dans la fiche
- coloriser les jets de dé (vert en cas d'explosition et rouge en cas de valeur minimale)
- compteurs essentiels (points de karma, de points de vie...)
