import re
import random
from tkinter import messagebox
from globals import globals

def lancer_des():
    talent = globals.talent_combo.get()
    if talent in globals.talent_cache:
        des = globals.talent_cache[talent]['des']
        karma = globals.talent_cache[talent]['karma']
        karma_dice = globals.karma_combo.get()

        if karma:
            add_karma = messagebox.askyesno("Ajouter un dé de karma", 
                                          f"Ajouter un dé de karma ({karma_dice}) ?")
        else:
            add_karma = False

        result, details = roll_dice(des, add_karma, karma_dice)
        globals.result_label.config(
            text=f'Résultat pour le talent "{talent}" (Niv. Tot. {int(globals.talent_cache[talent]["niv_tot"])}/{des}):')
        globals.result_value_label.config(text=f'{result}')
        globals.details_label.config(text=f'Détails: {details}')
    else:
        globals.result_label.config(text='Talent non trouvé dans le fichier.')
        globals.result_value_label.config(text='')
        globals.details_label.config(text='')

def roll_dice(des, add_karma=False, karma_dice=None):
    # Le reste de la fonction reste inchangé
    # [Garder le code existant de roll_dice]
    pass

def lancer_des_for_talent(talent):
    globals.talent_combo.set(talent)
    lancer_des()