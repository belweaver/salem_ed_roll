import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import random

class AutocompleteCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        """Use our completion list as our drop down selection menu, arrows move through menu."""
        self._completion_list = sorted(completion_list, key=str.lower)  # Work with a sorted list
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.bind('<KeyRelease>', self.handle_keyrelease)
        self['values'] = self._completion_list  # Setup our popup menu

    def autocomplete(self, delta=0):
        """autocomplete the Combobox, delta may be 0/1/-1 to cycle through possible hits"""
        if delta:  # need to delete selection otherwise we would fix the current position
            self.delete(self.position, tk.END)
        else:  # set position to end so selection starts where textentry ended
            self.position = len(self.get())
        # collect hits
        _hits = []
        for element in self._completion_list:
            if element.lower().startswith(self.get().lower()):  # Match case insensitively
                _hits.append(element)
        # if we have a new hit list, keep this in mind
        if _hits != self._hits:
            self._hit_index = 0
            self._hits = _hits
        # only allow cycling if we are in a known hit list
        if _hits == self._hits and self._hits:
            self._hit_index = (self._hit_index + delta) % len(self._hits)
        # now finally perform the auto completion
        if self._hits:
            self.delete(0, tk.END)
            self.insert(0, self._hits[self._hit_index])
            self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        """event handler for the keyrelease event on this widget"""
        if event.keysym == "BackSpace":
            self.delete(self.index(tk.INSERT), tk.END)
            self.position = self.index(tk.END)
        if event.keysym == "Left":
            if self.position < self.index(tk.END):  # delete the selection
                self.delete(self.position, tk.END)
            else:
                self.position = self.position - 1  # delete one character
                self.delete(self.position, tk.END)
        if event.keysym == "Right":
            self.position = self.index(tk.END)  # go to end (no selection)
        if len(event.keysym) == 1:
            self.autocomplete()
        # No need for up/down, we'll jump to the popup
        # list at the position of the autocompletion

# Fonction pour ouvrir un fichier excel et construire le cache de compétences
def load_excel_file():
    # Ouvrir un dialogue de fichier pour sélectionner le fichier excel
    file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")))
    if not file_path:
        return

    # Initialiser le cache des talents
    global talent_cache
    talent_cache = {}

    # Lire toutes les feuilles du fichier excel
    xls = pd.ExcelFile(file_path)

    # Zones des talents pour chaque feuille
    zones = {
        "Comp": [(4, 20), (35, 43), (52, 55)],
        "D1": [(5, 49)],
        "D2": [(5, 49)],
        "D3": [(5, 49)],
        "CH": [(5, 48)],
        "Passions+autres": [(4, 46), (82, 85), (90, 105), (106, 112)]
    }

    # Parcourir chaque feuille et charger les talents
    for sheet_name in xls.sheet_names:
        if sheet_name in zones:
            for zone in zones[sheet_name]:
                start, end = zone
                df = pd.read_excel(xls, sheet_name, skiprows=start-1, nrows=end-start+1, usecols="B:Q")

                # Nettoyer les noms de colonnes si ce sont des chaînes de caractères
                df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]

                # Vérifier si les colonnes nécessaires existent
                if 'Talents' in df.columns and 'Niv. Tot.' in df.columns and 'Dés' in df.columns and 'Classification' in df.columns:
                    for index, row in df.iterrows():
                        talent = row['Talents']
                        niv_tot = row['Niv. Tot.']
                        des = row['Dés']
                        classification = row['Classification']

                        # Filtrer les talents qui n'ont pas de valeur numérique dans la colonne Niv. Tot.
                        if pd.notna(talent) and pd.notna(niv_tot) and isinstance(niv_tot, (int, float)):
                            talent = talent.strip()
                            karma = bool(re.search(r' \(D\)$', talent))
                            talent_key = re.sub(r' \(D\)$', '', talent)
                            print(f"talent/talent_key/karma: {talent}/{talent_key}/{karma}")

                            # Vérifier si un talent avec " (D)" existe déjà
                            if not karma and f"{talent_key} (D)" in talent_cache:
                                print(f"talent_key: {talent_key} existe déjà avec karma")
                                continue

                            if pd.notna(des) and str(des).strip():
                                des = des.strip()
                                # Si le talent est déjà dans le cache, comparer les niveaux totaux
                                if talent in talent_cache:
                                    if niv_tot > talent_cache[talent]['niv_tot']:
                                        talent_cache[talent] = {'niv_tot': niv_tot, 'des': des, 'karma': karma, 'classification': classification}
                                        print(f"Mis à jour: Talent '{talent}' (Niv. Tot.: {niv_tot}, Dés: {des}, Karma: {karma}, Classification: {classification})")
                                else:
                                    talent_cache[talent] = {'niv_tot': niv_tot, 'des': des, 'karma': karma, 'classification': classification}
                                    print(f"Talent trouvé: '{talent}' (Niv. Tot.: {niv_tot}, Dés: {des}, Karma: {karma}, Classification: {classification})")
                else:
                    print(f"Les colonnes nécessaires ne sont pas présentes dans la feuille '{sheet_name}'.")

    # Mettre à jour le champ de typeahead avec la liste des talents
    update_talent_list(talent_cache)

    # Créer des boutons permanents pour les talents avec une classification non zéro
    create_permanent_buttons()

# Fonction pour mettre à jour le champ de typeahead à partir du cache des talents
def update_talent_list(talent_cache):
    global talent_list
    talent_list = list(talent_cache.keys())
    talent_list.sort()  # Classer les talents par ordre alphabétique
    talent_combo.set_completion_list(talent_list)

# Fonction pour lancer des dés en fonction du talent sélectionné dans le champ de typeahead
def lancer_des():
    talent = talent_combo.get()
    if talent in talent_cache:
        des = talent_cache[talent]['des']
        karma = talent_cache[talent]['karma']
        karma_dice = karma_combo.get()

        if karma:
            add_karma = messagebox.askyesno("Ajouter un dé de karma", f"Ajouter un dé de karma ({karma_dice}) ?")
        else:
            add_karma = False

        result, details = roll_dice(des, add_karma, karma_dice)
        result_label.config(text=f'Résultat pour le talent "{talent}" (Niv. Tot. {int(talent_cache[talent]["niv_tot"])}/{des}):')
        result_value_label.config(text=f'{result}')
        details_label.config(text=f'Détails: {details}')
    else:
        result_label.config(text='Talent non trouvé dans le fichier.')
        result_value_label.config(text='')
        details_label.config(text='')

# Fonction pour lancer le nombre de dés spécifié avec explosion
def roll_dice(des, add_karma=False, karma_dice=None):
    # Séparer les dés et les valeurs numériques
    dice_part = re.findall(r"(\d*D\d+)", des)
    numeric_part = None

    # Extraire la partie après le dernier "+"
    last_plus_index = des.rfind('+')
    if last_plus_index != -1:
        last_part = des[last_plus_index + 1:].strip()
        # Vérifier si la dernière partie contient un "D" ou "d"
        if 'D' in last_part or 'd' in last_part:
            numeric_part = None
        else:
            # Vérifier si la dernière partie est un nombre
            if last_part.isdigit():
                numeric_part = last_part

    result = 0
    details = []

    # Traiter les dés
    for dice in dice_part:
        rolls = re.findall(r"(\d*)D(\d+)", dice)
        for roll in rolls:
            num_dice = int(roll[0]) if roll[0] else 1
            faces = int(roll[1])
            for _ in range(num_dice):
                roll_result = random.randint(1, faces)
                result += roll_result
                details.append(f"{roll_result} (D{faces})")
                while roll_result == faces:
                    roll_result = random.randint(1, faces)
                    result += roll_result
                    details.append(f"EXP {roll_result} (D{faces})")

    # Ajouter les valeurs numériques à la fin
    if numeric_part:
        result += int(numeric_part)
        details.append(f"{numeric_part}")

    # Ajouter le dé de karma si nécessaire
    if add_karma and karma_dice:
        karma_faces = int(karma_dice[1:])
        karma_result = random.randint(1, karma_faces)
        result += karma_result
        details.append(f"+karma ({karma_dice}): {karma_result}")
        while karma_result == karma_faces:
            karma_result = random.randint(1, karma_faces)
            result += karma_result
            details.append(f"EXP {karma_result} (D{karma_faces})")

    return result, " + ".join(details)

# Fonction pour lancer les dés lorsque la touche [Return] est pressée
def on_return(event):
    lancer_des()

# Fonction pour créer des boutons permanents pour les talents avec une classification non zéro
def create_permanent_buttons():
    # Supprimer les anciens boutons permanents s'ils existent
    for widget in permanent_buttons_frame.winfo_children():
        widget.destroy()

    # Créer des boutons permanents pour les talents avec une classification non zéro
    talents_with_classification = [(talent, data['classification']) for talent, data in talent_cache.items() if isinstance(data['classification'], (int, float)) and data['classification'] > 0]
    talents_with_classification.sort(key=lambda x: (x[1], x[0]))  # Trier par classification puis par ordre alphabétique

    # Organiser les talents en colonnes
    columns = {i: [] for i in range(1, 5)}
    for talent, classification in talents_with_classification:
        if classification <= 4:
            columns[classification].append(talent)

    # Limiter à 5 boutons par colonne
    for col in columns.values():
        col[:] = col[:5]

    # Créer les boutons
    for col_num, talents in columns.items():
        for row_num, talent in enumerate(talents):
            button = ttk.Button(permanent_buttons_frame, text=talent, command=lambda t=talent: lancer_des_for_talent(t))
            button.grid(row=row_num, column=col_num-1, pady=5, padx=10, sticky=tk.W+tk.E)

# Fonction pour lancer les dés pour un talent spécifique
def lancer_des_for_talent(talent):
    talent_combo.set(talent)
    lancer_des()

# Initialiser l’application
root = tk.Tk()
root.title('Lancer de dés')
root.geometry("600x600")

# Créer un bouton ouvrir un fichier excel
open_file_button = ttk.Button(root, text="Ouvrir un fichier Excel", command=load_excel_file)
open_file_button.pack(pady=10, padx=20, fill=tk.X)

# Créer un sélecteur de dé de karma
karma_label = ttk.Label(root, text="Sélecteur de dé de karma:")
karma_label.pack(pady=5, padx=20)
karma_combo = ttk.Combobox(root, values=["D8", "D10", "D12"], state='readonly')
karma_combo.current(0)
karma_combo.pack(pady=5, padx=20)

# Créer un cadre pour les boutons permanents
permanent_buttons_frame = ttk.Frame(root)
permanent_buttons_frame.pack(pady=10, padx=20, fill=tk.X)

# Créer un champ de saisie avec typeahead pour sélectionner une compétence
talent_combo = AutocompleteCombobox(root, state='normal')
talent_combo.pack(pady=10, fill=tk.X, padx=20)
talent_combo.bind('<Return>', on_return)

# Créer un label pour le message de résultat
result_label = ttk.Label(root, wraplength=450, justify='center')
result_label.pack(pady=20, padx=20)

# Créer un label pour afficher le résultat principal en grand
result_value_label = ttk.Label(root, font=('Helvetica', 28), justify='center')
result_value_label.pack(pady=10, padx=20)

# Créer un label pour afficher les détails
details_label = ttk.Label(root, wraplength=450, justify='center')
details_label.pack(pady=20, padx=20)

# Créer un bouton lancer les dés
launch_button = ttk.Button(root, text="Lancer les dés", command=lancer_des)
launch_button.pack(pady=10, padx=20, fill=tk.X)

# Initialiser le cache des compétences vide
talent_cache = {}
talent_list = []

# Lancer l’application
root.mainloop()
