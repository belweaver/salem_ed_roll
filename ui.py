import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from excel_handler import load_excel_file
from dice_roller import lancer_des
from globals import globals
from autocomplete_combobox import AutocompleteCombobox

def create_ui(root):
    print("Création de l'interface utilisateur...")

    # Créer un bouton ouvrir un fichier excel
    open_file_button = ttk.Button(root, text="Ouvrir un fichier Excel", command=load_excel_file)
    open_file_button.pack(pady=10, padx=20, fill=tk.X)

    # Créer un sélecteur de dé de karma
    karma_label = ttk.Label(root, text="Sélecteur de dé de karma:")
    karma_label.pack(pady=5, padx=20)
    globals.karma_combo = ttk.Combobox(root, values=["D8", "D10", "D12"], state='readonly')
    globals.karma_combo.current(0)
    globals.karma_combo.pack(pady=5, padx=20)

    # Créer un cadre pour les boutons permanents
    globals.permanent_buttons_frame = ttk.Frame(root)
    globals.permanent_buttons_frame.pack(pady=10, padx=20, fill=tk.X)

    # Créer un champ de saisie avec typeahead pour sélectionner une compétence
    globals.talent_combo = AutocompleteCombobox(root, state='normal')
    globals.talent_combo.pack(pady=10, fill=tk.X, padx=20)
    globals.talent_combo.bind('<Return>', on_return)

    # Créer un label pour le message de résultat
    globals.result_label = ttk.Label(root, wraplength=450, justify='center')
    globals.result_label.pack(pady=20, padx=20)

    # Créer un label pour afficher le résultat principal en grand
    globals.result_value_label = ttk.Label(root, font=('Helvetica', 28), justify='center')
    globals.result_value_label.pack(pady=10, padx=20)

    # Créer un label pour afficher les détails
    globals.details_label = ttk.Label(root, wraplength=450, justify='center')
    globals.details_label.pack(pady=20, padx=20)

    # Créer un bouton lancer les dés
    launch_button = ttk.Button(root, text="Lancer les dés", command=lancer_des)
    launch_button.pack(pady=10, padx=20, fill=tk.X)

def on_return(event):
    lancer_des()