from tkinter import ttk
import tkinter as tk
from globals import globals
from dice_roller import lancer_des_for_talent

def update_talent_list(talent_cache):
    globals.talent_list = list(talent_cache.keys())
    globals.talent_list.sort()
    if globals.talent_combo is not None:
        globals.talent_combo.set_completion_list(globals.talent_list)
    else:
        print("Erreur : talent_combo n'est pas initialisé.")

def create_permanent_buttons():
    if globals.permanent_buttons_frame is None:
        print("Erreur : permanent_buttons_frame n'est pas initialisé.")
        return

    # Supprimer les anciens boutons permanents s'ils existent
    for widget in globals.permanent_buttons_frame.winfo_children():
        widget.destroy()

    # Créer des boutons permanents pour les talents avec une classification non zéro
    talents_with_classification = [(talent, data['classification']) 
                                 for talent, data in globals.talent_cache.items() 
                                 if isinstance(data['classification'], (int, float)) and data['classification'] > 0]
    talents_with_classification.sort(key=lambda x: (x[1], x[0]))

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
            button = ttk.Button(globals.permanent_buttons_frame, 
                              text=talent, 
                              command=lambda t=talent: lancer_des_for_talent(t))
            button.grid(row=row_num, column=col_num-1, pady=5, padx=10, sticky=tk.W+tk.E)