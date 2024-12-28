import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import random
import json
from typing import Dict, List, Tuple, Optional, Any

# Constants
EXCEL_ZONES = {
    "Comp": [(4, 20), (35, 43), (51, 55)],
    "D1": [(5, 49)],
    "D2": [(5, 49)],
    "D3": [(5, 49)],
    "CH": [(5, 48)],
    "Passions+autres": [(4, 46), (82, 85), (90, 105), (106, 112)]
}

SUBSTITUTE_TALENT_COLUMN_NAME = ['Connaissances', 'Compétences', 'Art / artisanat']
REQUIRED_COLUMNS = ['Talents', 'Niv. Tot.', 'Dés', 'Classification']
DEFAULT_KARMA_DIE = 'D12'
MAX_HISTORY_SIZE = 50

class ToolTip:
    def __init__(self, widget: tk.Widget, text: str):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25

        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")

        label = tk.Label(self.tooltip, text=self.text, 
                      justify=tk.LEFT, background="#ffffe0", 
                      relief=tk.SOLID, borderwidth=1)
        label.pack()

    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

class AutocompleteCombobox(ttk.Combobox):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._completion_list = []
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.bind('<KeyRelease>', self.handle_keyrelease)

    def set_completion_list(self, completion_list: List[str]):
        self._completion_list = sorted(completion_list, key=str.lower)
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self['values'] = self._completion_list

    def autocomplete(self, delta: int = 0):
        if delta:
            self.delete(self.position, tk.END)
        else:
            self.position = len(self.get())

        _hits = []
        for element in self._completion_list:
            if element.lower().startswith(self.get().lower()):
                _hits.append(element)

        if _hits != self._hits:
            self._hit_index = 0
            self._hits = _hits

        if _hits == self._hits and self._hits:
            self._hit_index = (self._hit_index + delta) % len(self._hits)

        if self._hits:
            self.delete(0, tk.END)
            self.insert(0, self._hits[self._hit_index])
            self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        if event.keysym == "BackSpace":
            self.delete(self.index(tk.INSERT), tk.END)
            self.position = self.index(tk.END)
        elif event.keysym == "Left":
            if self.position < self.index(tk.END):
                self.delete(self.position, tk.END)
            else:
                self.position = self.position - 1
                self.delete(self.position, tk.END)
        elif event.keysym == "Right":
            self.position = self.index(tk.END)
        elif len(event.keysym) == 1:
            self.autocomplete()

class TalentCache:
    def __init__(self):
        self._cache: Dict[str, Dict[str, Any]] = {}

    def add_talent(self, talent: str, niv_tot: float, des: str, 
                  karma: bool, classification: float) -> bool:
        if talent in self._cache:
            if niv_tot <= self._cache[talent]['niv_tot']:
                return False
        self._cache[talent] = {
            'niv_tot': niv_tot,
            'des': des,
            'karma': karma,
            'classification': classification
        }
        return True

    def get_talent(self, talent: str) -> Optional[Dict[str, Any]]:
        return self._cache.get(talent)

    def get_all_talents(self) -> List[str]:
        return sorted(self._cache.keys())

    def get_classified_talents(self) -> List[Tuple[str, float]]:
        return [(talent, data['classification']) 
                for talent, data in self._cache.items() 
                if isinstance(data['classification'], (int, float)) 
                and data['classification'] > 0]

class Config:
    def __init__(self, filename: str = 'dice_roller_config.json'):
        self.filename = filename
        self.data = self.load()

    def load(self) -> Dict[str, Any]:
        try:
            with open(self.filename, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            return self.get_default_config()

    def save(self):
        with open(self.filename, 'w') as f:
            json.dump(self.data, f)

    @staticmethod
    def get_default_config() -> Dict[str, Any]:
        return {
            'default_karma_die': DEFAULT_KARMA_DIE,
            'max_history_size': MAX_HISTORY_SIZE,
            'ui_theme': 'default',
            'window_size': '600x800'
        }

class DiceRollerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.config = Config()
        self.talent_cache = TalentCache()
        self.roll_history: List[Dict[str, Any]] = []
        
        self.setup_window()
        self.setup_ui()
        self.add_tooltips()

    def setup_window(self):
        self.root.title('Lanceur de dés amélioré')
        self.root.geometry(self.config.data['window_size'])

    def on_return(self, event):
        self.lancer_des()

    def setup_ui(self):
        # Frame principal
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Bouton pour ouvrir un fichier Excel
        self.open_file_button = ttk.Button(
            self.main_frame, 
            text="Ouvrir un fichier Excel", 
            command=self.load_excel_file
        )
        self.open_file_button.pack(pady=10, fill=tk.X)

        # Sélecteur de dé de karma
        karma_frame = ttk.LabelFrame(self.main_frame, text="Configuration du karma")
        karma_frame.pack(pady=10, fill=tk.X)
        
        self.karma_combo = ttk.Combobox(
            karma_frame, 
            values=["D8", "D10", "D12"], 
            state='readonly'
        )
        self.karma_combo.set(self.config.data['default_karma_die'])
        self.karma_combo.pack(pady=5, padx=5, fill=tk.X)

        # Frame pour les boutons permanents
        self.permanent_buttons_frame = ttk.LabelFrame(
            self.main_frame, 
            text="Talents fréquents"
        )
        self.permanent_buttons_frame.pack(pady=10, fill=tk.X)

        # Champ de saisie avec autocomplétion
        self.talent_combo = AutocompleteCombobox(self.main_frame)
        self.talent_combo.pack(pady=10, fill=tk.X)
        self.talent_combo.bind('<Return>', self.on_return)

        # Labels de résultat
        self.result_frame = ttk.LabelFrame(self.main_frame, text="Résultat")
        self.result_frame.pack(pady=10, fill=tk.X)

        self.result_label = ttk.Label(
            self.result_frame, 
            wraplength=450, 
            justify='center'
        )
        self.result_label.pack(pady=5)

        self.result_value_label = ttk.Label(
            self.result_frame, 
            font=('Helvetica', 28), 
            justify='center'
        )
        self.result_value_label.pack(pady=5)

        self.details_label = ttk.Label(
            self.result_frame, 
            wraplength=450, 
            justify='center'
        )
        self.details_label.pack(pady=5)

        # Bouton de lancer
        self.launch_button = ttk.Button(
            self.main_frame, 
            text="Lancer les dés", 
            command=self.lancer_des
        )
        self.launch_button.pack(pady=10, fill=tk.X)

        # Historique des lancers
        self.history_frame = ttk.LabelFrame(
            self.main_frame, 
            text="Historique des lancers"
        )

        self.history_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        # Creation du treeview pour l'historique
        self.history_tree = ttk.Treeview(
            self.history_frame,
            columns=('Talent', 'Résultat', 'Détails', 'Date'),
            show='headings'
        )
        
        # Configuration des colonnes
        self.history_tree.heading('Talent', text='Talent')
        self.history_tree.heading('Résultat', text='Résultat')
        self.history_tree.heading('Détails', text='Détails')
        self.history_tree.heading('Date', text='Date')
        
        self.history_tree.column('Talent', width=100)
        self.history_tree.column('Résultat', width=70)
        self.history_tree.column('Détails', width=200)
        self.history_tree.column('Date', width=100)
        
        # Ajout de la scrollbar
        history_scrollbar = ttk.Scrollbar(
            self.history_frame,
            orient=tk.VERTICAL,
            command=self.history_tree.yview
        )
        self.history_tree.configure(yscrollcommand=history_scrollbar.set)
        
        # Placement du treeview et de la scrollbar
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        history_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def add_tooltips(self):
        tooltips = {
            self.karma_combo: "Sélectionnez le type de dé de karma à utiliser",
            self.talent_combo: "Entrez ou sélectionnez un talent",
            self.launch_button: "Cliquez pour lancer les dés du talent sélectionné",
            self.open_file_button: "Ouvrir un fichier Excel contenant les talents"
        }
        
        for widget, text in tooltips.items():
            ToolTip(widget, text)

    def load_excel_file(self):
        try:
            file_path = filedialog.askopenfilename(
                filetypes=(
                    ("Excel files", "*.xlsx *.xls *.xlsm"),
                    ("All files", "*.*")
                )
            )
            
            if not file_path:
                return

            xls = pd.ExcelFile(file_path)
            self.talent_cache = TalentCache()  # Reset cache

            for sheet_name, zones in EXCEL_ZONES.items():
                if sheet_name not in xls.sheet_names:
                    continue

                for zone in zones:
                    start, end = zone
                    df = pd.read_excel(
                        xls,
                        sheet_name,
                        skiprows=start-1,
                        nrows=end-start+1,
                        usecols="B:Q"
                    )

                    # Nettoyer les noms de colonnes
                    df.columns = ['Talents' if col.strip() in SUBSTITUTE_TALENT_COLUMN_NAME else col.strip() if isinstance(col, str) else col 
                                for col in df.columns]

                    if not all(col in df.columns for col in REQUIRED_COLUMNS):
                        print(f"Colonnes manquantes dans {sheet_name}")
                        continue

                    self.process_talent_data(df)

            self.update_talent_list()
            self.create_permanent_buttons()
            messagebox.showinfo("Succès", "Fichier Excel chargé avec succès!")
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du chargement: {str(e)}")

    def process_talent_data(self, df: pd.DataFrame):
        for _, row in df.iterrows():
            talent = row['Talents']
            niv_tot = row['Niv. Tot.']
            des = row['Dés']
            classification = row['Classification']

            if not self.validate_talent_data(talent, niv_tot, des, classification):
                continue

            talent = talent.strip()
            karma = bool(re.search(r' \(D\)$', talent))
            talent_key = re.sub(r' \(D\)$', '', talent)

            if not karma and f"{talent_key} (D)" in self.talent_cache.get_all_talents():
                continue

            if pd.notna(des) and str(des).strip():
                self.talent_cache.add_talent(
                    talent,
                    float(niv_tot),
                    str(des).strip(),
                    karma,
                    classification
                )

    def validate_talent_data(
        self,
        talent: Any,
        niv_tot: Any,
        des: Any,
        classification: Any
    ) -> bool:
        if pd.isna(talent) or pd.isna(niv_tot) or not isinstance(niv_tot, (int, float)):
            return False
        if pd.isna(des) or not isinstance(des, str):
            return False
        if not re.match(r'^(\d*D\d+\s*[\+\-]?\s*)*\d*$', str(des).strip()):
            return False
        return True

    def update_talent_list(self):
        talents = self.talent_cache.get_all_talents()
        self.talent_combo.set_completion_list(talents)

    def create_permanent_buttons(self):
        for widget in self.permanent_buttons_frame.winfo_children():
            widget.destroy()

        talents_with_classification = self.talent_cache.get_classified_talents()
        talents_with_classification.sort(key=lambda x: (x[1], x[0]))

        columns = {i: [] for i in range(1, 5)}
        for talent, classification in talents_with_classification:
            if classification <= 4:
                columns[classification].append(talent)

        for col in columns.values():
            col[:] = col[:5]

        for col_num, talents in columns.items():
            for row_num, talent in enumerate(talents):
                button = ttk.Button(
                    self.permanent_buttons_frame,
                    text=talent,
                    command=lambda t=talent: self.lancer_des_for_talent(t)
                )
                button.grid(row=row_num, column=col_num-1, pady=5, padx=10, sticky=tk.W+tk.E)

    def roll_dice(self, des: str, add_karma: bool = False, karma_dice: Optional[str] = None) -> Tuple[int, str]:
        def roll_single_die(faces: int) -> Tuple[int, List[str]]:
            total = 0
            rolls = []
            current_roll = random.randint(1, faces)
            total += current_roll
            rolls.append(f"{current_roll} (D{faces})")
            
            while current_roll == faces:  # Explosion
                current_roll = random.randint(1, faces)
                total += current_roll
                rolls.append(f"EXP {current_roll} (D{faces})")
                
            return total, rolls

        total_result = 0
        all_rolls = []
        
        # Parse and roll dice
        for match in re.finditer(r"(\d*)D(\d+)", des):
            count = int(match.group(1)) if match.group(1) else 1
            faces = int(match.group(2))
            
            for _ in range(count):
                result, rolls = roll_single_die(faces)
                total_result += result
                all_rolls.extend(rolls)

        # Add any fixed modifiers
        modifier_match = re.search(r'[\+\-]\s*\d+$', des)
        if modifier_match:
            modifier = int(modifier_match.group())
            total_result += modifier
            all_rolls.append(str(modifier))
        
        # Handle karma dice
        if add_karma and karma_dice:
            karma_faces = int(karma_dice[1:])
            karma_result, karma_rolls = roll_single_die(karma_faces)
            total_result += karma_result
            all_rolls.extend([f"+karma: {roll}" for roll in karma_rolls])
        
        return total_result, " + ".join(all_rolls)

    def lancer_des(self):
        talent = self.talent_combo.get()
        talent_data = self.talent_cache.get_talent(talent)
        
        if not talent_data:
            messagebox.showerror("Erreur", "Talent non trouvé dans le fichier.")
            return

        des = talent_data['des']
        karma = talent_data['karma']
        karma_dice = self.karma_combo.get() if karma else None

        if karma:
            add_karma = messagebox.askyesno(
                "Ajouter un dé de karma",
                f"Ajouter un dé de karma ({karma_dice}) ?"
            )
        else:
            add_karma = False

        result, details = self.roll_dice(des, add_karma, karma_dice)
        
        # Mise à jour des labels
        self.result_label.config(
            text=f'Résultat pour le talent "{talent}" '
                 f'(Niv. Tot. {int(talent_data["niv_tot"])}/{des}):'
        )
        self.result_value_label.config(text=str(result))
        self.details_label.config(text=f'Détails: {details}')
        
        # Ajout à l'historique
        from datetime import datetime
        self.add_to_history(talent, result, details, datetime.now().strftime("%H:%M:%S"))

    def add_to_history(self, talent: str, result: int, details: str):
            from datetime import datetime
        
            # Limiter la taille de l'historique
            if len(self.roll_history) >= self.config.data['max_history_size']:
                self.roll_history.pop(0)
                # Supprimer la première entrée du Treeview
                first_item = self.history_tree.get_children()[0]
                self.history_tree.delete(first_item)
        
            # Ajouter le nouveau lancer à l'historique
            current_time = datetime.now().strftime('%H:%M:%S')
            self.roll_history.append({
                'talent': talent,
                'result': result,
                'details': details,
                'time': current_time
            })
        
            # Ajouter l'entrée dans le Treeview
            self.history_tree.insert('', 0, values=(
                talent,
                str(result),
                details,
                current_time
             ))

    def load_excel_file(self):
        try:
                file_path = filedialog.askopenfilename(
                    filetypes=(
                        ("Excel files", "*.xlsx *.xls *.xlsm"),
                        ("All files", "*.*")
                    )
                )
                if not file_path:
                    return

                xls = pd.ExcelFile(file_path)
                self.process_excel_file(xls)
                self.update_talent_list()
                self.create_permanent_buttons()
            
                messagebox.showinfo(
                    "Succès",
                    "Le fichier a été chargé avec succès!"
                )
            
        except Exception as e:
                messagebox.showerror(
                    "Erreur",
                    f"Impossible de charger le fichier: {str(e)}"
                )

    def process_excel_file(self, xls: pd.ExcelFile):
        for sheet_name, zones in EXCEL_ZONES.items():
            if sheet_name not in xls.sheet_names:
                continue
                
            for start, end in zones:
                try:
                    df = pd.read_excel(
                        xls,
                        sheet_name,
                        skiprows=start-1,
                        nrows=end-start+1,
                        usecols="B:Q"
                    )
                    
                    if not all(col in df.columns for col in REQUIRED_COLUMNS):
                        print(f"Colonnes manquantes dans {sheet_name}")
                        continue
                        
                    self.process_talent_data(df)
                    
                except Exception as e:
                    print(f"Erreur lors du traitement de {sheet_name}: {str(e)}")

    def process_talent_data(self, df: pd.DataFrame):
        for _, row in df.iterrows():
            talent = row['Talents']
            niv_tot = row['Niv. Tot.']
            des = row['Dés']
            classification = row['Classification']

            if not self.validate_talent_data(talent, niv_tot, des, classification):
                continue

            talent = talent.strip()
            karma = bool(re.search(r' \(D\)$', talent))
            talent_key = re.sub(r' \(D\)$', '', talent)

            if not karma and f"{talent_key} (D)" in self.talent_cache._cache:
                continue

            if pd.notna(des) and str(des).strip():
                des = des.strip()
                self.talent_cache.add_talent(
                    talent,
                    float(niv_tot),
                    des,
                    karma,
                    classification
                )

    def validate_talent_data(self, talent, niv_tot, des, classification) -> bool:
        if not pd.notna(talent) or not isinstance(talent, str):
            return False
        if not pd.notna(niv_tot) or not isinstance(niv_tot, (int, float)):
            return False
        if not pd.notna(des) or not isinstance(des, str):
            return False
        if not re.match(r'^(\d*D\d+\s*[\+\-]?\s*)*\d*$', des.strip()):
            return False
        return True

    def update_talent_list(self):
        self.talent_combo.set_completion_list(self.talent_cache.get_all_talents())

    def create_permanent_buttons(self):
        # Nettoyer les boutons existants
        for widget in self.permanent_buttons_frame.winfo_children():
            widget.destroy()

        # Récupérer et trier les talents classifiés
        talents_with_classification = self.talent_cache.get_classified_talents()
        talents_with_classification.sort(key=lambda x: (-x[1], x[0]))

        # Créer les colonnes de boutons
        columns = {i: [] for i in range(1, 5)}
        for talent, classification in talents_with_classification:
            if classification <= 4:
                columns[int(classification)].append(talent)

        # Limiter à 5 boutons par colonne
        for col in columns.values():
            col[:] = col[:5]

        # Créer les boutons
        for col_num, talents in columns.items():
            for row_num, talent in enumerate(talents):
                button = ttk.Button(
                    self.permanent_buttons_frame,
                    text=talent,
                    command=lambda t=talent: self.lancer_des_for_talent(t)
                )
                button.grid(
                    row=row_num,
                    column=col_num-1,
                    pady=5,
                    padx=10,
                    sticky=tk.W+tk.E
                )

    def lancer_des_for_talent(self, talent: str):
        self.talent_combo.set(talent)
        self.lancer_des()

    def lancer_des(self):
        talent = self.talent_combo.get()
        talent_data = self.talent_cache.get_talent(talent)
        
        if not talent_data:
            self.update_result_labels(
                "Talent non trouvé dans le fichier.",
                "",
                ""
            )
            return

        karma = talent_data['karma']
        karma_dice = self.karma_combo.get() if karma else None

        if karma:
            add_karma = messagebox.askyesno(
                "Ajouter un dé de karma",
                f"Ajouter un dé de karma ({karma_dice}) ?"
            )
        else:
            add_karma = False

        result, details = self.roll_dice(
            talent_data['des'],
            add_karma,
            karma_dice
        )

        self.update_result_labels(
            f'Résultat pour le talent "{talent}" '
            f'(Niv. Tot. {int(talent_data["niv_tot"])}/{talent_data["des"]}):',
            str(result),
            f'Détails: {details}'
        )
        
        self.add_to_history(talent, result, details)

    def update_result_labels(self, result_text: str, value_text: str, details_text: str):
        self.result_label.config(text=result_text)
        self.result_value_label.config(text=value_text)
        self.details_label.config(text=details_text)

    def roll_dice(self, des: str, add_karma: bool = False, karma_dice: Optional[str] = None) -> Tuple[int, str]:
        def roll_single_die(faces: int) -> Tuple[int, List[str]]:
            total = 0
            rolls = []
            current_roll = random.randint(1, faces)
            total += current_roll
            rolls.append(f"{current_roll} (D{faces})")
            
            while current_roll == faces:  # Explosion
                current_roll = random.randint(1, faces)
                total += current_roll
                rolls.append(f"EXP {current_roll} (D{faces})")
                
            return total, rolls

        total_result = 0
        all_rolls = []
        
        # Traiter les dés standards
        for match in re.finditer(r"(\d*)D(\d+)", des):
            count = int(match.group(1)) if match.group(1) else 1
            faces = int(match.group(2))
            
            for _ in range(count):
                result, rolls = roll_single_die(faces)
                total_result += result
                all_rolls.extend(rolls)
        
        # Traiter le bonus numérique
        bonus_match = re.search(r'\+\s*(\d+)$', des)
        if bonus_match:
            bonus = int(bonus_match.group(1))
            total_result += bonus
            all_rolls.append(str(bonus))
        
        # Traiter le dé de karma
        if add_karma and karma_dice:
            karma_faces = int(karma_dice[1:])
            karma_result, karma_rolls = roll_single_die(karma_faces)
            total_result += karma_result
            all_rolls.extend([f"+karma: {roll}" for roll in karma_rolls])
        
        return total_result, " + ".join(all_rolls)

    def on_return(self, event):
        self.lancer_des()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = DiceRollerApp()
    app.run()