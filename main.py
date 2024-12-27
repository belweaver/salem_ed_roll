import tkinter as tk
from ui import create_ui

def main():
    print("Initialisation de l'application...")  # Ajout du message de débogage
    # Initialiser l’application
    root = tk.Tk()
    root.title('Lancer de dés')
    root.geometry("600x600")

    create_ui(root)
    root.mainloop()

if __name__ == "__main__":
    main()