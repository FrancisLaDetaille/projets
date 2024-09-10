import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class CSVSorterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CSV File Sorter")
        self.geometry("400x200")

        # Label pour l'instruction
        self.label = tk.Label(self, text="Sélectionnez un fichier CSV", pady=10)
        self.label.pack()

        # Bouton pour sélectionner un fichier CSV
        self.btn_select_file = tk.Button(self, text="Choisir un fichier CSV", command=self.select_file)
        self.btn_select_file.pack(pady=10)

        # Bouton pour lancer le tri
        self.btn_sort_file = tk.Button(self, text="Lancer le tri", command=self.process_csv, state=tk.DISABLED)
        self.btn_sort_file.pack(pady=10)

        # Placeholder pour le chemin du fichier sélectionné
        self.file_path = ""

    def select_file(self):
        """Ouvre un dialogue pour sélectionner un fichier CSV."""
        self.file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.file_path:
            self.btn_sort_file.config(state=tk.NORMAL)
            messagebox.showinfo("Fichier Sélectionné", f"Fichier CSV sélectionné : {self.file_path}")

    def process_csv(self):
        """Traite le fichier CSV et le divise en différents fichiers en fonction des critères."""
        try:
            # Charger le fichier CSV
            df = pd.read_csv(self.file_path)

            # Vérifier que les colonnes nécessaires sont présentes
            required_columns = ['reviewsCount', 'totalScore', 'contactDetails/emails/0']
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("Erreur", "Le fichier CSV doit contenir les colonnes suivantes : 'reviewsCount', 'totalScore', 'contactDetails/emails/0'.")
                return

            # Convertir les colonnes 'reviewsCount' et 'totalScore' en numérique
            df['reviewsCount'] = pd.to_numeric(df['reviewsCount'], errors='coerce')
            df['totalScore'] = pd.to_numeric(df['totalScore'], errors='coerce')

            # Filtrage des fichiers
            # 1. Plus de 300 avis et note > 4 avec email non nul
            df_filtered_1 = df[(df['reviewsCount'] > 300) & (df['totalScore'] > 4) & (df['contactDetails/emails/0'].notna())]

            # 2. Autres (email non nul)
            df_filtered_2 = df[(~((df['reviewsCount'] > 300) & (df['totalScore'] > 4))) & (df['contactDetails/emails/0'].notna())]

            # 3. Sans email
            df_filtered_3 = df[df['contactDetails/emails/0'].isna()]

            # Enregistrer les fichiers sur le bureau
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

            # Fichier 1 : Plus de 300 avis et note > 4 avec email
            output_file_1 = os.path.join(desktop_path, "LisaCare mail.csv")
            df_filtered_1.to_csv(output_file_1, index=False)

            # Fichier 2 : Autres avec email
            output_file_2 = os.path.join(desktop_path, "Mails pour Trust.csv")
            df_filtered_2.to_csv(output_file_2, index=False)

            # Fichier 3 : Sans email
            output_file_3 = os.path.join(desktop_path, "Tel.csv")
            df_filtered_3.to_csv(output_file_3, index=False)

            # Message de succès
            messagebox.showinfo("Succès", "Les fichiers ont été triés et enregistrés sur le bureau.")

        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur est survenue : {e}")

if __name__ == "__main__":
    app = CSVSorterApp()
    app.mainloop()
