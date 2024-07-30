import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# URL du fichier blacklist
blacklist_url = 'https://docs.google.com/spreadsheets/d/1UJlFqLqJub0I9N_X_VbArX_QUlgnB-7cl95URrDzkEU/gviz/tq?tqx=out:csv&sheet=Blacklist'

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Nettoyage de CSV")
        self.geometry("400x200")

        # Bouton pour choisir le fichier CSV
        self.btn_choisir_fichier = tk.Button(self, text="Choisir le fichier CSV", command=self.choisir_fichier)
        self.btn_choisir_fichier.pack(pady=20)

        # Bouton pour lancer le nettoyage
        self.btn_nettoyer = tk.Button(self, text="Nettoyer le fichier", command=self.nettoyer_csv, state=tk.DISABLED)
        self.btn_nettoyer.pack(pady=20)

        self.fichier_csv = None

    def choisir_fichier(self):
        """Ouvre une boîte de dialogue pour choisir un fichier CSV."""
        self.fichier_csv = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.fichier_csv:
            self.btn_nettoyer.config(state=tk.NORMAL)
            messagebox.showinfo("Fichier sélectionné", "Fichier CSV sélectionné : " + self.fichier_csv)
        else:
            messagebox.showwarning("Aucun fichier", "Aucun fichier sélectionné.")

    def nettoyer_csv(self):
        """Nettoie le fichier CSV des informations et filtre les personnes blacklistées."""
        if not self.fichier_csv:
            messagebox.showerror("Erreur", "Aucun fichier CSV sélectionné.")
            return

        try:
            # Lire les données du fichier blacklist depuis Google Sheets
            df_blacklist = pd.read_csv(blacklist_url)
            df_blacklist.columns = df_blacklist.columns.str.strip()

            # Détecter le séparateur utilisé dans le fichier CSV
            separateur = self.detecter_separateur(self.fichier_csv)

            # Charger le fichier CSV avec les informations
            df = self.lire_csv(self.fichier_csv, separateur)
            df.columns = df.columns.str.strip()

            # Liste des colonnes nécessaires
            colonnes_necessaires = [
                'title',
                'cid',
                'address',
                'categories/0',
                'city',
                'location/lat',
                'location/lng',
                'phone',
                'url',
                'website',
                'mail'  # Assurez-vous que cette colonne existe dans le fichier à nettoyer
            ]

            # Vérifier et conserver uniquement les colonnes nécessaires qui existent dans le DataFrame
            colonnes_existantes = [col for col in colonnes_necessaires if col in df.columns]
            df = df[colonnes_existantes]

            # Supprimer les doublons basés sur les colonnes nécessaires
            df = df.drop_duplicates()

            # Filtrer les lignes en fonction de la blacklist
            def match_blacklist(row):
                """Vérifie si la ligne correspond à une valeur de la blacklist dans n'importe quelle colonne."""
                for col in df.columns:
                    if col in df_blacklist.columns:
                        blacklist_values = df_blacklist[col].dropna().astype(str)
                        if not blacklist_values.empty:
                            for value in blacklist_values:
                                try:
                                    # Assurer que la valeur à comparer est une chaîne
                                    if pd.notna(row[col]) and str(row[col]).strip() in blacklist_values.values:
                                        return True
                                except Exception as e:
                                    print(f"Erreur lors de la comparaison pour {value}: {e}")
                return False

            df_nettoye = df[~df.apply(match_blacklist, axis=1)]

            # Exclure les adresses e-mail avec des domaines spécifiques
            if 'mail' in df.columns and 'domaine' in df_blacklist.columns:
                domaines_exclus = df_blacklist['domaine'].dropna().unique()
                if len(domaines_exclus) > 0:
                    # Assurer que la colonne mail est en chaînes de caractères
                    df_nettoye = df_nettoye.copy()  # Créer une copie pour éviter SettingWithCopyWarning
                    # Remplacer les NaN par une chaîne vide avant la conversion
                    df_nettoye['mail'] = df_nettoye['mail'].fillna('').astype(str)
                    # Exclusion basée sur les domaines spécifiés
                    df_nettoye = df_nettoye[~df_nettoye['mail'].apply(
                        lambda email: any(email.lower().endswith(f'@{domaine.lower()}') for domaine in domaines_exclus)
                    )]

            # Exclure les adresses e-mail contenant "chu" après "@"
            if 'mail' in df.columns:
                df_nettoye = df_nettoye[~df_nettoye['mail'].str.contains(r'@.*chu', case=False, na=False)]

            # Obtenir le chemin du bureau
            bureau = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') if os.name == 'nt' else os.path.join(os.path.expanduser('~'), 'Desktop')

            # Préparer le nom du fichier nettoyé avec suffixe si nécessaire
            base, extension = os.path.splitext(os.path.basename(self.fichier_csv))
            fichier_nettoye = os.path.join(bureau, f"{base}_CLEAN{extension}")

            # Ajouter un suffixe pour éviter les conflits de noms
            suffixe = 1
            while os.path.exists(fichier_nettoye):
                fichier_nettoye = os.path.join(bureau, f"{base}_CLEAN_{suffixe}{extension}")
                suffixe += 1

            # Sauvegarder le fichier nettoyé sur le bureau
            df_nettoye.to_csv(fichier_nettoye, index=False)

            messagebox.showinfo("Succès", f"Fichier nettoyé enregistré sous : {fichier_nettoye}")

        except ValueError as e:
            messagebox.showerror("Erreur", f"Erreur : {e}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur est survenue : {e}")

    def detecter_separateur(self, fichier):
        """Détecte le séparateur utilisé dans le fichier CSV."""
        with open(fichier, 'r', encoding='utf-8') as file:
            lignes = [file.readline() for _ in range(5)]  # Lire les 5 premières lignes

        # Compter les occurrences de différents séparateurs
        separateurs = [',', ';', '\t']
        compte_separateurs = {sep: sum(ligne.count(sep) for ligne in lignes) for sep in separateurs}

        # Choisir le séparateur le plus fréquent
        separateur = max(compte_separateurs, key=compte_separateurs.get)

        print(f"Séparateur détecté : '{separateur}'")
        return separateur

    def lire_csv(self, fichier, sep):
        """Lit le fichier CSV avec le séparateur donné."""
        try:
            df = pd.read_csv(fichier, sep=sep)
            return df
        except pd.errors.ParserError:
            raise ValueError(f"Erreur lors de la lecture du fichier CSV avec le séparateur '{sep}'.")

if __name__ == "__main__":
    app = Application()
    app.mainloop()
