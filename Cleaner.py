import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# Lien vers le fichier Google Sheets contenant les groupes de spécialités
specialites_url = 'https://docs.google.com/spreadsheets/d/14Dzb6p-Y2nTdeNIKm5avvu2bHW8Iuj9PTrnz512TTZU/pub?gid=0&single=true&output=csv'
blacklist_url = 'https://docs.google.com/spreadsheets/d/1UJlFqLqJub0I9N_X_VbArX_QUlgnB-7cl95URrDzkEU/gviz/tq?tqx=out:csv&sheet=Blacklist'

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Nettoyage de CSV")
        self.geometry("400x250")

        # Initialisation des groupes de spécialités
        self.categories_pro = self.charger_specialites()

        # Configurer un grid avec des colonnes et lignes vides autour pour centrer
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(2, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # Boutons pour choisir fichier et lancer le nettoyage
        tk.Button(self, text="Choisir le fichier CSV", command=self.choisir_fichier, width=20).grid(row=1, column=1, pady=10, sticky="nsew")

        # Ajouter un menu déroulant pour les catégories professionnelles
        self.label_categorie = tk.Label(self, text="Choisissez une catégorie :")
        self.label_categorie.grid(row=2, column=1, pady=10, sticky="nsew")

        self.categorie_var = tk.StringVar()
        if self.categories_pro:
            self.categorie_var.set(list(self.categories_pro.keys())[0])  # Valeur par défaut
            self.menu_categorie = tk.OptionMenu(self, self.categorie_var, *self.categories_pro.keys())
            self.menu_categorie.config(width=20)  # Fixer la largeur du menu déroulant
            self.menu_categorie.grid(row=3, column=1, pady=10, sticky="nsew")
        else:
            self.categorie_var.set("Aucune catégorie trouvée")
            self.menu_categorie = tk.OptionMenu(self, self.categorie_var, "Aucune catégorie")
            self.menu_categorie.config(width=20)  # Fixer la largeur du menu déroulant
            self.menu_categorie.grid(row=3, column=1, pady=10, sticky="nsew")

        # Bouton pour nettoyer le fichier (désactivé au début)
        self.btn_nettoyer = tk.Button(self, text="Nettoyer le fichier", command=self.nettoyer_csv, state=tk.DISABLED, width=20)
        self.btn_nettoyer.grid(row=5, column=1, pady=20, sticky="nsew")

    def charger_specialites(self):
        """Charge les groupes de spécialités depuis le Google Sheets et retourne un dictionnaire."""
        try:
            df_specialites = pd.read_csv(specialites_url)

            # Nettoyer les données
            for col in df_specialites.columns:
                if df_specialites[col].dtype == 'object':
                    df_specialites[col] = df_specialites[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

            categories_pro = {col: df_specialites[col].dropna().tolist() for col in df_specialites.columns}
            return categories_pro
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de charger les spécialités depuis Google Sheets : {e}")
            return {}

    def choisir_fichier(self):
        self.fichier_csv = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.fichier_csv:
            self.btn_nettoyer.config(state=tk.NORMAL)
            messagebox.showinfo("Fichier sélectionné", "Fichier CSV sélectionné : " + self.fichier_csv)

    def nettoyer_csv(self):
        try:
            # Lire la blacklist
            df_blacklist = pd.read_csv(blacklist_url)
            df_blacklist = df_blacklist.applymap(lambda x: x.strip() if isinstance(x, str) else x)

            # Extraire les données de la blacklist
            blacklist_emails = df_blacklist['mail'].dropna().astype(str).tolist()
            blacklist_domains = df_blacklist['domaine'].dropna().astype(str).tolist()
            blacklist_tri_mail = df_blacklist['tri_mail'].dropna().astype(str).tolist()

            # Charger le fichier CSV et nettoyer les colonnes
            separateur = self.detecter_separateur(self.fichier_csv)
            df = pd.read_csv(self.fichier_csv, sep=separateur)
            df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

            # Colonnes à garder
            colonnes_necessaires = ['title', 'cid', 'address', 'categories/0', 'city', 'location/lat',
                                    'location/lng', 'phone', 'url', 'website', 'contactDetails/emails/0',
                                    'contactDetails/emails/1', 'contactDetails/emails/2']
            colonnes_necessaires = [col for col in colonnes_necessaires if col in df.columns]
            if not colonnes_necessaires:
                raise ValueError("Aucune colonne nécessaire trouvée dans le fichier CSV.")
            df = df[colonnes_necessaires].drop_duplicates()

            # Filtrer selon la catégorie sélectionnée
            categorie_choisie = self.categorie_var.get()
            categories_a_garder = self.categories_pro.get(categorie_choisie, [])
            df = df[df['categories/0'].isin(categories_a_garder)]

            # Filtrer les lignes en fonction des emails et domaines dans la blacklist
            def filter_row(row):
                # Vérifier les emails dans les trois colonnes
                for email_col in ['contactDetails/emails/0', 'contactDetails/emails/1', 'contactDetails/emails/2']:
                    if email_col in row and pd.notna(row[email_col]):
                        email = row[email_col]
                        domain = email.split('@')[-1] if '@' in email else ''

                        # 1. Vérifier si l'email exact est dans la colonne "mail" de la blacklist => suppression de la ligne
                        if email in blacklist_emails:
                            return None  # Si trouvé, la ligne est supprimée (None pour marquer la ligne à supprimer)

                        # 2. Vérifier si le domaine est dans la colonne "domaine" de la blacklist => suppression de la ligne
                        if any(bl_domain in domain for bl_domain in blacklist_domains):
                            return None  # Si trouvé, la ligne est supprimée

                        # 3. Vérifier si un morceau de "tri_mail" est dans l'email => suppression de la cellule (pas de la ligne)
                        if any(bl_item in email for bl_item in blacklist_tri_mail):
                            row[email_col] = None  # Efface uniquement l'email dans la cellule concernée

                return row

            # Appliquer le filtrage et supprimer les lignes marquées pour suppression
            df_nettoye = df.apply(filter_row, axis=1).dropna(how='any')

            # Sauvegarde du fichier nettoyé
            fichier_nettoye = self.sauvegarder_fichier(df_nettoye)
            messagebox.showinfo("Succès", f"Fichier nettoyé enregistré sous : {fichier_nettoye}")

        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur est survenue : {e}")

    def detecter_separateur(self, fichier):
        """Détecte le séparateur utilisé dans le fichier CSV."""
        with open(fichier, 'r', encoding='utf-8') as file:
            lignes = [file.readline() for _ in range(5)]
        compte_separateurs = {sep: sum(ligne.count(sep) for ligne in lignes) for sep in [',', ';', '\t']}
        return max(compte_separateurs, key=compte_separateurs.get)

    def sauvegarder_fichier(self, df):
        """Sauvegarde le fichier sur le bureau."""
        bureau = os.path.join(os.path.expanduser('~'), 'Desktop')
        base, ext = os.path.splitext(os.path.basename(self.fichier_csv))
        fichier_nettoye = os.path.join(bureau, f"{base}_CLEAN{ext}")

        # Ajouter suffixe en cas de conflit de nom
        suffixe = 1
        while os.path.exists(fichier_nettoye):
            fichier_nettoye = os.path.join(bureau, f"{base}_CLEAN_{suffixe}{ext}")
            suffixe += 1
        df.to_csv(fichier_nettoye, index=False)
        return fichier_nettoye

if __name__ == "__main__":
    app = Application()
    app.mainloop()
