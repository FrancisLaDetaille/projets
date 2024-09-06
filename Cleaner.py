import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# URL for Google Sheets with specialty groups
specialites_url = 'https://docs.google.com/spreadsheets/d/14Dzb6p-Y2nTdeNIKm5avvu2bHW8Iuj9PTrnz512TTZU/pub?gid=0&single=true&output=csv'
blacklist_url = 'https://docs.google.com/spreadsheets/d/1UJlFqLqJub0I9N_X_VbArX_QUlgnB-7cl95URrDzkEU/gviz/tq?tqx=out:csv&sheet=Blacklist'

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CSV Cleaner")
        self.geometry("400x250")

        # Load specialty groups
        self.categories_pro = self.load_specialties()

        # Configure grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(2, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # Button to select file
        tk.Button(self, text="Choose CSV File", command=self.choose_file, width=20).grid(row=1, column=1, pady=10, sticky="nsew")

        # Dropdown for professional categories
        self.label_categorie = tk.Label(self, text="Choose a category:")
        self.label_categorie.grid(row=2, column=1, pady=10, sticky="nsew")

        self.categorie_var = tk.StringVar()
        if self.categories_pro:
            self.categorie_var.set(list(self.categories_pro.keys())[0])
            self.menu_categorie = tk.OptionMenu(self, self.categorie_var, *self.categories_pro.keys())
            self.menu_categorie.config(width=20)
            self.menu_categorie.grid(row=3, column=1, pady=10, sticky="nsew")
        else:
            self.categorie_var.set("No category found")
            self.menu_categorie = tk.OptionMenu(self, self.categorie_var, "No category")
            self.menu_categorie.config(width=20)
            self.menu_categorie.grid(row=3, column=1, pady=10, sticky="nsew")

        # Button to clean file (initially disabled)
        self.btn_nettoyer = tk.Button(self, text="Clean File", command=self.clean_csv, state=tk.DISABLED, width=20)
        self.btn_nettoyer.grid(row=5, column=1, pady=20, sticky="nsew")

    def load_specialties(self):
        """Load specialty groups from Google Sheets."""
        try:
            df_specialites = pd.read_csv(specialites_url)

            # Clean data
            for col in df_specialites.columns:
                if df_specialites[col].dtype == 'object':
                    df_specialites[col] = df_specialites[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

            categories_pro = {col: df_specialites[col].dropna().tolist() for col in df_specialites.columns}
            return categories_pro
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load specialties: {e}")
            return {}

    def choose_file(self):
        """Open file dialog to choose a CSV file."""
        self.fichier_csv = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.fichier_csv:
            self.btn_nettoyer.config(state=tk.NORMAL)
            messagebox.showinfo("File Selected", "Selected CSV file: " + self.fichier_csv)

    def clean_csv(self):
        """Clean the CSV file."""
        try:
            # Load blacklist
            df_blacklist = pd.read_csv(blacklist_url)

            # Clean blacklist data
            df_blacklist = df_blacklist.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))

            # Extract blacklist data
            blacklist_emails = df_blacklist['mail'].dropna().astype(str).tolist()
            blacklist_domains = df_blacklist['domaine'].dropna().astype(str).tolist()
            blacklist_tri_mail = df_blacklist['tri_mail'].dropna().astype(str).tolist()

            # Load and clean CSV file
            separator = self.detect_separator(self.fichier_csv)
            df = pd.read_csv(self.fichier_csv, sep=separator)
            df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))

            # Columns to keep
            necessary_columns = ['title', 'cid', 'address', 'categories/0', 'city', 'location/lat',
                                 'location/lng', 'phone', 'url', 'website', 'contactDetails/emails/0',
                                 'contactDetails/emails/1', 'contactDetails/emails/2']
            necessary_columns = [col for col in necessary_columns if col in df.columns]
            if not necessary_columns:
                raise ValueError("No necessary columns found in CSV file.")
            df = df[necessary_columns].drop_duplicates()

            # Filter by selected category
            chosen_category = self.categorie_var.get()
            categories_to_keep = self.categories_pro.get(chosen_category, [])
            df = df[df['categories/0'].isin(categories_to_keep)]

            # 1. Filter out blacklisted emails and domains
            def filter_rows(row):
                for email_col in ['contactDetails/emails/0', 'contactDetails/emails/1', 'contactDetails/emails/2']:
                    if email_col in row and pd.notna(row[email_col]):
                        email = row[email_col]
                        domain = email.split('@')[-1]

                        if email in blacklist_emails or any(bl_domain in domain for bl_domain in blacklist_domains):
                            return False

                return True

            df_filtered = df[df.apply(filter_rows, axis=1)]

            # 2. Replace emails containing parts of tri_mail
            def replace_tri_mail(row):
                for email_col in ['contactDetails/emails/0', 'contactDetails/emails/1', 'contactDetails/emails/2']:
                    if email_col in row and pd.notna(row[email_col]):
                        email = row[email_col]
                        if any(bl_tri in email for bl_tri in blacklist_tri_mail):
                            row[email_col] = None

                return row

            df_cleaned = df_filtered.apply(replace_tri_mail, axis=1)

            # Save cleaned file
            cleaned_file = self.save_file(df_cleaned)
            messagebox.showinfo("Success", f"Cleaned file saved as: {cleaned_file}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def detect_separator(self, file):
        """Detect the CSV separator."""
        with open(file, 'r', encoding='utf-8') as f:
            lines = [f.readline() for _ in range(5)]
        sep_count = {sep: sum(line.count(sep) for line in lines) for sep in [',', ';', '\t']}
        return max(sep_count, key=sep_count.get)

    def save_file(self, df):
        """Save cleaned file to desktop."""
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        base, ext = os.path.splitext(os.path.basename(self.fichier_csv))
        cleaned_file = os.path.join(desktop, f"{base}_CLEAN{ext}")

        # Add suffix if name conflict
        suffix = 1
        while os.path.exists(cleaned_file):
            cleaned_file = os.path.join(desktop, f"{base}_CLEAN_{suffix}{ext}")
            suffix += 1
        df.to_csv(cleaned_file, index=False)
        return cleaned_file

if __name__ == "__main__":
    app = Application()
    app.mainloop()
