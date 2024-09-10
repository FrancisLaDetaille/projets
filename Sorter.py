import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class CSVSorterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CSV File Sorter")
        self.geometry("400x200")

        # Label for instructions
        self.label = tk.Label(self, text="Select a CSV file", pady=10)
        self.label.pack()

        # Button to select CSV file
        self.btn_select_file = tk.Button(self, text="Choose CSV File", command=self.select_file)
        self.btn_select_file.pack(pady=10)

        # Button to start sorting (disabled initially)
        self.btn_sort_file = tk.Button(self, text="Start Sorting", command=self.process_csv, state=tk.DISABLED)
        self.btn_sort_file.pack(pady=10)

        # Placeholder for file path
        self.file_path = ""

    def select_file(self):
        """Open dialog to select CSV file."""
        self.file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.file_path:
            self.btn_sort_file.config(state=tk.NORMAL)
            messagebox.showinfo("File Selected", f"CSV file selected: {self.file_path}")

    def process_csv(self):
        """Process and split the CSV based on criteria."""
        try:
            # Load CSV file
            df = pd.read_csv(self.file_path)

            # Check required columns
            required_columns = ['reviewsCount', 'totalScore', 'contactDetails/emails/0']
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("Error", "CSV must have columns: 'reviewsCount', 'totalScore', 'contactDetails/emails/0'.")
                return

            # Convert 'reviewsCount' and 'totalScore' to numeric
            df['reviewsCount'] = pd.to_numeric(df['reviewsCount'], errors='coerce')
            df['totalScore'] = pd.to_numeric(df['totalScore'], errors='coerce')

            # Filter files
            # 1. More than 300 reviews and score > 4 with email
            df_filtered_1 = df[(df['reviewsCount'] > 300) & (df['totalScore'] > 4) & (df['contactDetails/emails/0'].notna())]

            # 2. Others with email
            df_filtered_2 = df[(~((df['reviewsCount'] > 300) & (df['totalScore'] > 4))) & (df['contactDetails/emails/0'].notna())]

            # 3. Without email
            df_filtered_3 = df[df['contactDetails/emails/0'].isna()]

            # Save files to desktop
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

            # File 1: More than 300 reviews and score > 4 with email
            output_file_1 = os.path.join(desktop_path, "LisaCare mail.csv")
            df_filtered_1.to_csv(output_file_1, index=False)

            # File 2: Others with email
            output_file_2 = os.path.join(desktop_path, "Mails pour Trust.csv")
            df_filtered_2.to_csv(output_file_2, index=False)

            # File 3: Without email
            output_file_3 = os.path.join(desktop_path, "Tel.csv")
            df_filtered_3.to_csv(output_file_3, index=False)

            # Success message
            messagebox.showinfo("Success", "Files sorted and saved to desktop.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    app = CSVSorterApp()
    app.mainloop()
