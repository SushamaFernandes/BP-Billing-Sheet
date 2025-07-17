import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os
import re
import numpy as np

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI Bill - Excel Processor")
        self.file_path = None
        self.df = None

        # Title label
        self.title_label = tk.Label(root, text="AI Bill - Excel Processor", font=("Arial", 16, "bold"))
        self.title_label.pack(pady=10)

        # Upload button
        self.upload_btn = tk.Button(root, text="Upload File", command=self.upload_file, width=20)
        self.upload_btn.pack(pady=5)

        # Process and Save button
        self.process_btn = tk.Button(root, text="Process and Save File", command=self.process_and_save, width=20, state=tk.DISABLED)
        self.process_btn.pack(pady=5)

        # Status label
        self.status_label = tk.Label(root, text="", fg="blue")
        self.status_label.pack(pady=10)

    def upload_file(self):
        filetypes = [("Excel files", "*.xlsx;*.xls")]
        path = filedialog.askopenfilename(title="Select Excel File", filetypes=filetypes)
        if path:
            try:
                # Use correct engine for .xls files
                if path.lower().endswith('.xls'):
                    self.df = pd.read_excel(path, engine='xlrd')
                else:
                    self.df = pd.read_excel(path)
                self.file_path = path
                self.status_label.config(text=f"Loaded: {os.path.basename(path)}", fg="green")
                self.process_btn.config(state=tk.NORMAL)
            except Exception as e:
                self.status_label.config(text=f"Error loading file: {e}", fg="red")
                self.process_btn.config(state=tk.DISABLED)

    def process_and_save(self):
        if self.df is None:
            self.status_label.config(text="No file loaded.", fg="red")
            return
        try:

            df = self.df.copy()
            # Filter out unwanted Project Name and Resource Name
            if "Project Name" in df.columns:
                df = df[~df["Project Name"].str.contains(r"Training|NBT", case=False, na=False)]
            df = df[df["Resource Name"] != "Mary Stella"]

            # Ensure required columns exist
            required_cols = ["Entry Date", "Resource Name", "Task Name", "Actul Work(hrs)"]
            for col in required_cols:
                if col not in df.columns:
                    self.status_label.config(text=f"Missing column: {col}", fg="red")
                    return

            # Format Entry Date
            df["Entry Date"] = pd.to_datetime(df["Entry Date"], errors='coerce').dt.strftime('%d-%m-%Y')


            # Extract Issue# using regex (case-insensitive, e.g., EKDM-40, DPROD-10081, DQUAl-2164)
            df["Issue#"] = df["Task Name"].str.extract(r'([A-Za-z]+-\d+)', flags=re.IGNORECASE, expand=False)
            # Replace missing Issue# with pd.NA for consistency
            df["Issue#"] = df["Issue#"].where(df["Issue#"].notna(), pd.NA)

            # Clean Task Name (remove Issue# if present, case-insensitive)
            def clean_task_name(task_name):
                # Remove Issue# (case-insensitive)
                cleaned = re.sub(r'([A-Za-z]+-\d+)', '', str(task_name), flags=re.IGNORECASE).strip()
                # Remove all variations of '--> :', '-->:', '--> :', '-->: ' with optional numbers, colons, and spaces before
                cleaned = re.sub(r'.*?\d*\s*-->(\s*:\s*|:\s*|\s*:|:|\s*)', '', cleaned)
                # Remove ':' and spaces around it (if any left)
                cleaned = re.sub(r'\s*:\s*', '', cleaned)
                return cleaned.strip()

            df["Task Name Clean"] = df["Task Name"].apply(clean_task_name)
            df["Task Name Out"] = np.where(df["Task Name Clean"] != "", df["Task Name Clean"], df["Task Name"])

            # Teams mapping using Resource Name
            resource_to_team = {
                "Punam Patil": "Dcab",
                "Paresh Damani": "Dcab",
                "Suyog Vasage": "Digistyle",
                "Dattatray Awaghade": "Digistyle",
                "Prashant Bhayekar": "Product",
                "Dhawalshri Jadhav": "Product",
                "Anuja Redekar": "Product",
                "Abrar Shaha": "Product",
                "Ruchita Shetye": "Product",
                "Royston Rodrigues": "Product",
                "Sharad Kodag": "Product",
                "Vipin Verma": "Product",
                "Ritesh Salian": "Product",
                "Nayan Kale": "Product",
                "Nishu Shah": "Product",
                "Hrishikesh Dadhe": "Product",
                "Ashwini Kanojia": "Product",
                "Dattatray Awaghade": "Product",
                "Inderjeet Jethwani": "Product",
                "Narayan Panigrahi": "Rebuying",
                "Sushama Fernandes": "Rebuying",
                "Rushikesh Shete": "Rebuying",
                "Dhiraj Pawar": "Rebuying",
                "Reshma Kute": "Rebuying",
                "Mohammed Azim Ansari": "Rebuying",
                "Sagar Padwal": "Rebuying",
                "Shlok Patil": "Rebuying",
                "Pranav Dasamane": "Rebuying",
                "Vishal Naik": "Rebuying",
                "Mohd Waseem Shaikh": "Rebuying",
                "Ashwini Kanojia": "Rebuying",
                "Jyotikaur Jassi": "Rebuying",
                "Soham Kale": "AI",
                "Arun Kumar": "AI",
                "Rishi Misra": "AI",
                "Hardik Raja": "AI"
            }

            def infer_team(task, resource):
                # Use mapping if available
                if resource in resource_to_team:
                    return resource_to_team[resource]
                # Previous logic fallback
                if "DJ" in str(task):
                    return "Product"
                return ""

            df["Teams"] = df.apply(lambda row: infer_team(row["Task Name"], row["Resource Name"]), axis=1)

            # Module inference using Issue# mapping
            issue_to_module = {
                "EKDM": "EKDM",
                "DPROD": "Product",
                "DSDEV": "Digistyle",
                "DCOM": "Commerce",
                "TAC": "TAC",
                "DART": "Artwork",
                "STYLEPOOL": "Stylepool",
                "LABELS": "Labelgenerator",
                "DQUAL": "d:pat",
                "DOC": "d:document",
                "DBUS": "BusinessPartner",
                "CPT": "CPT",
                "CAB": "A2DSS",
                "PRICE": "d:pricing",
                "DSDATA": "Commerce",
                "DSCM": "d:iwa",
                "DCAB": "Dcab",
                "DSEL": "d:select",
                "DCIS": "d:cision",
                "APL": "APL",
                "DQUAN": "d:quantity",
                "DINV": "d:invoice",
                "DROSI": "Dcab",
                "DMILE": "Milestonemaster",
                "DSIGN": "Product"
            }


            def infer_module(row):
                # Meeting/support logic
                task = str(row["Task Name"])
                issue = str(row["Issue#"])
                # If Task Name is 'Framework Call', set Module as 'Framework'
                if task.strip().lower() == "framework call":
                    return "Framework"
                if (
                    "meeting" in task.lower() or
                    "call" in task.lower() or
                    "sprint-scrum-on call support" in task.lower()
                ):
                    return np.nan  # Placeholder, will fill later
                # Use Issue# prefix for mapping
                if issue and issue != 'nan':
                    prefix = issue.split('-')[0].upper()
                    if prefix == "DSDATA":
                        # Keyword to module mapping for DSDATA
                        keyword_to_module = {
                            "commerce": "Commerce",
                            "iwa": "d:iwa",
                            "select": "d:select",
                            "cision": "d:cision",
                            "quantity": "d:quantity",
                            "invoice": "d:invoice",
                            "cab": "Dcab",
                            "milestone": "Milestonemaster",
                            "label": "Labelgenerator",
                            "pricing": "d:pricing",
                            "api": "API",
                            "product": "Product",
                            "artwork": "Artwork",
                            "digistyle": "Digistyle",
                            "tac": "TAC",
                            "ekdm": "EKDM",
                            "stylepool": "Stylepool",
                            "businesspartner": "BusinessPartner",
                            "document": "d:document",
                            "pat": "d:pat",
                            "a2dss": "A2DSS",
                            "cpt": "CPT",
                            "apl": "APL"
                        }
                        task_lower = task.lower()
                        for keyword, module in keyword_to_module.items():
                            if keyword in task_lower:
                                return module
                        return ""  # If no keyword matches
                    module = issue_to_module.get(prefix, "")
                    if module:
                        return module
                # Fallback logic (optional)
                if "API" in task:
                    return "API"
                return ""

            df["Module"] = df.apply(infer_module, axis=1)

            # Task Type inference based on member mapping
            analysis_members = {
                "Sushama Fernandes",
                "Prashant Bhayekar",
                "Dhawalshri Jadhav",
                "Punam Patil",
                "Dattatray Awaghade",
                "Nishu Shah",
                "Paresh Damani",
                "Sagar Padwal"
            }
            testing_members = {
                "Anuja Redekar",
                "Reshma Kute",
                "Jyotikaur Jassi"
            }

            def infer_task_type(resource):
                if resource in analysis_members:
                    return "Analysis"
                if resource in testing_members:
                    return "Testing"
                return "Development"

            df["Task Type"] = df["Resource Name"].apply(infer_task_type)

            # Mandays calculation
            df["Mandays"] = pd.to_numeric(df["Actul Work(hrs)"], errors='coerce') / 8
            df["Mandays"] = df["Mandays"].round(2)

            # Billable logic (updated as per new rule)
            def infer_billable(row):
                task_name_lower = str(row["Task Name"]).lower()
                # Always billable if Task Name is 'Sprint-Scrum-On Call Support'
                if task_name_lower.strip() == "sprint-scrum-on call support":
                    return "Yes"
                # If Issue# is not empty or Task Name contains 'Project Management', mark as Yes
                if (pd.notna(row["Issue#"]) and str(row["Issue#"]).strip() != "") or ("project management" in task_name_lower):
                    return "Yes"
                if row["Task Type"] == "Analysis" and row["Teams"] == "Product":
                    return "Yes"
                return "No"

            df["Billable"] = df.apply(infer_billable, axis=1)

            # Fill Module for meetings/support based on other modules for the same resource and date
            def fill_meeting_module(row, df):
                if pd.isna(row["Module"]) or row["Module"] == "":
                    same_day = df[
                        (df["Resource Name"] == row["Resource Name"]) &
                        (df["Entry Date"] == row["Entry Date"]) &
                        (df["Module"].notna()) & (df["Module"] != "")
                    ]
                    if not same_day.empty:
                        return same_day["Module"].iloc[0]
                return row["Module"]

            df["Module"] = df.apply(lambda row: fill_meeting_module(row, df), axis=1)
            df["Module"] = df["Module"].replace(np.nan, "", regex=True)

            # Prepare output DataFrame
            output_cols = [
                "Entry Date", "Resource Name", "Issue#", "Task Name",
                "Actul Work(hrs)", "Teams", "Module", "Task Type", "Mandays", "Billable"
            ]
            df["Task Name"] = df["Task Name Out"]
            df_out = df[output_cols]

            # Save dialog
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Processed File"
            )
            if save_path:
                df_out.to_excel(save_path, index=False)
                self.status_label.config(text=f"File saved: {os.path.basename(save_path)}", fg="green")
            else:
                self.status_label.config(text="Save cancelled.", fg="blue")

        except Exception as e:
            self.status_label.config(text=f"Error: {e}", fg="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()
