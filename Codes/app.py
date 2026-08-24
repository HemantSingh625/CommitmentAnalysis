import os
import sys
import json
import traceback
import threading
import shutil
import tempfile
import pandas as pd
import openpyxl
import openpyxl.utils.dataframe
import openpyxl.cell.cell
import re
import calendar
import difflib
import warnings
import plotly
import plotly.express
import plotly.graph_objects
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Ensures all dependencies are bundled by PyInstaller
pd.options.mode.chained_assignment = None
warnings.filterwarnings('ignore')

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Tata Steel - Compliance AI")
        self.geometry("650x550")
        
        self.prod_file = ctk.StringVar(value="No file selected")
        self.comm_files = []
        self.comm_text = ctk.StringVar(value="0 files selected")
        self.profile = ctk.StringVar(value="742 CRCA")
        
        # Title
        self.lbl_title = ctk.CTkLabel(self, text="Compliance Report Generator", font=ctk.CTkFont(size=24, weight="bold"))
        self.lbl_title.pack(pady=(20, 10))
        
        # Step 1: Production File
        self.frame_prod = ctk.CTkFrame(self)
        self.frame_prod.pack(pady=10, padx=20, fill="x")
        
        self.lbl_prod_desc = ctk.CTkLabel(self.frame_prod, text="1. Select Production Data File:")
        self.lbl_prod_desc.pack(anchor="w", padx=10, pady=(10, 5))
        
        self.btn_prod = ctk.CTkButton(self.frame_prod, text="Browse Production", command=self.browse_prod)
        self.btn_prod.pack(side="left", padx=10, pady=(0, 10))
        
        self.lbl_prod = ctk.CTkLabel(self.frame_prod, textvariable=self.prod_file, text_color="gray", width=300, anchor="w")
        self.lbl_prod.pack(side="left", padx=10, pady=(0, 10))
        
        # Step 2: Commitment Files
        self.frame_comm = ctk.CTkFrame(self)
        self.frame_comm.pack(pady=10, padx=20, fill="x")
        
        self.lbl_comm_desc = ctk.CTkLabel(self.frame_comm, text="2. Select Monthly Commitment Files (Select multiple):")
        self.lbl_comm_desc.pack(anchor="w", padx=10, pady=(10, 5))
        
        self.btn_comm = ctk.CTkButton(self.frame_comm, text="Browse Commitments", command=self.browse_comm)
        self.btn_comm.pack(side="left", padx=10, pady=(0, 10))
        
        self.lbl_comm = ctk.CTkLabel(self.frame_comm, textvariable=self.comm_text, text_color="gray", width=300, anchor="w")
        self.lbl_comm.pack(side="left", padx=10, pady=(0, 10))
        
        # Step 3: Profile Selection
        self.frame_profile = ctk.CTkFrame(self)
        self.frame_profile.pack(pady=10, padx=20, fill="x")
        
        self.lbl_profile_desc = ctk.CTkLabel(self.frame_profile, text="3. Select the Report Profile:")
        self.lbl_profile_desc.pack(anchor="w", padx=10, pady=(10, 5))
        
        self.rad_742 = ctk.CTkRadioButton(self.frame_profile, text="742 CRCA", variable=self.profile, value="742 CRCA")
        self.rad_742.pack(side="left", padx=20, pady=(0, 10))
        
        self.rad_743 = ctk.CTkRadioButton(self.frame_profile, text="743 & 744 PPGL", variable=self.profile, value="743 & 744 PPGL")
        self.rad_743.pack(side="left", padx=20, pady=(0, 10))
        
        # Run Button
        self.btn_run = ctk.CTkButton(self, text="Generate Compliance Report", font=ctk.CTkFont(size=16, weight="bold"), height=40, command=self.start_generation)
        self.btn_run.pack(pady=20)
        
        # Status Label
        self.lbl_status = ctk.CTkLabel(self, text="", text_color="green")
        self.lbl_status.pack()

    def browse_prod(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file:
            self.prod_file.set(file)

    def browse_comm(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if files:
            self.comm_files = list(files)
            self.comm_text.set(f"{len(self.comm_files)} files selected")

    def start_generation(self):
        if self.prod_file.get() == "No file selected":
            messagebox.showerror("Error", "Please select the Production file.")
            return
        if len(self.comm_files) == 0:
            messagebox.showerror("Error", "Please select at least one Commitment file.")
            return
            
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Master Report As"
        )
        if not save_path:
            return
            
        self.lbl_status.configure(text="Generating report... Please wait (this may take a minute).", text_color="orange")
        self.btn_run.configure(state="disabled")
        
        # Run in thread to not freeze UI
        threading.Thread(target=self.run_pipeline, args=(self.prod_file.get(), self.comm_files, self.profile.get(), save_path)).start()

    def run_pipeline(self, prod_path, comm_paths, profile, save_path):
        try:
            # Create a temporary workspace manually to avoid strict WinError 32 on cleanup
            temp_dir = tempfile.mkdtemp()
            try:
                # Copy production file
                prod_filename = "742.XLSX" if profile == "742 CRCA" else "743&744.XLSX"
                temp_prod_path = os.path.join(temp_dir, prod_filename)
                shutil.copy2(prod_path, temp_prod_path)
                
                # Copy commitment files
                comm_dir = os.path.join(temp_dir, "monthly_commitments")
                os.makedirs(comm_dir)
                for c_path in comm_paths:
                    shutil.copy2(c_path, os.path.join(comm_dir, os.path.basename(c_path)))
                
                # Store original working dir and switch to temp
                original_cwd = os.getcwd()
                os.chdir(temp_dir)
                
                try:
                    # Write config.json
                    config = {}
                    if profile == "742 CRCA":
                        config = {
                            "profile_name": profile,
                            "production_file": prod_filename,
                            "plant": ["742"],
                            "plant_codes": ["742"],
                            "product_category": "CRCA",
                            "p_codes": ["B47", "C01", "B01", "C06", "B48", "B49", "B50", "C24", "C41"]
                        }
                    else:
                        config = {
                            "profile_name": profile,
                            "production_file": prod_filename,
                            "plant": ["743", "744", 743, 744],
                            "plant_codes": ["743", "744"],
                            "product_category": "PPGL,GL,GI,PPGI",
                            "p_codes": ["B04", "C05", "B03", "C02", "B18", "B20", "B06", "B11", "B17", "B14", "C04", "B22"]
                        }
                        
                    with open("config.json", "w") as f:
                        json.dump(config, f, indent=4)
                        
                    scripts = ["clean_production.py", "clean_commitments.py", "merge_master.py", "generate_dashboard.py"]
                    for script in scripts:
                        script_path = resource_path(script)
                        if not os.path.exists(script_path):
                            # Try looking in original directory as fallback (for testing)
                            script_path = os.path.join(original_cwd, script)
                            if not os.path.exists(script_path):
                                raise Exception(f"Missing internal script: {script}")
                        with open(script_path, "r", encoding="utf-8-sig") as s:
                            code = s.read()
                            # Execute in fresh namespace
                            exec(code, {"__name__": "__main__"})
                            
                    if not os.path.exists("Master_Compliance_Report.xlsx"):
                        raise Exception("Pipeline finished but output file was not created.")
                        
                    # Move to final destination
                    shutil.move("Master_Compliance_Report.xlsx", save_path)
                    
                    # Move dashboard if created
                    if os.path.exists("Compliance_Dashboard.html"):
                        html_path = save_path.replace(".xlsx", "_Dashboard.html")
                        shutil.move("Compliance_Dashboard.html", html_path)
                    
                finally:
                    os.chdir(original_cwd)
            finally:
                import gc
                import time
                gc.collect()
                time.sleep(0.5)
                shutil.rmtree(temp_dir, ignore_errors=True)
                
            self.after(0, self.finish_success, save_path)
        except Exception as e:
            err = traceback.format_exc()
            self.after(0, self.finish_error, str(e) + "\n" + err)

    def finish_success(self, filename):
        self.lbl_status.configure(text=f"Success! Saved as:\n{os.path.basename(filename)}", text_color="green")
        self.btn_run.configure(state="normal")
        messagebox.showinfo("Complete", f"Compliance Report generated successfully!\n\nSaved at:\n{filename}")

    def finish_error(self, err_msg):
        self.lbl_status.configure(text="Error occurred. See popup.", text_color="red")
        self.btn_run.configure(state="normal")
        messagebox.showerror("Error", f"Failed to generate report:\n\n{err_msg}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
