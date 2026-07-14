import os
import shutil
import re
import glob
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox


class new_request:
    """Encapsulate the 'New Project' dialog in a class.

    Instantiate with a `tk` root-like object (Tk or Toplevel). If `root`
    is omitted, a `tk.Tk()` instance is created and exposed as
    `self.root` so the caller can invoke `self.root.mainloop()`.
    """

    def __init__(self, root=None):
        self.root = root if root is not None else tk.Tk()
        self._own_root = root is None
        # Will be set to the created project folder name when creation succeeds
        self.project_name = None

        self.root.title("New Project")
        self.root.geometry("720x110")
        self.root.resizable(False, False)
        
        # Build the UI
        header = tk.Frame(self.root, bg="#7F78B8", height=30)
        header.pack(fill="x")

        title_label = tk.Label(
            header,
            text="  New Project",
            bg="#7F78B8",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        )
        title_label.pack(fill="both", padx=5)

        main = tk.Frame(self.root, bg="#F0F0F0")
        main.pack(fill="both", expand=True, padx=10, pady=5)

        labels = ["SiteID", "Utility", "City", "State", "SysType"]
        for i, text in enumerate(labels):
            lbl = tk.Label(main, text=text, font=("Segoe UI", 10), bg="#F0F0F0")
            lbl.grid(row=0, column=i, padx=10, sticky="w")

        # Inputs stored as attributes so handlers can access them
        self.siteid_entry = ttk.Entry(main, width=18)
        self.siteid_entry.grid(row=1, column=0, padx=10, pady=5)

        self.utility_entry = ttk.Entry(main, width=18)
        self.utility_entry.grid(row=1, column=1, padx=10, pady=5)

        self.city_entry = ttk.Entry(main, width=18)
        self.city_entry.insert(0, "(Optional)")
        self.city_entry.grid(row=1, column=2, padx=10, pady=5)

        self.state_var = tk.StringVar(value="AL")
        states_provinces = [
            # USA
            "AL",
            "AK",
            "AZ",
            "AR",
            "CA",
            "CO",
            "CT",
            "DE",
            "FL",
            "GA",
            "HI",
            "ID",
            "IL",
            "IN",
            "IA",
            "KS",
            "KY",
            "LA",
            "ME",
            "MD",
            "MA",
            "MI",
            "MN",
            "MS",
            "MO",
            "MT",
            "NE",
            "NV",
            "NH",
            "NJ",
            "NM",
            "NY",
            "NC",
            "ND",
            "OH",
            "OK",
            "OR",
            "PA",
            "RI",
            "SC",
            "SD",
            "TN",
            "TX",
            "UT",
            "VT",
            "VA",
            "WA",
            "WV",
            "WI",
            "WY",
            # Canada Provinces
            "AB",
            "BC",
            "MB",
            "NB",
            "NL",
            "NS",
            "ON",
            "PE",
            "QC",
            "SK",
            # Canada Territories
            "NT",
            "NU",
            "YT",
        ]

        state_dropdown = ttk.Combobox(
            main, textvariable=self.state_var, values=states_provinces, width=8, state="normal"
        )
        state_dropdown.grid(row=1, column=3, padx=10, pady=5)

        self.sys_var = tk.StringVar(value="R900")
        sys_dropdown = ttk.Combobox(
            main, textvariable=self.sys_var, values=["R900", "R450", "Cellular"], width=10, state="readonly"
        )
        sys_dropdown.grid(row=1, column=4, padx=10, pady=5)

        create_btn = ttk.Button(main, text="Create", command=self.create_project)
        create_btn.grid(row=1, column=5, padx=15)

    # ----------------------------
    # Helper methods
    # ----------------------------
    @staticmethod
    def _clean_text(value: str) -> str:
        value = value.strip()
        return re.sub(r'[\\/:*?"<>|]', '', value)

    @staticmethod
    def _copy_without_overwrite(src: str, dst: str) -> None:
        for root_dir, dirs, files in os.walk(src):
            relative_path = os.path.relpath(root_dir, src)
            target_dir = os.path.join(dst, relative_path)
            os.makedirs(target_dir, exist_ok=True)
            for file in files:
                src_file = os.path.join(root_dir, file)
                dst_file = os.path.join(target_dir, file)
                if not os.path.exists(dst_file):
                    shutil.copy2(src_file, dst_file)

    def create_project(self):
        # --- Get values from GUI ---
        siteid = self._clean_text(self.siteid_entry.get())
        utility = self._clean_text(self.utility_entry.get())
        city = self._clean_text(self.city_entry.get())
        state = self._clean_text(self.state_var.get())
        systype = self._clean_text(self.sys_var.get())

        # --- Build folder name ---
        if city and city.lower() != "(optional)":
            folder_name = f"{siteid} {utility} {city} {state} {systype}"
        else:
            folder_name = f"{siteid} {utility} {state} {systype}"

        # --- Paths ---
        base_dir = r"\\tetonriver\prop\Requests"
        template_dir = r"\\tetonriver\prop\Template\Data"
        new_folder_path = os.path.join(base_dir, folder_name)

        try:
            # 1) Create project folder
            os.makedirs(new_folder_path, exist_ok=True)

            # 2) Copy the entire Data folder into the new project folder
            destination_data_folder = os.path.join(new_folder_path, "Data")
            self._copy_without_overwrite(template_dir, destination_data_folder)

            # 3) Make renamed COPIES of specific files in Data\xReports
            reports_dir = os.path.join(destination_data_folder, "xReports")

            if not os.path.isdir(reports_dir):
                raise FileNotFoundError(f"xReports folder not found:\n{reports_dir}")

            yymmdd = datetime.now().strftime("%y%m%d")

            ppt_matches = glob.glob(os.path.join(reports_dir, "Project * New Prop Study.pptx"))
            xlsx_matches = glob.glob(os.path.join(reports_dir, "Project * - Services-Area Count.xlsx"))

            # --- PPT copy ---
            if ppt_matches:
                src_ppt = ppt_matches[0]
                new_ppt_name = f"{folder_name} {yymmdd} New Prop Study.pptx"
                dst_ppt = os.path.join(reports_dir, new_ppt_name)
                shutil.copy2(src_ppt, dst_ppt)
            else:
                raise FileNotFoundError(
                    f"Could not find PowerPoint template in:\n{reports_dir}\n"
                    f"Expected something like: Project R900 YYMMDD New Prop Study.pptx"
                )

            # --- Excel copy ---
            if xlsx_matches:
                src_xlsx = xlsx_matches[0]
                new_xlsx_name = f"{folder_name} {yymmdd} - Services-Area Count.xlsx"
                dst_xlsx = os.path.join(reports_dir, new_xlsx_name)
                shutil.copy2(src_xlsx, dst_xlsx)
            else:
                raise FileNotFoundError(
                    f"Could not find Excel template in:\n{reports_dir}\n"
                    f"Expected something like: Project Rxxx YYMMDD - Services-Area Count.xlsx"
                )

            # 4) Open xPropInfo folder
            propinfo_path = os.path.join(destination_data_folder, "xPropInfo")

            if os.path.isdir(propinfo_path):
                subprocess.Popen(f'explorer "{propinfo_path}"')
            else:
                messagebox.showwarning("Warning", f"xPropInfo folder not found:\n{propinfo_path}")

            # 5) Save project name for caller and close program
            self.project_name = folder_name
            self.root.destroy()

        except FileExistsError:
            messagebox.showerror("Error", "Project folder already exists.")
        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    app = new_request()
    app.root.mainloop()
