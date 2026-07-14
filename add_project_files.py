import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
import time
from datetime import datetime


def copy_folder_contents(src_folder, dest_folder):
    """
    Mimics AutoIt DirCopy behavior by copying the contents of src_folder
    into dest_folder.
    """
    os.makedirs(dest_folder, exist_ok=True)

    for item in os.listdir(src_folder):
        src_item = os.path.join(src_folder, item)
        dest_item = os.path.join(dest_folder, item)

        if os.path.isdir(src_item):
            shutil.copytree(src_item, dest_item, dirs_exist_ok=True)
        else:
            shutil.copy2(src_item, dest_item)

def add_proj_files(proj_path, proj_type):
    
    time.sleep(0.1)

    folder_name = os.path.basename(proj_path)
  
    template_reports_path = r"\\TetonRiver\prop\template\Data\xReports"
    dest = fr"{proj_path}\Data\xReports"
    copy_folder_contents(template_reports_path, dest)

    today = datetime.now().strftime("%y%m%d")
    base_file_name = f"{folder_name} {today} {proj_type}".strip()

    services_file_name = f"{base_file_name} Services-Area Count.xlsx"

    services_dest = os.path.join(dest, services_file_name)

    services_template = r"\\TetonRiver\prop\template\Data\xReports\Project Rxxx YYMMDD - Services-Area Count.xlsx"

    # handle duplicate
    if os.path.exists(services_dest):
        services_dest = os.path.join(dest, f"(1){services_file_name}")

    shutil.copy2(services_template, services_dest)

    os.startfile(dest)

# --- GUI CODE ---
def browse_folder():
    folder = filedialog.askdirectory(initialdir=r"C:\ProgramData\EDX\Projects",title="Select Project Folder")
    if folder:
        path_var.set(folder)

def create_project():
    proj_path = path_var.get()
    proj_type = dropdown_var.get()

    if not proj_path:
        messagebox.showerror("Error", "Please select a folder.")
        return

    try:
        add_proj_files(proj_path, proj_type)
        messagebox.showinfo("Success", "Project files created successfully.")
        root.destroy()
    except Exception as e:
        messagebox.showerror("Error", str(e))

# --- MAIN WINDOW ---
root = tk.Tk()
root.title("Project File Creator")
root.geometry("500x200")

# Always on top (optional, based on your previous requirement)
root.attributes("-topmost", True)

# Path selection
tk.Label(root, text="Project Folder:").pack(pady=(10, 0))

path_var = tk.StringVar()

frame_path = tk.Frame(root)
frame_path.pack(padx=10, fill="x")

tk.Entry(frame_path, textvariable=path_var).pack(side="left", fill="x", expand=True)
tk.Button(frame_path, text="Browse", command=browse_folder, width=10).pack(side="right", padx=5)

# Dropdown
tk.Label(root, text="Project Type:").pack(pady=(10, 0))

dropdown_var = tk.StringVar()
dropdown_var.set("")  # default blank

dropdown = tk.OptionMenu(root, dropdown_var, "", "Revision", "As-Built")
dropdown.pack()

# Create button
tk.Button(root, text="Create", command=create_project, width=20).pack(pady=20)


root.mainloop()
