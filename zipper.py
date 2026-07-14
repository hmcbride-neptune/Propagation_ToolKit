import os
import zipfile
import shutil
import datetime
import tkinter as tk
from tkinter import messagebox, ttk

# ----- CONFIG -----
SOURCE_DIR = r"C:\ProgramData\EDX\Projects"
DEST_BASE = r"\\tetonriver\prop\Propagations"
CURRENT_YEAR = str(datetime.datetime.now().year)
DEST_DIR = os.path.join(DEST_BASE, CURRENT_YEAR)


# ----- CLEAN FUNCTION -----
def clean_project_folder(folder_path):
    """
    Deletes everything except:
    - 'Data' subfolder
    - any .xml files in root
    """
    for item in os.listdir(folder_path):
        full_path = os.path.join(folder_path, item)

        # Keep Data folder
        if item.lower() == "data" and os.path.isdir(full_path):
            continue

        # Keep XML files (root only)
        if item.lower().endswith(".xml") and os.path.isfile(full_path):
            continue

        # Delete everything else
        if os.path.isdir(full_path):
            shutil.rmtree(full_path)
        else:
            os.remove(full_path)


# ----- ZIP FUNCTION -----
def zip_folder(folder_path, dest_dir):
    folder_name = os.path.basename(folder_path)
    zip_path = os.path.join(dest_dir, f"{folder_name}.zip")

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root_dir, dirs, files in os.walk(folder_path):
            for file in files:
                full_path = os.path.join(root_dir, file)
                rel_path = os.path.relpath(full_path, folder_path)
                zipf.write(full_path, rel_path)

    return zip_path


# ----- MAIN PROCESS -----
def process_selection():
    selected_indices = listbox.curselection()

    if not selected_indices:
        messagebox.showwarning("No Selection", "Please select at least one project.")
        return

    # Confirm destructive action
    confirm = messagebox.askyesno(
        "Confirm",
        "This will CLEAN and DELETE the selected project folders after zipping.\n\nContinue?"
    )
    if not confirm:
        return

    os.makedirs(DEST_DIR, exist_ok=True)

    selected_folders = [listbox.get(i) for i in selected_indices]

    for folder in selected_folders:
        full_path = os.path.join(SOURCE_DIR, folder)

        try:
            # Step 1: Clean
            clean_project_folder(full_path)

            # Step 2: Zip
            zip_path = zip_folder(full_path, DEST_DIR)

            # Step 3: Delete original ONLY if zip succeeded
            if os.path.exists(zip_path):
                shutil.rmtree(full_path)

        except Exception as e:
            messagebox.showerror("Error", f"Failed processing {folder}\n{e}")
            return

    messagebox.showinfo(
        "Complete",
        f"Processed and removed {len(selected_folders)} project(s)\nSaved to:\n{DEST_DIR}"
    )

    root.destroy()


# ----- GUI -----
root = tk.Tk()
root.title("EDX Project Archive Tool")
root.geometry("300x400")
root.attributes("-topmost", True)

frame = ttk.Frame(root, padding=10)
frame.pack(fill="both", expand=True)

label = ttk.Label(
    frame,
    text="Select projects to ZIP",
    font=("Segoe UI", 10)
)
label.pack(anchor="w", pady=(0, 10))

listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE)
listbox.pack(fill="both", expand=True, pady=(0, 10))

scrollbar = ttk.Scrollbar(listbox, orient="vertical", command=listbox.yview)
listbox.config(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")

# Populate folders
if os.path.exists(SOURCE_DIR):
    folders = sorted([
        f for f in os.listdir(SOURCE_DIR)
        if os.path.isdir(os.path.join(SOURCE_DIR, f))
    ])
    for folder in folders:
        listbox.insert(tk.END, folder)
else:
    messagebox.showerror("Error", f"Path not found:\n{SOURCE_DIR}")

# Buttons
button_frame = ttk.Frame(frame)
button_frame.pack(fill="x")

select_btn = ttk.Button(
    button_frame,
    text="Select",
    command=process_selection
)
select_btn.pack(side="right", padx=5)

cancel_btn = ttk.Button(
    button_frame,
    text="Cancel",
    command=root.destroy
)
cancel_btn.pack(side="right")

root.mainloop()
