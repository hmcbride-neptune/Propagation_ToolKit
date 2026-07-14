import argparse
import pyautogui
from pywinauto import Application, Desktop, mouse
import time
from edx_open import EDX_OPEN
import tkinter as tk
from tkinter import messagebox
import keyboard
import pyperclip
import os

def snapshot_texts(dlg):
    """Walk the UI tree once and cache each element's texts.

    dlg.descendants() and element.texts() are expensive cross-process
    calls, so we do a single traversal and reuse the result for every
    label lookup instead of re-walking the tree per label.
    """
    return [item.texts() for item in dlg.descendants()]

def get_value_after_label(all_texts, label_text):
    label_lower = label_text.lower()

    for i, texts in enumerate(all_texts):
        for t in texts:
            if label_lower in t.lower():
                # Check neighbor ABOVE first (EDX layout)
                if i - 1 >= 0:
                    prev = all_texts[i - 1]
                    if prev and "mi^2" in prev[0]:
                        return prev[0]

                # Fallback: check below
                if i + 1 < len(all_texts):
                    nxt = all_texts[i + 1]
                    if nxt and "mi^2" in nxt[0]:
                        return nxt[0]

    raise Exception(f"{label_text} value not found")

def run_demographics_query(project):
    pyautogui.hotkey('f10')
    pyautogui.press(['q', 'down', 'enter'], interval=0.2)
    dlg = Desktop(backend="win32").window(title_re=".*Query Studies.*")
    dlg.wait('exists visible', timeout=10)

    # Run Query
    dlg.child_window(auto_id="_btnRun").click_input()

    # Get Coverage Area and Total Area

    # Poll for results instead of a fixed sleep: return as soon as the
    # query has populated the "mi^2" values, capped at a timeout.
    coverage_text = total_text = None
    deadline = time.time() + 10
    while True:
        all_texts = snapshot_texts(dlg)  # single tree walk, reused below
        try:
            coverage_text = get_value_after_label(all_texts, "Signal Area")
            total_text = get_value_after_label(all_texts, "Total Area")
            break
        except Exception:
            if time.time() >= deadline:
                raise
            time.sleep(0.25)

    coverage_area = float(coverage_text.replace(" mi^2", "").strip())
    total_area = float(total_text.replace(" mi^2", "").strip())
    save_csv_file(dlg, project)
    dlg.child_window(auto_id="_btnClose").click_input()
    return coverage_area, total_area

def get_demographics(project):
    app = Application(backend="uia").connect(title_re=".*EDX.*")
    dlg = app.top_window()
    dlg.set_focus()
    coverage_area, total_area = run_demographics_query(project)
    return coverage_area, total_area 

def show_demographics_results(coverage_area, total_area):
    paste_count = 0
    hotkey_id = None

    root = tk.Tk()
    root.withdraw()

    popup = tk.Toplevel(root)
    popup.title("Demographics Results")
    popup.attributes("-topmost", True)
    popup.resizable(False, False)

    message = f"({coverage_area}/{total_area}) mi\u00B2"

    label = tk.Label(
        popup,
        text=message,
        font=("Segoe UI", 12),
        padx=25,
        pady=15
    )
    label.pack()

    status_label = tk.Label(
        popup,
        text=f"Clipboard: {coverage_area}",
        font=("Segoe UI", 9),
        fg="gray"
    )
    status_label.pack(pady=(0, 10))

    def copy_to_clipboard(value):
        root.clipboard_clear()
        root.clipboard_append(str(value))
        root.update()

    def close_popup():
        nonlocal hotkey_id

        try:
            if hotkey_id is not None:
                keyboard.remove_hotkey(hotkey_id)
        except Exception:
            pass

        popup.destroy()
        root.destroy()

    def on_ctrl_v():
        nonlocal paste_count

        paste_count += 1

        if paste_count == 1:
            # Let the first paste use coverage_area, then switch clipboard to total_area
            root.after(150, lambda: copy_second_value())

        elif paste_count == 2:
            # Let the second paste use total_area, then close popup
            root.after(150, close_popup)

    def copy_second_value():
        copy_to_clipboard(total_area)
        status_label.config(text=f"Clipboard: {total_area}")

    popup.protocol("WM_DELETE_WINDOW", close_popup)

    # Copy coverage_area first
    copy_to_clipboard(coverage_area)

    # Detect Ctrl+V globally
    hotkey_id = keyboard.add_hotkey("ctrl+v", on_ctrl_v)

    # Center-ish popup
    popup.update_idletasks()
    width = popup.winfo_width()
    height = popup.winfo_height()
    x = popup.winfo_screenwidth() // 2 - width // 2
    y = popup.winfo_screenheight() // 2 - height // 2
    popup.geometry(f"+{x}+{y}")

    root.mainloop()

def get_unique_filename(folder, base_name):
    name, ext = os.path.splitext(base_name)
    candidate = base_name
    i = 1

    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{name} ({i}){ext}"
        i += 1

    return candidate

def save_csv_file(dlg, project):
    
    base_path = r"C:\ProgramData\EDX\Projects"
    project_path = os.path.join(base_path, project)
    save_dir = os.path.join(project_path, "Data", "PQ")

    base_name = "Query_studies.txt"
    unique_name = get_unique_filename(save_dir, base_name)

    save_path = os.path.join(save_dir, unique_name)

    # Copy full path to clipboard
    pyperclip.copy(save_path)

    # Click save button
    dlg.child_window(auto_id="_btnSaveText").click_input()
    time.sleep(0.5)

    # Paste path and save
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.1)
    pyautogui.press('enter')

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("project_name", nargs="?", default=None)
    args = parser.parse_args()
    try:
        edx = EDX_OPEN()
        project = args.project_name or edx.get_project_name_from_edx()
        edx.confirm_edx_signalpro_open()
        
    except:
        messagebox.showerror("", "Demographics Study Failed")
        exit()

    coverage_area, total_area = get_demographics(project)
    show_demographics_results(coverage_area, total_area)
    

if __name__ == "__main__":

    main()
