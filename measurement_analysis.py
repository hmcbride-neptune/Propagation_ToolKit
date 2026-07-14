import re
import time
import tkinter as tk
from tkinter import messagebox
import keyboard
import pyperclip
from edx_open import EDX_OPEN

from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError

def open_measurement_analysis(win):
    """
    Uses the menu path shown in EDX.
    If this fails in your build, use print_control_identifiers() once to inspect.
    """
    try:
        win.menu_select("Studies->Measurement Analysis...")
    except Exception as e:
        raise

def wait_for_measurement_dialog(app):
    try:
        dlg = app.window(title_re=r"^Measurement Analysis$")
        dlg.wait("visible enabled ready", timeout=15)
        return dlg
    except ElementNotFoundError:
        raise RuntimeError("Measurement Analysis dialog not found")
    except Exception:
        raise RuntimeError("Measurement Analysis dialog not found")

def click_perform_area_analysis(dlg):
    try:
        # dlg.print_control_identifiers()
        dlg.child_window(title="&Perform Area Analysis", class_name="Button").click()
        time.sleep(4)  # allow results to refresh
    except Exception as e:
        raise

def get_dialog_texts(dlg):
    """
    Collect all visible texts in the Measurement Analysis dialog.
    """
    texts = []
    try:
        for c in dlg.descendants():
            try:
                txt = c.window_text()
                if txt:
                    texts.append(txt.strip())
            except Exception:
                pass
    except Exception:
        pass
    return [t for t in texts if t]

def parse_no_of_points_from_dialog_texts(texts):
    """
    The dialog has at least one 'No. of points' label.
    We try to find the first numeric value that follows the first occurrence.
    """
    for i, txt in enumerate(texts):
        if txt == "No. of points":
            for nxt in texts[i + 1:i + 6]:
                cleaned = nxt.replace(",", "").strip()
                if re.fullmatch(r"\d+", cleaned):
                    return int(cleaned)
    return None

def run_measurement_analysis_and_get_metrics(app, win):

    open_measurement_analysis(win)
    dlg = wait_for_measurement_dialog(app)

    click_perform_area_analysis(dlg)

    # 1) Read the visible dialog results
    texts = get_dialog_texts(dlg)
    for i, txt in enumerate(texts):
        if txt.strip() == "No. of points:":
            if i + 1 < len(texts):
                no_of_points =  int(texts[i + 1].strip().replace(",",""))

    file_path = dlg.child_window(class_name="Edit").window_text().strip()
    with open(file_path, "r") as f:
        total_points = sum(1 for _ in f) - 1

     # 2) Close dialog so coverage text (if any) is easier to read on main window
    try:
        dlg.child_window(title="&OK", class_name="Button").click()
        time.sleep(1)
    except Exception:
        pass

    return no_of_points, total_points

def show_and_copy(no_of_points, total_points):
    percentage = round(no_of_points / total_points * 100, 2)
    message = f"{percentage}% ({no_of_points}/{total_points})\n\n(Press Ctrl+v or Right-Click to continue)"

    pyperclip.copy(str(no_of_points))

    root = tk.Tk()
    root.title("Measurement Analysis Result")

    # Label
    label = tk.Label(root, text=message, font=("Segoe UI", 16), padx=20, pady=20)
    label.pack()

    # Keep window on top (good UX)
    root.attributes("-topmost", True)
    root.bind("Destroy>", lambda e: keyboard.unhook_all_hotkeys())
    keyboard.add_hotkey('ctrl+v', lambda: root.after(150, root.destroy))
    root.bind("<Button-3>", lambda e: root.destroy())

    root.mainloop()

if __name__ == "__main__":
    try:
        edx = EDX_OPEN()
        edx.confirm_edx_signalpro_open()
        app, win = edx.get_edx_main_window()
        no_of_points, total_points = run_measurement_analysis_and_get_metrics(app, win)  
        show_and_copy(no_of_points, total_points)
    except Exception as e:
        print(f"Automation failed: {e}")
