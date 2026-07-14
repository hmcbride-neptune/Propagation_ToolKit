import os
import re
import time
import subprocess
import argparse
from pywinauto import Application, Desktop
import pyautogui

# ----------------------------
# CONFIG
# ----------------------------

GATEWAY_SCRIPT = r"C:\ProgramData\Python\Gateways\Gateway Locations.py"
EDX_PROJECTS_DIR = r"C:\ProgramData\EDX\Projects"

# ----------------------------
# HELPERS
# ----------------------------

def sanitize_filename_part(text):
    return re.sub(r'[<>:"/\\|?*]', "_", str(text).strip())


def get_edx_main_window():
    app = Application(backend="win32").connect(title_re=r".*\[initial view\].*")
    main = app.window(title_re=r".*\[initial view\].*")
    main.wait("visible enabled", timeout=5)
    return app, main


def confirm_edx_signalpro_open():
    _, main = get_edx_main_window()
    main.set_focus()


def get_project_name_from_edx():
    _, main = get_edx_main_window()
    title = main.window_text()

    match = re.search(
        r"64-bit\s*-\s*(.*?)\s*-\s*\[initial view\]",
        title,
        re.IGNORECASE
    )

    if not match:
        raise RuntimeError("Could not parse project name")

    return match.group(1).strip()


# ----------------------------
# DIALOG HANDLING (RETRY BASED)
# ----------------------------

def get_export_transmitters_dialog(retries=15, delay=0.15):
    for _ in range(retries):
        try:
            dlg = Desktop(backend="win32").window(
                title="Export transmitters",
                class_name="#32770"
            )
            if dlg.exists(timeout=0.5):
                return dlg
        except:
            pass

        time.sleep(delay)

    raise RuntimeError("Export transmitters dialog not found")


def open_export_transmitters_dialog():
    _, main = get_edx_main_window()

    main.set_focus()
    main.type_keys("%f")
    time.sleep(0.05)
    main.type_keys("e")

    return get_export_transmitters_dialog()


# ----------------------------
# POPUP HANDLER
# ----------------------------

def acknowledge_export_popup():
    try:
        
        
        # popup = Desktop(backend="win32").window(title="Information")
        # popup.wait("visible", timeout=10)

        # popup.type_keys("{ENTER}")
        # popup.wait_not("visible", timeout=5)

        # time.sleep(0.25)  # allow UI rebuild
        time.sleep(0.05)
        pyautogui.press("enter")
        return True
    except:
        return False


# ----------------------------
# GROUP HANDLING (FAST INDEX)
# ----------------------------

def get_group_data(dlg):
    combo = dlg["Transmitter group:ComboBox"]
    groups = [g.strip() for g in combo.item_texts() if g.strip()]
    index_map = {g: i for i, g in enumerate(groups)}
    return groups, index_map


def select_transmitter_group(group_name, index_map):
    dlg = get_export_transmitters_dialog()
    combo = dlg["Transmitter group:ComboBox"]

    combo.click_input()
    time.sleep(0.05)

    combo.type_keys("{HOME}")

    for _ in range(index_map[group_name]):
        combo.type_keys("{DOWN}")

    combo.type_keys("{TAB}")

    return dlg


# ----------------------------
# EXPORT
# ----------------------------


def set_export_file_name(dlg, file_name):
    # ✅ Wait until edit control is ready (NOT just dialog)
    edit = None
    for _ in range(10):
        try:
            edit = dlg["Export file:Edit"]
            if edit.exists(timeout=0.3):
                break
        except:
            pass
        time.sleep(0.1)

    if not edit:
        raise RuntimeError("Export file Edit control not found")

    current_path = edit.window_text()
    folder = os.path.dirname(current_path)

    new_path = os.path.join(folder, file_name)

    edit.set_edit_text(new_path)

    return new_path

def export_current_group_to_csv(dlg, filename):
    path = set_export_file_name(dlg, filename)

    print("Exporting:", path)

    dlg["Export to CSV"].click_input()
    time.sleep(1)
    acknowledge_export_popup()

    dlg = get_export_transmitters_dialog()

    return dlg, path


# ----------------------------
# MAIN LOGIC
# ----------------------------

def export_transmitters(project_name):

    confirm_edx_signalpro_open()

    dlg = open_export_transmitters_dialog()

    groups, index_map = get_group_data(dlg)

    print("Groups:", groups)

    exported = []

    master = "Master group" if "Master group" in groups else groups[0]

    print("Master:", master)

    dlg = select_transmitter_group(master, index_map)
    dlg, path = export_current_group_to_csv(dlg, "GateWay Locations - Master.csv")
    exported.append(path)

    for g in groups:
        if g in {"", "Master group"} or g == master:
            continue

        print("Selecting:", g)

        dlg = select_transmitter_group(g, index_map)

        file = f"GateWay Locations - {sanitize_filename_part(g)}.csv"

        dlg, path = export_current_group_to_csv(dlg, file)

        exported.append(path)

    dlg["OK"].click_input()

    subprocess.Popen(
        [
            "python",
            GATEWAY_SCRIPT,
            os.path.join(
                EDX_PROJECTS_DIR,
                project_name,
                "Data",
                "xPropInfo",
                "GateWay Locations - Master.csv"
            )
        ]
    )

    return exported


# ----------------------------
# ENTRY POINT
# ----------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("project_name", nargs="?", default=None)
    args = parser.parse_args()

    project = args.project_name or get_project_name_from_edx()

    export_transmitters(project)


if __name__ == "__main__":
    main()