import os
import re
import time
import tkinter as tk
from tkinter import messagebox
from pywinauto import Desktop
from pywinauto.keyboard import send_keys


class EDXAutomation:
    def __init__(self):
        self.desktop = Desktop(backend="win32")

    # ------------------------------------------------------------
    # Main function equivalent to SwitchDailyHourly()
    # ------------------------------------------------------------
    def switch_daily_hourly(self):
        try:
            initial_win = self.activate_initial_view()
            title_str = initial_win.window_text()

            project_name = self.extract_project_name_from_title(title_str)

            if not project_name:
                messagebox.showerror(
                    "Error",
                    f"Could not determine project name from active title:\n\n{title_str}"
                )
                return

            project_xml = os.path.join(
                r"C:\ProgramData\EDX\Projects",
                project_name,
                "project.map.xml"
            )

            radio_template = self.get_radio_template_from_project_xml(project_xml)

            if "R450" in radio_template:
                selected_threshold = self.select_threshold_gui(
                    title="Select R450 coverage threshold.",
                    options=[
                        "Daily (-95)",
                        "Hourly (-90)",
                        "Billing (-100)"
                    ]
                )

                if selected_threshold:
                    threshold_value = self.get_r450_threshold_value(selected_threshold)
                    self.apply_area_study_threshold(threshold_value)

            elif "R900" in radio_template:
                selected_threshold = self.select_threshold_gui(
                    title="Select R900 coverage threshold.",
                    options=[
                        "Daily V4 (-100)",
                        "Hourly V4 (-90)",
                        "Billing V4 (-105)",
                        "BillingHG V4 (-110)",
                        "Daily V5 (-110)",
                        "Hourly V5 (-100)",
                        "Billing V5 (-115)"
                    ]
                )

                if selected_threshold:
                    threshold_value = self.get_r900_threshold_value(selected_threshold)
                    self.apply_area_study_threshold(threshold_value)

            else:
                messagebox.showerror(
                    "Error",
                    "Either project is not R450 or R900, or something is wrong with the project's project.map.xml file."
                )

        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ------------------------------------------------------------
    # Find and activate the EDX initial view window
    # ------------------------------------------------------------
    def activate_initial_view(self):
        win = self.desktop.window(title_re=r".*-\s*\[initial view\].*")
        win.set_focus()
        return win

    # ------------------------------------------------------------
    # AutoIt title parsing equivalent
    #
    # AutoIt logic:
    # Something like:
    # ProjectName - [initial view]
    # or:
    # EDX - ProjectName - [initial view]
    # ------------------------------------------------------------
    def extract_project_name_from_title(self, title_str):
        marker = " - [initial view]"

        if marker not in title_str:
            return None

        left_part = title_str.split(marker)[0]

        # Match the AutoIt behavior:
        # it takes text after the first " - "
        if " - " in left_part:
            return left_part.split(" - ", 1)[1].strip()

        return left_part.strip()

    # ------------------------------------------------------------
    # Read project.map.xml and determine R450/R900
    # ------------------------------------------------------------
    def get_radio_template_from_project_xml(self, xml_path):
        if not os.path.exists(xml_path):
            raise FileNotFoundError(f"Could not find project.map.xml:\n{xml_path}")

        with open(xml_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()

        if len(lines) < 60:
            raise ValueError("project.map.xml does not have at least 60 lines.")

        # AutoIt reads line 60.
        template_line = lines[59].strip()

        # Original AutoIt looked for text around a_R and "/>
        # This keeps it simple and robust.
        if "R450" in template_line:
            return "R450"

        if "R900" in template_line:
            return "R900"

        # fallback: try regex in the whole file in case line 60 shifts
        full_text = "".join(lines)

        match = re.search(r"a_(R450|R900)", full_text, re.IGNORECASE)
        if match:
            return match.group(1).upper()

        return template_line

    # ------------------------------------------------------------
    # tkinter replacement for AutoIt GUICreate/List/Button
    # ------------------------------------------------------------
    def select_threshold_gui(self, title, options):
        selected_value = {"value": None}

        root = tk.Tk()
        root.title(title)
        root.geometry("300x220")
        root.resizable(False, False)
        root.attributes("-topmost", True)

        listbox = tk.Listbox(root, font=("MS Sans Serif", 10), height=8)
        listbox.pack(padx=8, pady=8, fill="both", expand=True)

        for option in options:
            listbox.insert(tk.END, option)

        if options:
            listbox.selection_set(0)

        def select_item():
            selection = listbox.curselection()

            if not selection:
                messagebox.showwarning("Selection Required", "Please select a threshold.")
                return

            selected_value["value"] = listbox.get(selection[0])
            root.destroy()

        def on_double_click(event):
            select_item()

        select_button = tk.Button(root, text="Select", width=14, command=select_item)
        select_button.pack(pady=(0, 8), anchor="w", padx=8)

        listbox.bind("<Double-Button-1>", on_double_click)

        root.mainloop()

        return selected_value["value"]

    # ------------------------------------------------------------
    # R450 threshold mapping
    # ------------------------------------------------------------
    def get_r450_threshold_value(self, selected_text):
        mapping = {
            "Daily (-95)": "-95",
            "Hourly (-90)": "-90",
            "Billing (-100)": "-100"
        }

        if selected_text not in mapping:
            raise ValueError(f"Unknown R450 threshold selection: {selected_text}")

        return mapping[selected_text]

    # ------------------------------------------------------------
    # R900 threshold mapping
    # ------------------------------------------------------------
    def get_r900_threshold_value(self, selected_text):
        mapping = {
            "Daily V4 (-100)": "-100",
            "Hourly V4 (-90)": "-90",
            "Billing V4 (-105)": "-105",
            "BillingHG V4 (-110)": "-110",
            "Daily V5 (-110)": "-110",
            "Hourly V5 (-100)": "-100",
            "Billing V5 (-115)": "-115"
        }

        if selected_text not in mapping:
            raise ValueError(f"Unknown R900 threshold selection: {selected_text}")

        return mapping[selected_text]

    # ------------------------------------------------------------
    # Equivalent automation for both R450 and R900
    # ------------------------------------------------------------
    def apply_area_study_threshold(self, threshold_value):
        self.activate_initial_view()

        # AutoIt:
        # send("{AltDown}")
        # send("s")
        # send("a")
        # send("{AltUp}")
        #
        # Python equivalent:
        send_keys("%s")
        send_keys("a")

        area_studies = self.desktop.window(title="Area Studies for Map View: initial view")
        area_studies.set_focus()

        # AutoIt:
        # ControlClick("Area Studies for Map View: initial view","","[CLASS:SysListView32; INSTANCE:1]","left",2, 426, 25)
        #
        # In pywinauto/win32, found_index is zero-based.
        listview = area_studies.child_window(class_name="SysListView32", found_index=0)

        # Double-click around same position as AutoIt.
        listview.click_input(button="left", coords=(426, 25), double=True)
        time.sleep(0.2)

        details = self.desktop.window(title="Area Study Details")
        details.set_focus()
        self.set_area_study_details_threshold(details, threshold_value)

    # ------------------------------------------------------------
    # Equivalent to ControlSetText and the style/color level clicks
    # ------------------------------------------------------------
    def set_area_study_details_threshold(self, details, threshold_value):
        # AutoIt:
        # ControlSetText("Area Study Details","","[CLASS:Edit; INSTANCE:1]", threshold)
        edit = details.child_window(class_name="Edit", found_index=0)
        edit.set_edit_text(threshold_value)

        # AutoIt:
        # ControlClick("Area Study Details","Study Display Style","[CLASS:Button; INSTANCE:2]")
        self.click_button_by_instance(details, instance=2)

        style = self.desktop.window(title="Area Study Style")
        style.set_focus()

        # AutoIt:
        # ControlClick("Area Study Style","Set signal/color &levels...","[CLASS:Button; INSTANCE:2]")
        self.click_button_by_instance(style, instance=2)

        colors = self.desktop.window(title="Colors and levels")
        colors.set_focus()

        # AutoIt:
        # send("{Tab}")
        # send("{Enter}")
        # sleep(100)
        # send("{Tab 4}")
        # send(threshold)
        #
        # This keeps the same keyboard flow.
        # send_keys("{TAB}")
        # send_keys("{ENTER}")
        # time.sleep(0.1)
        send_keys("{TAB 4}")
        send_keys(threshold_value)
        time.sleep(0.1)
        send_keys("{ENTER}")
        send_keys("{ESC}")
        send_keys("{TAB}")
        send_keys("{ENTER}")
        send_keys("{ENTER}")


    # ------------------------------------------------------------
    # Helper for AutoIt [CLASS:Button; INSTANCE:n]
    #
    # AutoIt INSTANCE is 1-based.
    # pywinauto found_index is 0-based.
    # ------------------------------------------------------------
    def click_button_by_instance(self, window, instance):
        button = window.child_window(class_name="Button", found_index=instance - 1)
        button.click_input()


# ------------------------------------------------------------
# Standalone run
# ------------------------------------------------------------
if __name__ == "__main__":
    edx = EDXAutomation()
    edx.switch_daily_hourly()