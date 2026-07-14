import os
import sys
import time
import pyautogui
import tkinter as tk
from tkinter import ttk, messagebox
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
import threading


class EDXAutomation:
    def __init__(self, group_combo, file_name):
        self.group_combo = group_combo
        self.file_name = file_name

        self.groups = []
        self.groups_gathered = False
        self.first_populate = False
        self.all_buttons_pressed = ""

    # ---------------------------------------------------------
    # AutoIt: PopulateGroups()
    # ---------------------------------------------------------
    def populate_groups(self):
        """
        Reads project.std.xml and populates the group combo box with:
            Master Group
            TxGroup names from the XML
        """

        # Clear combo box
        self.group_combo.set("")
        self.group_combo["values"] = ()

        # Reset group list
        self.groups = []

        self.first_populate = False

        xml_path = os.path.join(
            r"C:\ProgramData\EDX\Projects",
            self.file_name,
            "project.std.xml"
        )

        if not os.path.exists(xml_path):
            messagebox.showerror(
                "Missing XML",
                f"Could not find project.std.xml:\n\n{xml_path}"
            )
            return False

        try:
            with open(xml_path, "r", encoding="utf-8", errors="ignore") as f:
                xml_lines = f.readlines()

        except Exception as e:
            messagebox.showerror(
                "XML Read Error",
                f"Could not read project.std.xml:\n\n{e}"
            )
            return False

        group_size = 0

        # Find TxGroups size
        for line in xml_lines:
            if 'TxGroups size="' in line:
                try:
                    parts = line.split('"')
                    group_size = int(parts[1])
                except Exception:
                    group_size = 0
                break

        # Find TxGroup names
        for line in xml_lines:
            if 'TxGroup name="' in line:
                try:
                    parts = line.split('"')
                    group_name = parts[1]
                    self.groups.append(group_name)
                except Exception:
                    pass

            if group_size and len(self.groups) >= group_size:
                break

        if len(self.groups) > 1:
            self.groups_gathered = True

        # Populate combo box
        combo_values = self.groups
        self.group_combo["values"] = combo_values
        self.group_combo.set("Master group")

        return True

    # ---------------------------------------------------------
    # Helper functions
    # ---------------------------------------------------------
    def wait_for_window(self, title, timeout=15):
        """
        Waits for a window to exist and become active.
        """
        end_time = time.time() + timeout

        while time.time() < end_time:
            try:
                win = Desktop(backend="win32").window(title_re=title)
                if win.exists(timeout=1):
                    win.set_focus()
                    return win
            except Exception:
                pass

            time.sleep(0.25)

        raise TimeoutError(f"Window not found or not active: {title}")

    def get_window_rect(self, title):
        """
        Returns pywinauto rectangle for a window.
        """
        win = Desktop(backend="win32").window(title_re=title)
        win.wait("exists visible ready", timeout=10)
        return win.rectangle()

    def click_control_by_class_instance(self, window_title, control_type="Button", instance=1):
        """
        Rough replacement for AutoIt's:
        ControlClick("Window", "", "[CLASS:Button; INSTANCE:x]")

        pywinauto UIA does not use INSTANCE the same way AutoIt does,
        so this grabs all matching controls and clicks the requested index.
        """
        win = Desktop(backend="win32").window(title_re=window_title)
        win.wait("exists visible ready", timeout=10)

        controls = win.descendants(class_name=control_type)

        index = instance - 1

        if index < 0 or index >= len(controls):
            raise IndexError(
                f"Could not find {control_type} INSTANCE:{instance} "
                f"in window '{window_title}'. Found {len(controls)} controls."
            )

        controls[index].click_input()

    def get_static_text_by_instance(self, window_title, instance=3):
        """
        Replacement for:
        ControlGetText("Map Layers", "", "[CLASS:Static; INSTANCE:3]")
        """
        win = Desktop(backend="win32").window(title_re=window_title)
        win.wait("exists visible ready", timeout=10)

        statics = win.descendants(class_name="Static")

        index = instance - 1

        if index < 0 or index >= len(statics):
            return ""

        return statics[index].window_text()

    def pixel_search(self, left, top, right, bottom, target_rgb):
        """
        Basic replacement for AutoIt PixelSearch.

        target_rgb should be a tuple:
            (R, G, B)

        AutoIt color 0x0078D7 means:
            R = 0
            G = 120
            B = 215
        """
        screenshot = pyautogui.screenshot(region=(left, top, right - left, bottom - top))
        width, height = screenshot.size

        for y in range(height):
            for x in range(width):
                pixel = screenshot.getpixel((x, y))[:3]

                if pixel == target_rgb:
                    return left + x, top + y

        return None

    # ---------------------------------------------------------
    # AutoIt: SwitchGroups()
    # ---------------------------------------------------------
    def switch_groups(self):
        """
        Switches EDX display/calculation group based on selected combo value.
        """

        # Track action
        self.all_buttons_pressed += ",Switched group"

        selected_group = self.group_combo.get().strip()

        if selected_group == "":
            messagebox.showwarning("", "You didn't select a group!")
            return False

        # Determine group number using AutoIt-style logic
        #
        # AutoIt behavior:
        #   $number starts at 0
        #   If input matches $group[$number], then $number += 1
        #
        # So Python group_number is 1-based for actual TxGroups.
        group_number = None

        if selected_group.lower() == "master group":
            group_number = 1
        else:
            for idx, group_name in enumerate(self.groups):
                if selected_group == group_name:
                    group_number = idx + 1
                    break

        if group_number is None:
            messagebox.showerror(
                "Group Not Found",
                f"Could not find selected group:\n\n{selected_group}"
            )
            return False

        # Activate main EDX window
        main_win = self.wait_for_window(".*initial view.*")
        main_win.set_focus()

        # -----------------------------------------------------
        # Set all transmitters to active for selected group
        # except Master Group
        # -----------------------------------------------------
        if group_number != 1:
            send_keys("%r")
            send_keys("s")

            tx_win = self.wait_for_window(".*Transmitter/BaseHub/Router Sites Groups.*")
            tx_rect = tx_win.rectangle()

            tx_layer_size = 14

            # AutoIt:
            # MouseClick("left",$TxWinPos[0]+56,$TxWinPos[1]+73+($number-1)*$TxLayerSize,2)
            click_x = tx_rect.left + 56
            click_y = tx_rect.top + 73 + (group_number - 1) * tx_layer_size

            pyautogui.click(click_x, click_y, clicks=2, interval=0.1)

            self.wait_for_window(".*Edit Transmitter Group.*")
            time.sleep(0.2)

            # AutoIt:
            # ControlClick("Edit Transmitter Group","Set all Transmitters to active","[CLASS:Button; INSTANCE:17]")
            self.click_control_by_class_instance(
                "Edit Transmitter Group",
                control_type="Button",
                instance=17
            )

            time.sleep(0.1)

            # AutoIt:
            # ControlClick("Edit Transmitter Group","OK","[CLASS:Button; INSTANCE:11]")
            self.click_control_by_class_instance(
                "Edit Transmitter Group",
                control_type="Button",
                instance=11
            )

            self.wait_for_window(".*Transmitter/BaseHub/Router Sites Groups.*")

            # AutoIt:
            # MouseClick("left",$TxWinPos[0]+222,$TxWinPos[1]+435)
            pyautogui.click(tx_rect.left + 222, tx_rect.top + 435)

        # -----------------------------------------------------
        # Open Map Layers
        # -----------------------------------------------------
        self.wait_for_window(".*initial view.*")

        send_keys("%d")
        send_keys("{RIGHT 1}")
        send_keys("{DOWN}")
        send_keys("{ENTER}")

        self.wait_for_window(".*Map Layers.*")

        dlg = Desktop(backend="win32").window(title="Map Layers")
        dlg.set_focus()
        listview = dlg.child_window(class_name="SysListView32").wrapper_object()
        listview.set_focus()
        listview.get_item("transmitter/base/hub sites").select()

        # Open style dialog
        send_keys("%s")

        self.wait_for_window(".*Transmitter/Sector Site Style.*")
        style_rect = self.get_window_rect("Transmitter/Sector Site Style")

        time.sleep(0.05)

        # AutoIt:
        # PixelSearch($styleWin[0]+50, $styleWin[1], $styleWin[0]+$styleWin[3], $styleWin[1]+300, 0x0078D7)
        #
        # 0x0078D7 = RGB(0, 120, 215)
        found_pixel = self.pixel_search(
            style_rect.left + 50,
            style_rect.top,
            style_rect.right,
            style_rect.top + 300,
            target_rgb=(0, 120, 215)
        )

        if found_pixel is None:
            messagebox.showerror(
                "Pixel Search Failed",
                "Could not find the selected blue row in Transmitter/Sector Site Style."
            )
            return False

        pyautogui.click(found_pixel[0] + 4, found_pixel[1] + 4)
        pyautogui.moveTo(found_pixel[0] + 4, found_pixel[1] + 4)

        send_keys("{UP 20}")

        if selected_group.lower() != "master group":
            for _ in range(group_number - 1):
                send_keys("{DOWN}")
                time.sleep(0.02)

        send_keys("{SPACE}")

        # OK on Transmitter/Sector Site Style
        send_keys("%o")

        # Wait until style window closes
        Desktop(backend="win32").window(
            title_re=".*Transmitter/Sector Site Style.*"
        ).wait_not("visible", timeout=10)

        time.sleep(0.1)

        # OK on Map Layers
        send_keys("%o")

        Desktop(backend="win32").window(
            title_re=".*Map Layers.*"
        ).wait_not("visible", timeout=10)

        # -----------------------------------------------------
        # Area Study Details
        # -----------------------------------------------------
        send_keys("%s")
        send_keys("a")

        self.wait_for_window(".*Area Studies for Map View.*")

        # AutoIt:
        # ControlClick("Area Studies for Map View", "", "[CLASS:Button; INSTANCE:2]")
        self.click_control_by_class_instance(
            "Area Studies for Map View",
            control_type="Button",
            instance=2
        )

        self.wait_for_window(".*Area Study Details.*")

        send_keys("{TAB 2}")
        send_keys("{UP 20}")

        if selected_group.lower() != "master group":
            for _ in range(group_number - 1):
                send_keys("{DOWN}")
                time.sleep(0.02)

        time.sleep(0.2)

        send_keys("%o")

        Desktop(backend="uia").window(
            title_re=".*Area Study Details.*"
        ).wait_not("visible", timeout=10)

        self.wait_for_window(".*Area Studies for Map View.*")

        # AutoIt:
        # ControlClick("Area Studies for Map View", "", "[CLASS:Button; INSTANCE:7]")
        self.click_control_by_class_instance(
            "Area Studies for Map View",
            control_type="Button",
            instance=7
        )

        return True
    


class GroupSelectorApp:
    def __init__(self, project_name):
        self.project_name = project_name

        # Create window
        self.root = tk.Tk()
        self.root.title("Select EDX Group")
        self.root.geometry("300x120")
        self.root.attributes('-topmost', True)

        # Label
        ttk.Label(self.root, text="Select Transmitter Group:").pack(pady=5)

        # Combobox
        self.group_combo = ttk.Combobox(self.root, width=30, state="readonly")
        self.group_combo.pack(pady=5)

        # Button
        ttk.Button(self.root, text="Run", command=self.run_switch).pack(pady=5)

        # Create EDX automation object
        self.edx = EDXAutomation(
            group_combo=self.group_combo,   # now valid
            file_name=self.project_name
        )

        # Populate dropdown
        if not self.edx.populate_groups():
            messagebox.showerror("Error", "Failed to load groups.")
            self.root.destroy()

    def run_switch(self):
        selected = self.group_combo.get()

        if not selected:
            messagebox.showwarning("No Selection", "Select a group first.")
            return

        self.root.iconify()

        def task():
            try:
                self.edx.switch_groups()
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally:
                self.root.after(0, self.root.destroy)

        threading.Thread(target=task, daemon=True).start()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    project_name = sys.argv[1] if len(sys.argv) > 1 else ""
    GroupSelectorApp(project_name).run()
