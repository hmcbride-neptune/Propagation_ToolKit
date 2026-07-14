import time
import tkinter as tk
from tkinter import ttk, messagebox
import pyautogui
import pyperclip
from pywinauto import Desktop
from pywinauto.keyboard import send_keys


class EDXAutomation:
    def __init__(self, study_type, startfrom):
        self.study_type = study_type
        self.startfrom = int(startfrom)

        self.total_finds = 0

        self.find_num_1 = 1
        self.find_num_2 = 1
        self.find_num_3 = 1
        self.find_num_4 = 1
        self.find_num_5 = 1

        self.col_array = []

    # --------------------------------------------------
    # Main entry point called by GUI
    # --------------------------------------------------

    def main(self):
        return self.name_finds2()

    # --------------------------------------------------
    # Window / UI helpers
    # --------------------------------------------------

    def wait_for_window(self, title_re, backend="win32", timeout=10):
        win = Desktop(backend=backend).window(title_re=title_re)
        # win.wait("exists visible ready", timeout=timeout)
        win.set_focus()
        return win

    def get_window_rect(self, win):
        rect = win.rectangle()
        return rect.left, rect.top, rect.width(), rect.height()

    def pixel_search(self, x1, y1, x2, y2, target_color):
        width = x2 - x1
        height = y2 - y1

        if width <= 0 or height <= 0:
            return None

        screenshot = pyautogui.screenshot(region=(x1, y1, width, height))

        target_r = (target_color >> 16) & 0xFF
        target_g = (target_color >> 8) & 0xFF
        target_b = target_color & 0xFF

        img_width, img_height = screenshot.size

        for y in range(img_height):
            for x in range(img_width):
                r, g, b = screenshot.getpixel((x, y))[:3]

                if r == target_r and g == target_g and b == target_b:
                    return x1 + x, y1 + y

        return None

    def copy_selected_text(self):
        send_keys("^c")
        time.sleep(0.06)
        return pyperclip.paste()

    def right_3_as_int(self, text, default=0):
        try:
            return int(text[-3:])
        except Exception:
            return default

    def autoit_mid_digit_check(self, text):
        """
        AutoIt:
            StringIsDigit(StringMid(text, 3, 1))

        AutoIt is 1-based.
        Python is 0-based.
        Position 3 = index 2.
        """
        if len(text) < 3:
            return False

        return text[2].isdigit()

    def send_find_name(self, study_num, d):
        find_name = f"{study_num}F{d:03d}"
        send_keys(find_name)
        return find_name

    def click_sector_ok(self, sector_win):
        try:
            sector_win.child_window(auto_id="m_BtnOK").click_input()
            return
        except Exception:
            pass

        try:
            sector_win.child_window(title="&OK").click_input()
            return
        except Exception:
            pass

        send_keys("{ENTER}")

    # --------------------------------------------------
    # AutoIt getArray() converted
    # --------------------------------------------------

    def get_array(self):
        """
        Python equivalent of AutoIt getArray().
        """

        # Ctrl + Shift + Down
        send_keys("^+{DOWN}")
        time.sleep(0.3)

        # Ctrl + C
        send_keys("^c")
        time.sleep(0.3)

        collectors = pyperclip.paste()
        self.col_array = collectors.splitlines()

        self.total_finds = len(self.col_array) + 1

        find_num_arr_1 = [x for x in self.col_array if x.startswith("3F0")]
        find_num_arr_2 = [x for x in self.col_array if x.startswith("4F0")]
        find_num_arr_3 = [x for x in self.col_array if x.startswith("5F0")]
        find_num_arr_4 = [x for x in self.col_array if x.startswith("m")]
        find_num_arr_5 = [x for x in self.col_array if x.startswith("C")]

        self.find_num_1 = len(find_num_arr_1) + 1
        self.find_num_2 = len(find_num_arr_2) + 1
        self.find_num_3 = len(find_num_arr_3) + 1
        self.find_num_4 = len(find_num_arr_4) + 1
        self.find_num_5 = len(find_num_arr_5) + 1

        return {
            "total_finds": self.total_finds,
            "find_num_1": self.find_num_1,
            "find_num_2": self.find_num_2,
            "find_num_3": self.find_num_3,
            "find_num_4": self.find_num_4,
            "find_num_5": self.find_num_5,
            "col_array": self.col_array,
        }

    # --------------------------------------------------
    # Rename logic
    # --------------------------------------------------

    def handle_transmitter_name(self, transmitter_text, study_num, d):
        first_char = transmitter_text[:1]

        all_prefixes = {"3", "4", "5", "m", "C"}
        other_prefixes = all_prefixes - {str(study_num)}

        did_change = False

        if str(study_num) == "3":

            if first_char == "3":
                if "3find" in transmitter_text:
                    d = self.right_3_as_int(transmitter_text) + 1

                elif "3F" in transmitter_text:
                    if self.autoit_mid_digit_check(transmitter_text):
                        d = self.right_3_as_int(transmitter_text) + 1

                send_keys("{ENTER}")

            elif first_char in other_prefixes:
                send_keys("{ENTER}")

            else:
                self.send_find_name(study_num, d)
                send_keys("{TAB}")
                self.send_find_name(study_num, d)

                d += 1
                send_keys("{DOWN}")
                send_keys("{LEFT}")
                did_change = True

        elif str(study_num) == "4":

            if first_char == "4":
                if "find" in transmitter_text:
                    d = self.right_3_as_int(transmitter_text) + 1

                elif "4F" in transmitter_text:
                    if self.autoit_mid_digit_check(transmitter_text):
                        d = self.right_3_as_int(transmitter_text) + 1

                send_keys("{ENTER}")

            elif first_char in other_prefixes:
                send_keys("{ENTER}")

            else:
                self.send_find_name(study_num, d)
                send_keys("{TAB}")
                self.send_find_name(study_num, d)

                d += 1
                send_keys("{DOWN}")
                send_keys("{LEFT}")
                did_change = True

        elif str(study_num) == "5":

            if first_char == "5":
                if "find" in transmitter_text:
                    d = self.right_3_as_int(transmitter_text) + self.find_num_3

                elif "5F" in transmitter_text:
                    if self.autoit_mid_digit_check(transmitter_text):
                        d = self.right_3_as_int(transmitter_text) + 1

                send_keys("{ENTER}")

            elif first_char in other_prefixes:
                send_keys("{ENTER}")

            else:
                self.send_find_name(study_num, d)
                send_keys("{TAB}")
                self.send_find_name(study_num, d)

                d += 1
                send_keys("{DOWN}")
                send_keys("{LEFT}")
                did_change = True

        elif str(study_num) == "m":

            if first_char == "m":
                if "find" in transmitter_text:
                    d = self.right_3_as_int(transmitter_text) + 1 + self.find_num_4

                elif "mF" in transmitter_text:
                    if self.autoit_mid_digit_check(transmitter_text):
                        d = self.right_3_as_int(transmitter_text) + 1 + self.find_num_4

                send_keys("{ENTER}")

            elif first_char in other_prefixes:
                send_keys("{ENTER}")

            else:
                self.send_find_name(study_num, d)
                send_keys("{TAB}")
                self.send_find_name(study_num, d)

                d += 1
                send_keys("{DOWN}")
                send_keys("{LEFT}")
                did_change = True

        elif str(study_num) == "C":

            if first_char == "C":
                if "find" in transmitter_text:
                    d = self.right_3_as_int(transmitter_text) + 1 + self.find_num_5

                elif "CF" in transmitter_text:
                    if self.autoit_mid_digit_check(transmitter_text):
                        d = self.right_3_as_int(transmitter_text) + 1 + self.find_num_5

                send_keys("{ENTER}")

            elif first_char in other_prefixes:
                send_keys("{ENTER}")

            else:
                self.send_find_name(study_num, d)
                send_keys("{TAB}")
                self.send_find_name(study_num, d)

                d += 1
                send_keys("{DOWN}")
                send_keys("{LEFT}")
                did_change = True

        return d, did_change

    # --------------------------------------------------
    # AutoIt NameFinds2() converted
    # --------------------------------------------------

    def name_finds2(self):
        d = 1
        num_checked = 0
        change_count = 0

        if self.study_type == "V5":
            study_num = "5"
        elif self.study_type == "V4":
            study_num = "4"
        elif self.study_type == "V3":
            study_num = "3"
        elif self.study_type == "450":
            study_num = "C"
        elif self.study_type == "m":
            study_num = "m"
        else:
            raise ValueError(f"Unsupported study_type: {self.study_type}")

        time.sleep(0.5)

        # Activate EDX initial view
        self.wait_for_window(r".*-\s*\[initial view\].*", backend="win32", timeout=10)

        # Alt + R
        send_keys("%r")
        time.sleep(0.1)

        send_keys("{DOWN 8}")
        send_keys("{RIGHT}")
        send_keys("{DOWN}")
        send_keys("{ENTER}")

        # Wait for Sector Details
        sector_win = self.wait_for_window(r".*Sector Details.*", backend="win32", timeout=10)

        sect_left, sect_top, sect_width, sect_height = self.get_window_rect(sector_win)

        # Build array and find existing names
        self.get_array()

        # Ctrl + Up
        send_keys("^{UP}")
        time.sleep(0.1)

        # Click row/search area
        pyautogui.click(
            sect_left + 258,
            sect_top + 82,
            clicks=2,
            interval=0.05
        )

        time.sleep(0.1)

        # Type start collector
        send_keys(str(self.startfrom))
        send_keys("{ENTER}")
        time.sleep(0.3)

        # Find highlighted row
        highlighted = self.pixel_search(
            sect_left + 257,
            sect_top + 100,
            sect_left + 320,
            sect_top + sect_height,
            0x3399FF
        )

        if highlighted is None:
            highlighted = self.pixel_search(
                sect_left + 257,
                sect_top + 100,
                sect_left + 320,
                sect_top + sect_height,
                0x0078D7
            )

        if highlighted is None:
            highlighted = self.pixel_search(
                sect_left + 257,
                sect_top + 100,
                sect_left + 320,
                sect_top + sect_height,
                0x0078D4
            )

        if highlighted is None:
            raise RuntimeError("Could not find highlighted row color in Sector Details.")

        pyautogui.click(highlighted[0] + 5, highlighted[1])
        time.sleep(0.2)

        max_checks = self.total_finds - self.startfrom

        while num_checked < max_checks:
            transmitter_text = self.copy_selected_text()

            d, did_change = self.handle_transmitter_name(
                transmitter_text=transmitter_text,
                study_num=study_num,
                d=d
            )

            if did_change:
                change_count += 1

            num_checked += 1

        send_keys("{UP}")

        self.click_sector_ok(sector_win)

        return change_count


# ---------------- GUI ---------------- #

def run_script():
    study_type = study_var.get()
    collector = collector_entry.get()

    if not collector.isdigit():
        messagebox.showerror("Input Error", "Collector must be a number")
        return

    try:
        app = EDXAutomation(study_type, int(collector))
        app.main()
        root.destroy()

    except Exception as e:
        messagebox.showerror("Automation Error", str(e))


root = tk.Tk()
root.title("Rename Finds")

root.geometry("250x140")
root.resizable(False, False)

tk.Label(root, text="Study Type:").grid(
    row=0,
    column=0,
    padx=10,
    pady=5,
    sticky="w"
)

study_var = tk.StringVar(value="V4")

study_dropdown = ttk.Combobox(
    root,
    textvariable=study_var,
    values=["V4", "V3", "450", "V5"],
    width=10,
    state="readonly"
)

study_dropdown.grid(
    row=0,
    column=1,
    padx=10,
    pady=5
)

tk.Label(root, text="Start from collector:").grid(
    row=1,
    column=0,
    padx=10,
    pady=5,
    sticky="w"
)

collector_entry = tk.Entry(root, width=10)
collector_entry.insert(0, "1")
collector_entry.grid(
    row=1,
    column=1,
    padx=10,
    pady=5
)

start_btn = tk.Button(
    root,
    text="Start",
    command=run_script,
    width=15
)

start_btn.grid(
    row=2,
    column=0,
    columnspan=2,
    pady=10
)

root.mainloop()