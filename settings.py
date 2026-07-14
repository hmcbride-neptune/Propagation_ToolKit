import os
import tkinter as tk
from tkinter import ttk
from datetime import datetime


class SettingsGUI:
    def __init__(self, parent, project_name, username):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("Settings")
        self.window.geometry("420x700")
        self.window.resizable(False, False)

        # Optional: bring to front
        self.window.attributes('-topmost', True)
        self.window.lift()
        self.window.focus_force()

        # Store project name
        if project_name != None:
            self.project_name = project_name
        else:
            self.project_name = "None"

        # Needed to locate this user's settings file.
        self.username = username

        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill="both", expand=True)

        # ------------------------------------------------------------------
        # Project Information
        # ------------------------------------------------------------------
        proj_frame = ttk.LabelFrame(main_frame, text="Project Information", padding=10)
        proj_frame.pack(fill="x", pady=5)

        self.project_name_var = tk.StringVar(value=self.project_name)

        ttk.Label(proj_frame, text="Project Name:").pack(anchor="w")
        ttk.Entry(proj_frame, textvariable=self.project_name_var, state="readonly").pack(fill="x", pady=2)

        ttk.Label(proj_frame, text=f"Path: C:\\ProgramData\\EDX\\Projects\\{project_name}").pack(anchor="w", pady=2)

        today = datetime.now().strftime("%m/%d/%Y")
        ttk.Label(proj_frame, text=f"Date: {today}").pack(anchor="w", pady=2)

        # ------------------------------------------------------------------
        # Prop Study Code Settings
        # ------------------------------------------------------------------
        settings_frame = ttk.LabelFrame(main_frame, text="Prop Study Code Settings", padding=10)
        settings_frame.pack(fill="x", pady=5)

        ttk.Label(settings_frame, text="Graphical User Interface Transparency").pack(anchor="w")

        self.transparency = tk.DoubleVar(value=0.9)
        ttk.Scale(settings_frame, from_=0.2, to=1.0, variable=self.transparency)\
            .pack(fill="x", pady=5)

        # Checkboxes
        self.take_screenshots = tk.BooleanVar()
        self.fill_map_indicator = tk.BooleanVar(value=True)
        self.move_data_only = tk.BooleanVar(value=True)

        ttk.Checkbutton(
            settings_frame,
            text="Take screenshots automatically instead of opening GUI",
            variable=self.take_screenshots
        ).pack(anchor="w", pady=2)

        # Disabled for now: this feature is not yet available.
        ttk.Checkbutton(
            settings_frame,
            text='Open "Fill-in Map Indicator" instead of Make Presentation when\nbutton is pressed',
            variable=self.fill_map_indicator,
            state="disabled"
        ).pack(anchor="w", pady=2)

        ttk.Checkbutton(
            settings_frame,
            text="Only move Data folder when 'Make Project' is pressed",
            variable=self.move_data_only
        ).pack(anchor="w", pady=2)

        # Dropdown
        ttk.Label(settings_frame, text="Elevation search accuracy:").pack(anchor="w", pady=(8, 0))

        self.accuracy = tk.StringVar(value="Moderate (suggested)")
        ttk.Combobox(
            settings_frame,
            textvariable=self.accuracy,
            values=["Low", "Moderate (suggested)", "High"],
            state="readonly"
        ).pack(fill="x", pady=3)

        # ------------------------------------------------------------------
        # Appearance Settings
        # ------------------------------------------------------------------
        appearance_frame = ttk.LabelFrame(main_frame, text="Appearance Settings", padding=10)
        appearance_frame.pack(fill="x", pady=5)

        self.red_val = tk.IntVar(value=200)
        self.green_val = tk.IntVar(value=150)
        self.blue_val = tk.IntVar(value=80)

        self.create_color_slider(appearance_frame, "RED", self.red_val)
        self.create_color_slider(appearance_frame, "GREEN", self.green_val)
        self.create_color_slider(appearance_frame, "BLUE", self.blue_val)

        # Buttons
        btn_frame = ttk.Frame(appearance_frame)
        btn_frame.pack(fill="x", pady=5)

        ttk.Button(btn_frame, text="Default Settings", command=self.default_settings)\
            .pack(side="left")

        ttk.Button(btn_frame, text="Retrieve Newest Script", command=self.retrieve_script, state="disabled")\
            .pack(side="right")
        

        # ------------------------------------------------------------------
        # Save Button
        # ------------------------------------------------------------------
        ttk.Button(main_frame, text="Save All", command=self.save_all)\
            .pack(fill="x", pady=10)

        # Load this user's persisted settings (falls back to defaults if the
        # file does not exist) and apply them to the widgets above.
        lines = self.read_user_settings()

        # GUI Transparency: stored 0-255, the scale/window use a 0.2-1.0 alpha.
        try:
            alpha = int(lines[1]) / 255.0
        except (ValueError, TypeError):
            alpha = 0.9
        alpha = max(0.2, min(1.0, alpha))
        self.transparency.set(alpha)
        self.parent.attributes("-alpha", alpha)

        # Background Color: stored as 0xRRGGBB or #RRGGBB.
        r, g, b = self._parse_rgb(lines[2])
        if r is not None:
            self.red_val.set(r)
            self.green_val.set(g)
            self.blue_val.set(b)

        # Checkbox / dropdown preferences.
        self.take_screenshots.set(self._parse_bool(lines[0]))   # Open_Screenshot_GUI_Pref
        self.fill_map_indicator.set(self._parse_bool(lines[4]))  # Button10Preference
        self.move_data_only.set(self._parse_bool(lines[5]))      # MoveDataPreference

        if lines[6] in ("Low", "Moderate (suggested)", "High"):  # searchPref
            self.accuracy.set(lines[6])

        # Values not tied to a widget are preserved so save_all round-trips them.
        self.view = lines[3]             # view
        self.ribbon_view = lines[7]      # Ribbonview
        self.ss_bounds = lines[8:12]     # SSBoundsSave1..4

    def settings_file_path(self):
        return fr"\\TetonRiver\prop\template\zAutoit\User Settings\{self.username} settings.txt"

    def read_user_settings(self):
        settings_file = self.settings_file_path()

        # Defaults used when the file is missing or a value is absent. These
        # mirror the widget defaults set up in __init__. Index positions match
        # the load block above.
        defaults = [
            "False",                # 0  Open_Screenshot_GUI_Pref
            "230",                  # 1  PersonalGUITransparency (~0.9 alpha)
            "0xc89650",             # 2  Background color (200, 150, 80)
            "",                     # 3  view
            "True",                 # 4  Button10Preference
            "True",                 # 5  MoveDataPreference
            "Moderate (suggested)", # 6  searchPref
            "",                     # 7  Ribbonview
            "",                     # 8  SSBoundsSave1
            "",                     # 9  SSBoundsSave2
            "",                     # 10 SSBoundsSave3
            "",                     # 11 SSBoundsSave4
        ]

        if not os.path.exists(settings_file):
            return defaults

        # Settings are positional (line N == a specific setting), so keep every
        # line rather than dropping blanks, which would shift later values.
        with open(settings_file, "r", encoding="utf-8") as f:
            lines = [line.strip() for line in f.read().splitlines()]

        # Pad any missing trailing values with their defaults.
        while len(lines) < len(defaults):
            lines.append(defaults[len(lines)])

        return lines

    # ----------------------------------------------------------------------
    # Helpers: parse persisted values
    # ----------------------------------------------------------------------
    @staticmethod
    def _parse_bool(value):
        return str(value).strip().lower() in ("true", "1", "yes", "on")

    @staticmethod
    def _parse_rgb(value):
        text = str(value).strip()
        if text.startswith("0x"):
            text = text[2:]
        elif text.startswith("#"):
            text = text[1:]
        if len(text) == 6:
            try:
                return int(text[0:2], 16), int(text[2:4], 16), int(text[4:6], 16)
            except ValueError:
                pass
        return None, None, None


    # ----------------------------------------------------------------------
    # Helper: Color Sliders
    # ----------------------------------------------------------------------
    def create_color_slider(self, parent, label, variable):
        ttk.Label(parent, text=label).pack(anchor="w")
        ttk.Scale(parent, from_=0, to=255, variable=variable)\
            .pack(fill="x", pady=5)

    # ----------------------------------------------------------------------
    # Default Settings
    # ----------------------------------------------------------------------
    def default_settings(self):
        self.transparency.set(0.9)
        self.red_val.set(200)
        self.green_val.set(150)
        self.blue_val.set(80)

    # ----------------------------------------------------------------------
    # Retrieve Script (hook your logic here)
    # ----------------------------------------------------------------------
    def retrieve_script(self):
        print("Retrieve newest script clicked")

    # ----------------------------------------------------------------------
    # Save All Settings
    # ----------------------------------------------------------------------
    def save_all(self):
        r = self.red_val.get()
        g = self.green_val.get()
        b = self.blue_val.get()

        data = {
            "project_name": self.project_name_var.get(),
            "transparency": self.transparency.get(),
            "take_screenshots": self.take_screenshots.get(),
            "fill_map_indicator": self.fill_map_indicator.get(),
            "move_data_only": self.move_data_only.get(),
            "accuracy": self.accuracy.get(),
            "rgb": {
                "r": r,
                "g": g,
                "b": b,
            }
        }

        # print("Saved Settings:")
        # print(data)

        # Persist to the per-user settings file (overwrite if it exists,
        # create it otherwise).
        self.write_user_settings()

        # ✅ Convert RGB → HEX

        color_hex = f"#{r:02x}{g:02x}{b:02x}"

        # Apply transparency to the main (parent) window.
        alpha = max(0.2, min(1.0, self.transparency.get()))
        self.parent.attributes("-alpha", alpha)

        # Apply to main window
        self.parent.configure(bg=color_hex)

        # Apply to main grid container
        if hasattr(self.parent, "grid_frame"):
            self.parent.grid_frame.configure(bg=color_hex)

        # print(f"Applied background color: {color_hex}")

        self.window.destroy()

    def write_user_settings(self):
        r = self.red_val.get()
        g = self.green_val.get()
        b = self.blue_val.get()

        # Build the 12 positional lines in the same order read_user_settings
        # expects. Values without a widget are preserved from load.
        lines = [
            "True" if self.take_screenshots.get() else "False",  # 0  Open_Screenshot_GUI_Pref
            str(int(round(self.transparency.get() * 255))),       # 1  PersonalGUITransparency
            f"0x{r:02x}{g:02x}{b:02x}",                           # 2  Background color
            getattr(self, "view", ""),                            # 3  view
            "True" if self.fill_map_indicator.get() else "False", # 4  Button10Preference
            "True" if self.move_data_only.get() else "False",     # 5  MoveDataPreference
            self.accuracy.get(),                                  # 6  searchPref
            getattr(self, "ribbon_view", ""),                     # 7  Ribbonview
        ]
        # 8-11  SSBoundsSave1..4
        ss_bounds = getattr(self, "ss_bounds", [])
        for i in range(4):
            lines.append(ss_bounds[i] if i < len(ss_bounds) else "")

        settings_file = self.settings_file_path()
        try:
            os.makedirs(os.path.dirname(settings_file), exist_ok=True)
            with open(settings_file, "w", encoding="utf-8") as f:
                f.write("\n".join(lines) + "\n")
            print(f"Saved settings to: {settings_file}")
        except OSError as e:
            print(f"Could not write settings file '{settings_file}': {e}")

