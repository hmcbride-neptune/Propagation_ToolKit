import sys
import os
import runpy
import tempfile

# ----------------------------------------------------------------------------
# Frozen sub-tool dispatcher
# ----------------------------------------------------------------------------
# When packaged with PyInstaller there is no separate python.exe, so the GUI
# can't launch the other scripts as "python script.py". Instead it re-launches
# itself as "Main_GUI.exe --run <module> [args...]"; this block intercepts that
# and runs the requested bundled module's __main__ instead of opening the GUI.
if len(sys.argv) >= 3 and sys.argv[1] == "--run":
    _module = sys.argv[2]
    # Present the target script with a normal argv (its own name + its args).
    sys.argv = [_module] + sys.argv[3:]
    runpy.run_module(_module, run_name="__main__", alter_sys=True)
    sys.exit(0)

import tkinter as tk
from tkinter import messagebox
import new_request
import Data_Import
import subprocess
import move_data
from datetime import datetime
import win32gui
import settings
import re
import threading
import getpass

from PIL import Image, ImageTk

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGOS_DIR = os.path.join(BASE_DIR, "Logos")


def script_cmd(module_name):
    """Build the command to launch one of the sibling sub-tool scripts.

    From source this runs "python <module>.py"; when frozen it runs
    "Main_GUI.exe --run <module>", which the dispatcher at the top of this file
    turns back into that module's __main__.
    """
    if getattr(sys, "frozen", False):
        return [sys.executable, "--run", module_name]
    return [sys.executable, os.path.join(BASE_DIR, module_name + ".py")]

# Windows flag that stops a child process from opening its own console window.
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

def gui_python():
    """Return a windowless Python interpreter for launching external GUI scripts.

    When frozen, sys.executable is the packaged .exe (not a real Python), so
    fall back to whatever Python is on PATH. Prefer pythonw.exe so no console
    window flashes up behind the launched tool.
    """
    if not getattr(sys, "frozen", False):
        exe = sys.executable
        pythonw = os.path.join(os.path.dirname(exe), "pythonw.exe")
        return pythonw if os.path.exists(pythonw) else exe
    return "pythonw"

# Last-created project name (set by launch_new_request)
PROJECT_NAME = None
PROJECT_SYNC_MS = 2000

COLUMNS = 24                      # single row of 24 buttons
ROWS = 1
ICON_SIZE = 30                    # px; icons are square so the buttons are square
BUTTON_PADDING = 2               # px of empty space around each button
RELEASE_DATE = "260714"

USERNAME = getpass.getuser()


class Tooltip:
    """Minimal hover tooltip -- a row of bare icons is hard to read otherwise."""

    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _event=None):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self.tip, text=self.text, background="#ffffe0",
            relief="solid", borderwidth=1, padx=4, pady=1,
        ).pack()

    def hide(self, _event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None
   
def get_project_name_from_edx():
    """
    Gets the active EDX project name from the open EDX SignalPro window.

    Expected title example:
    EDX® SignalPro®  for 64-bit - 97035 Lake Oswego Modeling OR R900 - [initial view]
    """

    found_projects = []

    def enum_callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return

        title = win32gui.GetWindowText(hwnd)

        if not title:
            return

        if "[initial view]" not in title:
            return

        if "EDX" not in title and "SignalPro" not in title:
            return

        
        match = re.search(
            r"64-bit\s*-\s*(.*?)\s*-\s*\[initial view\]",
            title,
            re.IGNORECASE
        )

        if match:
            found_projects.append(match.group(1).strip())

    try:
        win32gui.EnumWindows(enum_callback, None)
    except Exception:
        return None

    if found_projects:
        return found_projects[0]

    return None
    
class MainGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{RELEASE_DATE}")
        self.configure(padx=BUTTON_PADDING, pady=BUTTON_PADDING)
        self.attributes('-topmost', True)

        # Clear any screenshot boundary left over from a previous session. The
        # screenshot tool persists the boundary the user sets (Auto/Manual) to
        # this temp file so it survives relaunches while this GUI is open;
        # wiping it here means each new Main GUI session starts with the
        # default boundary from settings.
        try:
            os.remove(os.path.join(tempfile.gettempdir(), "edx_snip_bounds.txt"))
        except OSError:
            pass

        # Id of the pending sync_project_name_from_edx() callback so it can be
        # cancelled on shutdown (otherwise it fires against a destroyed window).
        self._project_sync_after_id = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # ----------------------------------------------------------------------------
        # Button table:  (icon filename in Logos/, tooltip text, command)
        # ----------------------------------------------------------------------------
        self.BUTTONS = [
            ("New Request.bmp",  "Make new request folder",        self.launch_new_request),
            ("fileformat2.png",       "Format files. Meters.csv, esri export, etc.",       self.launch_format_files),
            ("Meters_combined.bmp", "Create Meter files from Inbound and Outbound",  self.launch_meters_combined),
            ("nEW pROJECT.bmp",    "Moves Data Folder",          self.launch_move_data),
            ("addfile.bmp",      "Add Project Folder. Browse to folder's xReports subfolder",        self.launch_add_project_files),
            ("terrain.bmp",       "Creates terrain map based on max and min in study area",     lambda: self.launch_terrain()),
            ("switch group.bmp",       "Switch Group being Displayed/Calculated",        self.launch_switch_group),
            ("measurement.bmp",   "Measurement Analysis",       self.launch_measurement_analysis),
            ("demographic.png",  "Demographic Studies",       self.launch_demographics_study),
            ("finds.bmp",    "Re-name Find Locations",        self.launch_rename_finds),
            ("screenshot.png",   "Take screenshots",        self.launch_screenshot), # NEED UPDATED
            ("showhide.bmp","Show or Hide inactive transmitter locations",  self.launch_show_hide),
            ("hourglass.png",      "Switch between daily, hourly, and billing",   self.launch_switch_DHB),
            ("movecoverage.bmp",    "Move Coverage layer either above or below meters",  self.launch_move_layer),
            ("export (1).bmp",      "Export Transmitters",            lambda: self.launch_export_transmitters()),
            ("presentation.png",  "Either make presentation or Fill in Map indicator based on preferences",      self.launch_presentation),
            ("reset.bmp",      "Reset All Settings",           self.reset),
            ("airplane.png",     "Run FAA/ASR Script",       self.launch_faa_tool),
            ("zipup.png", "Zip and Upload",       self.launch_zipper),
            ("question.png",      "Assumptions",           self.launch_assumptions),
            ("TailoredLogo.bmp",   "Create a tailored output",         self.launch_tailored_output),
            ("calc.bmp", "Get by collector coverage statistics",       self.launch_best_collector),
            ("dataImport.bmp",     "Data Import",           self.launch_data_import),
            ("settings.png",     "Change Settings",        lambda: self.launch_settings()),
        ]
                
        # Keep references so the images aren't garbage collected.
        self.icons = [self.load_icon(icon) for icon, _, _ in self.BUTTONS]
        self.buttons = []

        
        self.grid_frame = tk.Frame(self)
        self.grid_frame.pack(fill="both", expand=True)


        for index, (_, tip, command) in enumerate(self.BUTTONS):
            row, col = divmod(index, COLUMNS)
            btn = tk.Button(
                self.grid_frame,
                image=self.icons[index],
                width=ICON_SIZE,
                height=ICON_SIZE,
                command=command,
            )
            btn.grid(
                row=row,
                column=col,
                padx=BUTTON_PADDING,
                pady=BUTTON_PADDING,
                sticky="nsew",
            )
            Tooltip(btn, tip)
            self.buttons.append(btn)

        # Even spacing: every column/row gets equal weight and uniform sizing.
        for c in range(COLUMNS):
            self.grid_frame.columnconfigure(c, weight=1, uniform="btn")
        for r in range(ROWS):
            self.grid_frame.rowconfigure(r, weight=1, uniform="btn")

        self.resizable(False, False)
        self.sync_project_name_from_edx()
        
        lines = self.read_user_settings()
        # Background Color
        color = lines[2]
        if color.startswith("0x"):
            color = "#" + color[2:]
        self.grid_frame.configure(bg=color)

        # GUI Transparency
        PersonalGUITransparency = int(lines[1])
        alpha = PersonalGUITransparency / 255.0
        self.attributes("-alpha", alpha)

        Open_Screenshot_GUI_Pref = lines[0]
        # view = lines[3]
        Button10Preference = lines[4]
        MoveDataPreference = lines[5]
        searchPref = lines[6]
        Ribbonview = lines[7]
        SSBoundsSave1 = lines[8]
        SSBoundsSave2 = lines[9]
        SSBoundsSave3 = lines[10]
        SSBoundsSave4 = lines[11]
        
    def set_project_name(self, project_name, source=""):
            """
            Sets the global PROJECT_NAME and updates the GUI title.
            """
            global PROJECT_NAME

            if not project_name:
                return

            project_name = str(project_name).strip()

            if not project_name:
                return

            if PROJECT_NAME != project_name:
                PROJECT_NAME = project_name

                if source:
                    print(f"PROJECT_NAME set from {source}: {PROJECT_NAME}")
                else:
                    print(f"PROJECT_NAME set: {PROJECT_NAME}")

            self.title(f"Propagation - Main | {PROJECT_NAME}")
    
    def sync_project_name_from_edx(self):
        """
        Dynamically checks for the active EDX project and updates PROJECT_NAME.

        Important:
        - If EDX is open, PROJECT_NAME updates from EDX.
        - If EDX is not open, PROJECT_NAME is left alone.
        - This allows launch_new_request() to set PROJECT_NAME when no EDX project is open.
        """
        # The window may already be gone when a queued callback fires.
        if not self.winfo_exists():
            return

        edx_project = get_project_name_from_edx()

        if edx_project:
            self.set_project_name(edx_project, "EDX")

        self._project_sync_after_id = self.after(
            PROJECT_SYNC_MS, self.sync_project_name_from_edx
        )

    def on_close(self):
        """Cancel the pending EDX sync callback before tearing down the window."""
        if self._project_sync_after_id is not None:
            try:
                self.after_cancel(self._project_sync_after_id)
            except Exception:
                pass
            self._project_sync_after_id = None
        self.destroy()

    def get_active_project_name(self):
        """
        Returns the best available project name.

        Priority:
        1. Current open EDX project
        2. Existing global PROJECT_NAME
        """
        global PROJECT_NAME

        edx_project = get_project_name_from_edx()

        if edx_project:
            self.set_project_name(edx_project, "EDX")

        return PROJECT_NAME
    
    def launch_new_request(self):
        top = tk.Toplevel()
        top.attributes('-topmost', True)
        top.lift()
        top.focus_force()

        nr = new_request.new_request(top)

        # Wait until the new-request window is closed, then read project name.
        try:
            top.wait_window()
        except Exception:
            pass

        proj = getattr(nr, "project_name", None)

        if proj:
            self.set_project_name(proj, "New Request")

        return PROJECT_NAME  
    
    def launch_settings(self):
        try:
            project_name = self.get_active_project_name()
        except:
            project_name = None

        settings.SettingsGUI(self, project_name, USERNAME)
    
    def launch_zipper(self):
        script = "zipper"
        try:
            subprocess.Popen(script_cmd(script), cwd=BASE_DIR)
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to start Zip tool: {e}")
    
    def launch_meters_combined(self):
        
        base_dir = r"C:\ProgramData\Python"
        root = tk._default_root
        full_path = os.path.join(base_dir, "Meter Combine", "metersCombined.py")
        try:
            root.attributes('-topmost', False)
            root.lower()
        except Exception:
            pass

        subprocess.run([gui_python(), full_path], creationflags=NO_WINDOW)
        root.after(500, lambda: root.attributes('-topmost', True))

    def launch_format_files(self):
        self.attributes('-topmost', False)
        self.iconify()
        project_name = self.get_active_project_name()

        script = "format_files"
        
        args = script_cmd(script)

        if project_name:
            args.append(project_name)
        def task():
            subprocess.run(args)
            self.after(0, self.restore_window)
        threading.Thread(target=task, daemon=True).start()
        
    def launch_data_import(self):
        # Create the DataImport window as a Toplevel child of the main root
        # without creating an extra blank Toplevel.
        root = tk._default_root
        win = Data_Import.DataImportApp(master=root)
        try:
            win.attributes('-topmost', True)
            win.lift()
            win.focus_force()
        except Exception:
            pass
        return win
    
    def launch_faa_tool(self):
        script = "faa_tool"
        try:
            subprocess.Popen(script_cmd(script), cwd=BASE_DIR)
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to start FAA tool: {e}")
    
    def launch_tailored_output(self):
        
        base_dir = r"C:\ProgramData\Python"
        root = tk._default_root

        try:
            root.attributes('-topmost', False)
            root.lower()
        except Exception:
            pass

        latest_folder = None
        latest_date = None

        for name in os.listdir(base_dir):
            full_path = os.path.join(base_dir, name)

            # Make sure it's a directory and matches pattern
            if os.path.isdir(full_path) and name.startswith("Tailored_Output_"):
                try:
                    # Extract YYMMDD
                    date_part = name.replace("Tailored_Output_", "").strip()
                    # Convert to datetime (assume 20YY)
                    folder_date = datetime.strptime(date_part, "%y%m%d")

                    # Compare
                    if latest_date is None or folder_date > latest_date:
                        latest_date = folder_date
                        latest_folder = full_path
                        latest_datestr = date_part
                        tailored_file = fr"{latest_folder}\Tailored_Output_{latest_datestr}.py"
                except ValueError:
                    # Skip anything that doesn't match YYMMDD
                    continue
        subprocess.run([gui_python(), tailored_file], creationflags=NO_WINDOW)
        root.after(500, lambda: root.attributes('-topmost', True))
    
    def launch_measurement_analysis(self):
        # Get out of the way before the tool drives EDX. This window is topmost
        # (see __init__), and a topmost window fighting for focus makes the
        # pywinauto menu_select / dialog-read steps land on the wrong window,
        # which is why the 'No. of points' read fails when launched from here
        # but works when the script is run directly. Match the other EDX tools.
        self.attributes('-topmost', False)
        self.iconify()

        script = "measurement_analysis"
        args = script_cmd(script)

        def task():
            try:
                subprocess.run(args, cwd=BASE_DIR)
            except Exception as e:
                self.after(0, lambda: tk.messagebox.showerror(
                    "Error", f"Failed to start Measurement Analysis tool: {e}"))
            finally:
                self.after(0, self.restore_window)

        threading.Thread(target=task, daemon=True).start()
    
    def launch_move_data(self):
            try:
                project_name = self.get_active_project_name()
            except:
                project_name = None


            # if not project_name:
            #     messagebox.showerror(
            #         "No Project Found",
            #         "No active project was found.\n\n"
            #         "Open an EDX project or run New Request first."
            #     )
            #     return

            move_data.MoveData.copy_data_folder(project_name)
    
    def launch_best_collector(self):
        subprocess.run(r"C:\ProgramData\Python\Best_Collector\Best_Collector.exe")
    
    def reset(self):
        # Relaunch the app. Frozen: argv[0] is the exe itself; from source we
        # must re-pass the script path (argv includes it).
        if getattr(sys, "frozen", False):
            subprocess.Popen([sys.executable] + sys.argv[1:])
        else:
            subprocess.Popen([sys.executable] + sys.argv)
        sys.exit()
    
    def launch_assumptions(self):
        script = "assumptions"
        try:
            subprocess.Popen(script_cmd(script), cwd=BASE_DIR)
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to start Assumptions tool: {e}")
    
    def launch_add_project_files(self):
        script = "add_project_files"
        try:
            subprocess.Popen(script_cmd(script), cwd=BASE_DIR)
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to start Add Project Files tool: {e}")
    
    def load_icon(self, filename, size=ICON_SIZE):
        """Load a Logos PNG and fit it into a transparent square so buttons stay square."""
        img = Image.open(os.path.join(LOGOS_DIR, filename)).convert("RGBA")
        img.thumbnail((size, size), Image.LANCZOS)
        square = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        square.paste(img, ((size - img.width) // 2, (size - img.height) // 2), img)
        return ImageTk.PhotoImage(square)
    
    def launch_export_transmitters(self):
        project_name = self.get_active_project_name()

        if not project_name:
            messagebox.showerror(
                "No Project Found",
                "No active project was found.\n\n"
                "Open an EDX project or run New Request first."
            )
            return

        script = "export_transmitters"
       
        args = script_cmd(script)

        if project_name:
            args.append(project_name)

        # subprocess.Popen(
        #     args,
        #     creationflags=subprocess.CREATE_NO_WINDOW
        # )
        subprocess.Popen(args)
    
    def launch_terrain(self):
        self.attributes('-topmost', False)
        self.iconify()
        project_name = self.get_active_project_name()
        if not project_name:
            messagebox.showerror(
                "No Project Found",
                "No active project was found.\n\n"
                "Open an EDX project or run New Request first."
            )
            return
        script = "terrain"
        
        args = script_cmd(script)

        if project_name:
            args.append(project_name)
        def task():
            subprocess.run(args)
            self.after(0, self.restore_window)
        threading.Thread(target=task, daemon=True).start()
    
    def launch_demographics_study(self):
        self.attributes('-topmost', False)
        self.iconify()
        project_name = self.get_active_project_name()
        if not project_name:
            messagebox.showerror(
                "No Project Found",
                "No active project was found.\n\n"
                "Open an EDX project or run New Request first."
            )
            return
        script = "demographics"
        
        args = script_cmd(script)

        if project_name:
            args.append(project_name)
        def task():
            subprocess.run(args)
            self.after(0, self.restore_window)
        threading.Thread(target=task, daemon=True).start()
    
    def restore_window(self):
        self.deiconify()
        self.attributes('-topmost', True)
    
    def launch_switch_group(self):
        self.attributes('-topmost', False)
        self.iconify()

        project_name = self.get_active_project_name()

        if not project_name:
            messagebox.showerror(
                "No Project Found",
                "No active project was found.\n\n"
                "Open an EDX project or run New Request first."
            )
            self.restore_window()
            return
        script = "switch_group"

        args = script_cmd(script) + [project_name]

        def task():
            subprocess.run(args)
            self.after(0, self.restore_window)
        threading.Thread(target=task, daemon=True).start()
        
    def launch_move_layer(self):
        self.attributes('-topmost', False)
        self.iconify()
        project_name = self.get_active_project_name()
        if not project_name:
            messagebox.showerror(
                "No Project Found",
                "No active project was found.\n\n"
                "Open an EDX project or run New Request first."
            )
            return
        script = "move_layer"
        
        args = script_cmd(script)

        if project_name:
            args.append(project_name)
        def task():
            subprocess.run(args)
            self.after(0, self.restore_window)
        threading.Thread(target=task, daemon=True).start()

    def launch_show_hide(self):
        self.attributes('-topmost', False)
        self.iconify()
        project_name = self.get_active_project_name()
        if not project_name:
            messagebox.showerror(
                "No Project Found",
                "No active project was found.\n\n"
                "Open an EDX project or run New Request first."
            )
            return
        script = "show_hide"
        
        args = script_cmd(script)

        if project_name:
            args.append(project_name)
        def task():
            subprocess.run(args)
            self.after(0, self.restore_window)
        threading.Thread(target=task, daemon=True).start()

    def launch_rename_finds(self):
        self.attributes('-topmost', False)
        self.iconify()
        project_name = self.get_active_project_name()
        if not project_name:
            messagebox.showerror(
                "No Project Found",
                "No active project was found.\n\n"
                "Open an EDX project or run New Request first."
            )
            return
        script = "rename_finds"
        
        args = script_cmd(script)

        if project_name:
            args.append(project_name)
        def task():
            subprocess.run(args)
            self.after(0, self.restore_window)
        threading.Thread(target=task, daemon=True).start() 

    def launch_switch_DHB(self):
        self.attributes('-topmost', False)
        self.iconify()
        project_name = self.get_active_project_name()
        if not project_name:
            messagebox.showerror(
                "No Project Found",
                "No active project was found.\n\n"
                "Open an EDX project or run New Request first."
            )
            return
        script = "switch_dhb"
        
        args = script_cmd(script)

        if project_name:
            args.append(project_name)
        def task():
            subprocess.run(args)
            self.after(0, self.restore_window)
        threading.Thread(target=task, daemon=True).start() 

    def launch_presentation(self):
        self.attributes('-topmost', False)
        self.iconify()
        project_name = self.get_active_project_name()

        script = "presentation"
        
        args = script_cmd(script)

        if project_name:
            args.append(project_name)
        def task():
            subprocess.run(args)
            self.after(0, self.restore_window)
        threading.Thread(target=task, daemon=True).start()

    def read_user_settings(self):
        settings_file = fr"\\TetonRiver\prop\template\zAutoit\User Settings\{USERNAME} settings.txt"

        if settings_file:
           
            with open(settings_file, "r", encoding="utf-8") as f:
                lines = [line.strip() for line in f if line.strip()]

            while len(lines) < 12:
                lines.append("")

            return lines

    def launch_screenshot(self):
        self.attributes('-topmost', False)
        self.iconify()

        script = "screenshots"
        
        args = script_cmd(script)

        args.append(USERNAME)
        def task():
            subprocess.run(args)
            self.after(0, self.restore_window)
        threading.Thread(target=task, daemon=True).start()

if __name__ == "__main__":
    MainGUI().mainloop()
