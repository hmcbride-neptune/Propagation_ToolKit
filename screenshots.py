import sys
# Force the process COM apartment to STA (COINIT_APARTMENTTHREADED) BEFORE
# importing pywinauto/comtypes. Otherwise pywinauto initializes COM as MTA,
# which deadlocks Tkinter's native file dialogs (e.g. askdirectory hangs).
sys.coinit_flags = 2

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from contextlib import contextmanager
import time
import os
import tempfile
import pyautogui
import pygetwindow as gw
import win32gui
import win32con
import numpy as np
from PIL import Image, ImageTk
from pywinauto.keyboard import send_keys
from pywinauto import Desktop
from pywinauto import Application

todays_date = datetime.now().strftime("%y%m%d") #YYMMDD

class SnipWindow(tk.Toplevel):

    # Max pixel size of the live preview image; the capture is scaled down to
    # fit inside this box while preserving aspect ratio.
    PREVIEW_W = 200
    PREVIEW_H = 170

    def __init__(self, username, parent=None):
        super().__init__(parent)

        self.parent = parent
        self.username = username
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # Reference to the current preview image to keep it from being GC'd.
        self.preview_image = None

        self.title("Snip")
        self.geometry("700x260")
        self.resizable(False, False)

        self.latitude = ""
        self.longitude = ""

        settings = self.get_user_settings()

        self.top_bound = int(settings[8]) if settings[8] else 0
        self.bottom_bound = int(settings[9]) if settings[9] else 0
        self.left_bound = int(settings[10]) if settings[10] else 0
        self.right_bound = int(settings[11]) if settings[11] else 0

        # If the user set a boundary earlier this Main GUI session (via Auto or
        # Manual), it was saved to a temp file. Prefer it over the settings
        # default so the boundary sticks across relaunches of this script. The
        # Main GUI clears this file on startup, so a new session starts fresh.
        session = self.load_session_bounds()
        if session:
            self.left_bound, self.top_bound, self.right_bound, \
                self.bottom_bound = session

        if str(settings[0]).lower() in ("true", "1", "yes"):
            # Auto mode: run headless. Don't build/show the GUI -- just capture
            # the boundary saved in settings to a numbered "screenshot N" file
            # and exit.
            self.withdraw()
            self.after(1000, self._run_auto_and_exit)
            return

        self.build_gui()

        # Start the live preview loop.
        self.update_preview()

    def get_window(self, title_contains, exclude=None):
        for win in gw.getAllWindows():
            title = win.title.lower()
            if title_contains.lower() not in title:
                continue
            if exclude and exclude.lower() in title:
                continue
            return win
        return None

    def get_project_name(self):
        # The EDX title looks like: "... 64-bit - <project name> - [initial
        # view]". Pull the project name out of the middle, or return "" if the
        # window/markers aren't found.
        edx = self.get_window("SignalPro", exclude="Legend")
        if not edx:
            return ""

        title = edx.title
        if "64-bit -" not in title or "- [" not in title:
            return ""

        name = title.split("64-bit -", 1)[1]
        name = name.split("- [", 1)[0]
        return name.strip()

    def get_default_save_location(self):
        # Default save folder: C:\PROGRAMDATA\EDX\Projects\<project>\Data\Images,
        # or the base Projects folder if the project name can't be determined.
        base_save = r"C:\PROGRAMDATA\EDX\Projects"
        project_name = self.get_project_name()
        if project_name:
            base_save = os.path.join(base_save, project_name, r"Data\Images")
        return base_save

    def _run_auto_and_exit(self):
        # Headless auto mode: capture the saved boundary, then close the app.
        try:
            self.auto_capture()
        finally:
            self.on_close()

    def build_gui(self):

        # =====================================================
        # ROOT LAYOUT
        # =====================================================
        self.columnconfigure(0, weight=3)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # =====================================================
        # VARIABLES & SETTINGS
        # =====================================================
        group1 = ttk.LabelFrame(self, text="Variables and Settings")
        group1.grid(
            row=0,
            column=0,
            padx=10,
            pady=10,
            sticky="nsew"
        )

        group1.columnconfigure(0, weight=1)

        # =====================================================
        # BOUNDARY GROUP
        # =====================================================
        boundary = ttk.LabelFrame(
            group1,
            text="Image Boundary Information"
        )

        boundary.grid(
            row=0,
            column=0,
            padx=10,
            pady=10,
            sticky="ew"
        )

        boundary.columnconfigure(1, weight=1)

        # -----------------------------------------------------
        # LEFT SIDE BUTTONS
        # -----------------------------------------------------

        ttk.Button(
            boundary,
            text="Auto Set Boundaries",
            command=self.edx_bounds
        ).grid(row=0, column=0, padx=5, pady=2, sticky="ew")

        ttk.Button(
            boundary,
            text="Manually Set Boundaries",
            command=self.manual_bounds
        ).grid(row=1, column=0, padx=5, pady=2, sticky="ew")

        ttk.Button(
            boundary,
            text="Preview Current Boundary",
            command=self.preview_boundary
        ).grid(row=2, column=0, padx=5, pady=2, sticky="ew")

        # -----------------------------------------------------
        # STATUS / COORDINATES
        # -----------------------------------------------------

        status_frame = ttk.Frame(boundary)
        status_frame.grid(
            row=0,
            column=1,
            rowspan=3,
            padx=15,
            sticky="nw"
        )

        self.scale_status_label = ttk.Label(
            status_frame,
            text="Scale/Pos Status: Not Set"
        )
        self.scale_status_label.grid(
            row=0,
            column=0,
            sticky="w"
        )

        self.status_label = ttk.Label(
            status_frame,
            text="Status: Unknown"
        )
        self.status_label.grid(
            row=1,
            column=0,
            sticky="w"
        )

        coord_frame = ttk.Frame(status_frame)
        coord_frame.grid(
            row=2,
            column=0,
            pady=(5, 0),
            sticky="w"
        )


        self.labelx1 = ttk.Label(coord_frame, text="x1: 0")
        self.labelx1.grid(row=0, column=0)

        self.labely1 = ttk.Label(coord_frame, text="y1: 0")
        self.labely1.grid(row=0, column=1, padx=15)

        self.labelx2 = ttk.Label(coord_frame, text="x2: 0")
        self.labelx2.grid(row=1, column=0)

        self.labely2 = ttk.Label(coord_frame, text="y2: 0")
        self.labely2.grid(row=1, column=1, padx=15)


        # -----------------------------------------------------
        # RECORD / RETURN
        # -----------------------------------------------------

        btn_frame = ttk.Frame(boundary)
        btn_frame.grid(
            row=0,
            column=2,
            rowspan=3,
            padx=10,
            sticky="n"
        )

        ttk.Button(
            btn_frame,
            text="Record",
            width=10,
            command=self.record
        ).grid(row=0, column=0, pady=2)

        ttk.Button(
            btn_frame,
            text="Return",
            width=10,
            command=self.recall
        ).grid(row=1, column=0, pady=2)

        # =====================================================
        # IMAGE NAME ROW
        # =====================================================

        image_frame = ttk.Frame(group1)
        image_frame.grid(
            row=1,
            column=0,
            padx=10,
            pady=5,
            sticky="ew"
        )

        image_frame.columnconfigure(1, weight=1)

        ttk.Label(
            image_frame,
            text="Image Name:"
        ).grid(row=0, column=0, sticky="w")

        self.name_combo = ttk.Combobox(
            image_frame,
            values=[
                f"{todays_date} Terrain",
                f"{todays_date} Provided",
                f"{todays_date} 99Percent",
                f"{todays_date} 90Percent",
                f"{todays_date} Best Provided"
            ],
            width=20
        )

        self.name_combo.grid(
            row=0,
            column=1,
            padx=5,
            sticky="ew"
        )

        self.name_combo.set(f"{todays_date} Terrain")

        self.type_combo = ttk.Combobox(
            image_frame,
            values=[
                ".png",
                ".bmp",
                ".jpg",
                ".jpeg"
            ],
            width=8
        )

        self.type_combo.grid(
            row=0,
            column=2,
            padx=5
        )

        self.type_combo.set(".png")

        ttk.Button(
            image_frame,
            text="Capture",
            command=self.capture
        ).grid(
            row=0,
            column=3,
            padx=5
        )

        # =====================================================
        # SAVE LOCATION ROW
        # =====================================================

        save_frame = ttk.Frame(group1)
        save_frame.grid(
            row=2,
            column=0,
            padx=10,
            pady=5,
            sticky="ew"
        )

        save_frame.columnconfigure(1, weight=1)

        ttk.Label(
            save_frame,
            text="Save Location:"
        ).grid(
            row=0,
            column=0,
            sticky="w"
        )

        self.save_entry = ttk.Entry(save_frame)
        self.save_entry.grid(
            row=0,
            column=1,
            padx=5,
            sticky="ew"
        )

        self.save_entry.insert(0, self.get_default_save_location())

        ttk.Button(
            save_frame,
            text="Browse",
            command=self.browse_save_location
        ).grid(
            row=0,
            column=2,
            padx=5
        )

        # =====================================================
        # PREVIEW GROUP
        # =====================================================

        preview = ttk.LabelFrame(
            self,
            text="Previewer"
        )

        preview.grid(
            row=0,
            column=1,
            padx=(0, 10),
            pady=10,
            sticky="nsew"
        )

        preview.columnconfigure(0, weight=1)

        self.preview_label = tk.Label(
            preview,
            bg="white",
            relief="sunken",
            text="No boundary set",
            cursor="hand2"
        )

        self.preview_label.grid(
            row=0,
            column=0,
            padx=10,
            pady=10
        )

        # Click the preview to open a larger view of the bounded region.
        self.preview_label.bind("<Button-1>", self.open_large_preview)

        ttk.Label(
            preview,
            text="Click image for larger preview"
        ).grid(
            row=1,
            column=0,
            pady=(0, 10)
        )
    
    @contextmanager
    def _hidden(self):
        # Temporarily hide the Snip window so it isn't captured on top of the
        # map. A brief pause lets the windows underneath repaint before we grab.
        # Restore only to the previous state, so a headless (already withdrawn)
        # window is never revealed.
        was_visible = self.state() == "normal"
        self.withdraw()
        self.update()
        time.sleep(0.2)
        try:
            yield
        finally:
            if was_visible:
                self.deiconify()

    def update_preview(self):
        # Show the screen region defined by the current boundaries in the
        # Previewer label, hiding this window so it isn't captured.
        with self._hidden():
            self._render_preview()

    def _render_preview(self):
        # Render the current boundary region into the preview label, scaled to
        # fit. Assumes any needed window hiding is handled by the caller.
        left = int(self.left_bound)
        top = int(self.top_bound)
        right = int(self.right_bound)
        bottom = int(self.bottom_bound)

        width = right - left
        height = bottom - top

        if width > 0 and height > 0:
            try:
                shot = pyautogui.screenshot(
                    region=(left, top, width, height)
                )

                # Scale down to fit the preview box, preserving aspect ratio.
                scale = min(self.PREVIEW_W / width, self.PREVIEW_H / height)
                new_size = (
                    max(1, int(width * scale)),
                    max(1, int(height * scale))
                )
                shot = shot.resize(new_size, Image.LANCZOS)

                self.preview_image = ImageTk.PhotoImage(shot)
                self.preview_label.config(image=self.preview_image, text="")
            except Exception as e:
                self.preview_label.config(image="", text="Preview error")
                print(f"Preview error: {e}")
        else:
            self.preview_label.config(image="", text="No boundary set")

    def open_large_preview(self, event=None):
        # Capture the bounded region fresh and show it in a large popup window.
        left = int(self.left_bound)
        top = int(self.top_bound)
        right = int(self.right_bound)
        bottom = int(self.bottom_bound)

        width = right - left
        height = bottom - top

        if width <= 0 or height <= 0:
            messagebox.showinfo(
                "Preview",
                "No boundary set yet."
            )
            return

        with self._hidden():
            shot = pyautogui.screenshot(region=(left, top, width, height))

        # Scale to fit within most of the screen, preserving aspect ratio, but
        # never upscale past the captured resolution.
        max_w = int(self.winfo_screenwidth() * 0.9)
        max_h = int(self.winfo_screenheight() * 0.9)
        scale = min(max_w / width, max_h / height, 1)
        disp_size = (max(1, int(width * scale)), max(1, int(height * scale)))
        shot = shot.resize(disp_size, Image.LANCZOS)

        popup = tk.Toplevel(self)
        popup.title("Preview")
        popup.transient(self)

        # Keep a reference on the popup so the image isn't garbage collected.
        popup.image = ImageTk.PhotoImage(shot)

        label = tk.Label(popup, image=popup.image, bg="black")
        label.pack()

        # Click anywhere or press Escape to dismiss.
        label.bind("<Button-1>", lambda e: popup.destroy())
        popup.bind("<Escape>", lambda e: popup.destroy())

    def preview_boundary(self):
        # Show a light-grey translucent box on the screen over the current
        # boundary so the user can see exactly what will be captured.
        left = int(self.left_bound)
        top = int(self.top_bound)
        right = int(self.right_bound)
        bottom = int(self.bottom_bound)

        width = right - left
        height = bottom - top

        if width <= 0 or height <= 0:
            messagebox.showinfo(
                "Preview",
                "No boundary set yet."
            )
            return

        # Borderless, always-on-top, semi-transparent window sized to the bounds.
        overlay = tk.Toplevel(self)
        overlay.overrideredirect(True)
        overlay.geometry(f"{width}x{height}+{left}+{top}")
        overlay.configure(bg="#d3d3d3")
        overlay.attributes("-alpha", 0.35)
        overlay.attributes("-topmost", True)

        # Dismiss on click, Escape, or automatically after a few seconds.
        overlay.bind("<Button-1>", lambda e: overlay.destroy())
        overlay.bind("<Escape>", lambda e: overlay.destroy())
        overlay.focus_set()
        overlay.after(3000, overlay.destroy)

    def capture(self):
        # Save a screenshot of the current boundary region to the save location,
        # named from the Image Name + type comboboxes.
        left = int(self.left_bound)
        top = int(self.top_bound)
        right = int(self.right_bound)
        bottom = int(self.bottom_bound)

        width = right - left
        height = bottom - top

        if width <= 0 or height <= 0:
            messagebox.showinfo(
                "Capture",
                "No boundary set yet."
            )
            return

        save_dir = self.save_entry.get().strip()
        name = self.name_combo.get().strip()
        ext = self.type_combo.get().strip()

        if not save_dir:
            messagebox.showerror("Capture", "No save location set.")
            return
        if not name:
            messagebox.showerror("Capture", "No image name set.")
            return

        if not ext.startswith("."):
            ext = "." + ext

        path = os.path.join(save_dir, name + ext)

        # Confirm before clobbering an existing file.
        if os.path.exists(path):
            if not messagebox.askyesno(
                "Capture",
                f"{name + ext} already exists.\nOverwrite it?"
            ):
                return

        if self._capture_region(left, top, right, bottom, path):
            self.status_label.config(text="Status: Captured")

    def auto_capture(self):
        # Startup auto mode: capture the boundary loaded from settings to a
        # numbered "screenshot N" file in the current save location, without
        # any prompts.
        left = int(self.left_bound)
        top = int(self.top_bound)
        right = int(self.right_bound)
        bottom = int(self.bottom_bound)

        if right - left <= 0 or bottom - top <= 0:
            messagebox.showerror(
                "Capture",
                "No boundary is saved in settings to auto-capture."
            )
            return

        # In headless auto mode the GUI widgets don't exist, so fall back to the
        # default save location / .png.
        if hasattr(self, "save_entry"):
            save_dir = self.save_entry.get().strip()
        else:
            save_dir = self.get_default_save_location()

        if not save_dir:
            messagebox.showerror("Capture", "No save location set.")
            return

        if hasattr(self, "type_combo"):
            ext = self.type_combo.get().strip() or ".png"
        else:
            ext = ".png"
        if not ext.startswith("."):
            ext = "." + ext

        # Find the next unused "screenshot N" name so we never overwrite.
        n = 1
        while os.path.exists(os.path.join(save_dir, f"screenshot {n}{ext}")):
            n += 1
        path = os.path.join(save_dir, f"screenshot {n}{ext}")

        if self._capture_region(left, top, right, bottom, path):
            if hasattr(self, "status_label"):
                self.status_label.config(text=f"Status: Captured screenshot {n}")

    def _capture_region(self, left, top, right, bottom, path):
        # Screenshot the boundary region and save it to path. Returns True on
        # success. Hides the Snip window and parks the EDX crosshair first.
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)

            with self._hidden():
                # Park the EDX crosshair in the map corner furthest from the
                # boundary so it isn't captured in the screenshot.
                self._move_crosshair_out(left, top, right, bottom)

                shot = pyautogui.screenshot(
                    region=(left, top, right - left, bottom - top)
                )
            shot.save(path)
            return True
        except Exception as e:
            messagebox.showerror(
                "Capture",
                f"Could not save screenshot:\n{e}"
            )
            return False

    def _move_crosshair_out(self, left, top, right, bottom):
        # EDX draws a crosshair at the mouse position over the map, which would
        # otherwise show up in the screenshot. Click the corner of the EDX map
        # furthest from the boundary being captured so the crosshair lands well
        # outside the region.
        edx = self.get_window("SignalPro", exclude="Legend")
        if not edx:
            return

        # Corner points near the edges of the EDX map area, inset to stay off
        # the toolbar (top), tree panel (left) and status bar (bottom).
        left_inset, right_inset = 270, 15
        top_inset, bottom_inset = 94, 40
        corners = [
            (edx.left + left_inset, edx.top + top_inset),
            (edx.left + edx.width - right_inset, edx.top + top_inset),
            (edx.left + left_inset, edx.top + edx.height - bottom_inset),
            (edx.left + edx.width - right_inset,
             edx.top + edx.height - bottom_inset),
        ]

        # Pick the corner furthest from the center of the boundary region.
        cx = (left + right) / 2
        cy = (top + bottom) / 2
        x, y = max(corners, key=lambda p: (p[0] - cx) ** 2 + (p[1] - cy) ** 2)

        pyautogui.click(x, y)
        time.sleep(0.15)

    def browse_save_location(self):
        # Let the user pick a directory and overwrite the save location entry.
        current = self.save_entry.get()
        initial = current if os.path.isdir(current) else None

        chosen = filedialog.askdirectory(
            parent=self,
            title="Select Save Location",
            initialdir=initial
        )

        if chosen:
            self.save_entry.delete(0, tk.END)
            self.save_entry.insert(0, os.path.normpath(chosen))

    def _set_bounds(self, left, top, right, bottom, source="Manual",
                    refresh=True):
        # Store new boundary coordinates and refresh the labels + preview.
        # source names how the bounds were set (e.g. "Manual", "Auto") and
        # drives the status labels. Pass refresh=False when the caller will
        # render the preview itself (e.g. while already hidden).
        self.left_bound = int(left)
        self.top_bound = int(top)
        self.right_bound = int(right)
        self.bottom_bound = int(bottom)

        self.labelx1.config(text=f"x1: {self.left_bound}")
        self.labely1.config(text=f"y1: {self.top_bound}")
        self.labelx2.config(text=f"x2: {self.right_bound}")
        self.labely2.config(text=f"y2: {self.bottom_bound}")

        self.scale_status_label.config(text="Scale/Pos Status: Set")
        self.status_label.config(text=f"Status: {source} Set")

        # Persist so the boundary survives relaunches of this script for the
        # rest of the current Main GUI session.
        self.save_session_bounds()

        if refresh:
            self.update_preview()

    @staticmethod
    def session_bounds_path():
        # Per-user temp file holding the boundary set during the current Main
        # GUI session. tempfile.gettempdir() is already user-specific on
        # Windows, so a fixed name is enough. The Main GUI deletes this file on
        # startup, so each new session begins with the settings default.
        return os.path.join(tempfile.gettempdir(), "edx_snip_bounds.txt")

    def load_session_bounds(self):
        # Return (left, top, right, bottom) saved earlier this session, or None
        # if there's no valid session file yet.
        path = self.session_bounds_path()
        if not os.path.exists(path):
            return None
        try:
            with open(path, "r", encoding="utf-8") as f:
                parts = f.read().strip().split(",")
            left, top, right, bottom = (int(p) for p in parts)
            return left, top, right, bottom
        except Exception:
            return None

    def save_session_bounds(self):
        # Write the current boundary to the session temp file.
        try:
            with open(self.session_bounds_path(), "w", encoding="utf-8") as f:
                f.write(
                    f"{self.left_bound},{self.top_bound},"
                    f"{self.right_bound},{self.bottom_bound}"
                )
        except Exception as e:
            print(f"Could not save session bounds: {e}")

    def manual_bounds(self):
        # Let the user drag a rectangle over the screen to set the boundary.
        # The whole screen is dimmed grey; the true colors show through inside
        # the selection as it's dragged.
        with self._hidden():
            full = pyautogui.screenshot().convert("RGB")

        # Dimmed (greyed-out) version shown everywhere by default.
        grey = Image.new("RGB", full.size, (128, 128, 128))
        dimmed = Image.blend(full, grey, 0.6)

        overlay = tk.Toplevel(self)
        overlay.overrideredirect(True)
        overlay.geometry(f"{full.width}x{full.height}+0+0")
        overlay.attributes("-topmost", True)

        canvas = tk.Canvas(overlay, highlightthickness=0, cursor="cross")
        canvas.pack(fill="both", expand=True)

        # Background = dimmed screen. Keep references so images aren't GC'd.
        canvas.dim_tk = ImageTk.PhotoImage(dimmed)
        canvas.create_image(0, 0, anchor="nw", image=canvas.dim_tk)
        canvas.full = full
        canvas.reveal_tk = None
        canvas.reveal_id = None
        canvas.rect_id = None
        canvas.start = None

        def on_press(event):
            canvas.start = (event.x, event.y)

        def on_motion(event):
            if not canvas.start:
                return

            x1, y1 = canvas.start
            l, t = min(x1, event.x), min(y1, event.y)
            r, b = max(x1, event.x), max(y1, event.y)

            if r - l < 1 or b - t < 1:
                return

            # Reveal the true (un-dimmed) colors inside the selection.
            crop = canvas.full.crop((l, t, r, b))
            canvas.reveal_tk = ImageTk.PhotoImage(crop)

            if canvas.reveal_id is None:
                canvas.reveal_id = canvas.create_image(
                    l, t, anchor="nw", image=canvas.reveal_tk
                )
            else:
                canvas.coords(canvas.reveal_id, l, t)
                canvas.itemconfig(canvas.reveal_id, image=canvas.reveal_tk)

            # Selection outline drawn on top.
            if canvas.rect_id is None:
                canvas.rect_id = canvas.create_rectangle(
                    l, t, r, b, outline="red", width=2
                )
            else:
                canvas.coords(canvas.rect_id, l, t, r, b)
            canvas.tag_raise(canvas.rect_id)

        def on_release(event):
            start = canvas.start
            overlay.destroy()

            if not start:
                return

            x1, y1 = start
            l, t = min(x1, event.x), min(y1, event.y)
            r, b = max(x1, event.x), max(y1, event.y)

            # Ignore stray clicks that didn't drag a real region.
            if r - l < 2 or b - t < 2:
                return

            self._set_bounds(l, t, r, b)

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_motion)
        canvas.bind("<ButtonRelease-1>", on_release)
        overlay.bind("<Escape>", lambda e: overlay.destroy())
        overlay.focus_set()

    def on_close(self):
        # Tear down the whole app when the window is closed. Destroy the root
        # (which ends mainloop); fall back to self if there's no parent.
        target = self.parent if self.parent is not None else self
        target.destroy()
    
    def settings_file_path(self):
        return fr"\\TetonRiver\prop\template\zAutoit\User Settings\{self.username} settings.txt"
   
    def get_user_settings(self):
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

    def _trim_to_content(self, left, top, right, bottom,
                         sat_level=30, fill_frac=0.4):
        # Screenshot the given screen rectangle and return the bounding box of
        # the solid colored map block, in screen coordinates, or None if none.
        #
        # The map is a solid rectangle of vivid color (the terrain fills it edge
        # to edge), whereas the window chrome is only sparsely colored: toolbar
        # icons, a legend swatch, tree-panel highlights. So rather than bounding
        # every colored pixel (which stray chrome would stretch), we keep only
        # the rows and columns that are *mostly* colored -- that isolates the
        # map block and ignores the sparse chrome and white margins.
        left, top, right, bottom = int(left), int(top), int(right), int(bottom)
        width = right - left
        height = bottom - top

        if width <= 0 or height <= 0:
            return None

        # Boolean mask of colored pixels via HSV saturation (0 = gray/white).
        sat = pyautogui.screenshot(
            region=(left, top, width, height)
        ).convert("HSV").getchannel("S")
        mask = np.asarray(sat) > sat_level

        if not mask.any():
            return None

        # Fraction of each row/column that is colored.
        row_frac = mask.mean(axis=1)
        col_frac = mask.mean(axis=0)

        rows = np.where(row_frac > fill_frac)[0]
        cols = np.where(col_frac > fill_frac)[0]

        if rows.size == 0 or cols.size == 0:
            return None

        bx1, bx2 = int(cols[0]), int(cols[-1]) + 1
        by1, by2 = int(rows[0]), int(rows[-1]) + 1

        return (left + bx1, top + by1, left + bx2, top + by2)

    def edx_bounds(self):

        # The legend window's title also contains "SignalPro", so exclude it to
        # avoid grabbing the legend as the main map window.
        edx = self.get_window("SignalPro", exclude="Legend")

        if not edx:
            messagebox.showerror(
                "Error",
                "Could not find EDX SignalPro window."
            )
            return

        #
        # Main EDX window
        #
        edx_left = edx.left
        edx_top = edx.top
        edx_right = edx.left + edx.width
        edx_bottom = edx.top + edx.height


        # Hide the Snip window for the whole operation so it isn't captured on
        # top of the map by the content scan or the preview render.
        with self._hidden():
            #
            # Scan the whole EDX window. The projection-based trim isolates the
            # solid colored map block on its own, dropping the toolbar, tree
            # panel, status bar and white margins -- no manual offsets needed.
            # The legend floats over the map, so it falls inside the box too.
            #
            box = self._trim_to_content(
                edx_left,
                edx_top,
                edx_right,
                edx_bottom
            )

            if box:
                # Update coords + status labels. refresh=False because we're
                # already hidden and render the preview directly below.
                self._set_bounds(*box, source="Auto", refresh=False)

                # Boundary changed -> refresh the preview (already hidden).
                self._render_preview()

        if not box:
            messagebox.showerror(
                "Error",
                "Could not find any colored map content to bound."
            )
            return

        # print(
        #     f"
        # Bounds: "
        #     f"{self.left_bound}, "
        #     f"{self.top_bound}, "
        #     f"{self.right_bound}, "
        #     f"{self.bottom_bound}"
        # )   

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
    
    def record(self):
        try:
            self.wait_for_window(".*initial view.*")
        except:
            raise "Make sure an EDX Project is open."

        send_keys("%d")
        send_keys("{RIGHT 1}")
        send_keys("{ENTER}")

        app = Application(backend="win32").connect(title="Current View Setup")
        dlg = app.window(title="Current View Setup")

        # Latitude
        self.latitude = dlg["Edit9"].window_text()
        
        # Longitude
        self.longitude = dlg["Edit10"].window_text()

        dlg.child_window(title="&OK", class_name="Button").click()

    def recall(self):
        try:
            self.wait_for_window(".*initial view.*")
        except:
            raise "Make sure an EDX Project is open."

        send_keys("%d")
        send_keys("{RIGHT 1}")
        send_keys("{ENTER}")

        app = Application(backend="win32").connect(title="Current View Setup")
        dlg = app.window(title="Current View Setup")

        lat_edit = dlg.child_window(best_match="Latitude:Edit")
        lon_edit = dlg.child_window(best_match="Longitude:Edit")

        lat_edit.set_edit_text(self.latitude)
        lon_edit.set_edit_text(self.longitude)

        dlg.child_window(title="&OK", class_name="Button").click()

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    SnipWindow(username=os.getlogin(), parent=root)

    root.mainloop()