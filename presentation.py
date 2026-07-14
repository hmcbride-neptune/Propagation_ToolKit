import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import os
import sys
import psutil
import win32com.client
import pandas as pd

class PresentationApp:
    def __init__(self, master):
        self.project = sys.argv[1] if len(sys.argv) > 1 else None

        self.master = master
        self.master.title("Map Indicator")
        self.master.geometry("593x289")
        self.master.configure(bg="#B9D1EA")

        self.Label = tk.Label(self.master, text="CSV Location", font=("MS Sans Serif", 12, "underline"))
        self.Label.place(x=280, y=8, width=102, height=24)
        self.Label.configure(bg="#B9D1EA")
        
        self.Label = tk.Label(self.master, text="Map", font=("MS Sans Serif", 12,"underline"))
        self.Label.place(x=536, y=8, width=35, height=24)
        self.Label.configure(bg="#B9D1EA")

        self.Label = tk.Label(self.master, text="Browse to Master Group:", font=("MS Sans Serif", 10))
        self.Label.place(x=48, y=40, width=150, height=20)
        self.Label.configure(bg="#B9D1EA")

        self.Label = tk.Label(self.master, text="Browse to intermediate Group:", font=("MS Sans Serif", 10))
        self.Label.place(x=16, y=72, width=183, height=20)
        self.Label.configure(bg="#B9D1EA")

        self.Label = tk.Label(self.master, text="Browse to Suggested Group:", font=("MS Sans Serif", 10))
        self.Label.place(x=22, y=104, width=178, height=20)
        self.Label.configure(bg="#B9D1EA")

        self.Label = tk.Label(self.master, text="Browse to 4th Group (If needed):", font=("MS Sans Serif", 10))
        self.Label.place(x=5, y=136, width=190, height=20)
        self.Label.configure(bg="#B9D1EA")

        self.Label = tk.Label(self.master, text="Browse to 5th Group (If needed):", font=("MS Sans Serif", 10))
        self.Label.place(x=5, y=168, width=190, height=20)
        self.Label.configure(bg="#B9D1EA")

        self.Label = tk.Label(self.master, text="Browse to 6th Group (If needed):", font=("MS Sans Serif", 10))
        self.Label.place(x=5, y=200, width=190, height=20)
        self.Label.configure(bg="#B9D1EA")

        self.run_button = tk.Button(self.master, text="Run", command=self.run_script, font=("MS Sans Serif", 16))
        self.run_button.place(x=8, y=224, width=209, height=33)
        self.run_button.configure(bg="#FFFFFF")

        self.progress = ttk.Progressbar(self.master, orient="horizontal", length=577, mode="determinate")
        self.progress.place(x=8, y=264, width=577, height=17)

        self.checking_label = tk.Label(self.master, text="Checking Line: -/-", font=("Palatino Linotype", 12), anchor="w")
        self.checking_label.place(x=224, y=232, width=231, height=26)
        self.checking_label.configure(bg="#B9D1EA")
        
        self.mapNo1 = tk.Label(self.master, text="1", font=("Rockwell", 15, "underline"), anchor="w")
        self.mapNo1.place(x=545, y=40, width=41, height=21)
        self.mapNo1.configure(bg="#B9D1EA")

        self.mapNo2 = tk.Label(self.master, text="2", font=("Rockwell", 15, "underline"), anchor="w")
        self.mapNo2.place(x=545, y=72, width=41, height=21)
        self.mapNo2.configure(bg="#B9D1EA")

        self.mapNo3 = tk.Label(self.master, text="3", font=("Rockwell", 15, "underline"), anchor="w")
        self.mapNo3.place(x=545, y=104, width=41, height=21)
        self.mapNo3.configure(bg="#B9D1EA")

        self.mapNo4 = tk.Label(self.master, text="4", font=("Rockwell", 15, "underline"), anchor="w")
        self.mapNo4.place(x=545, y=136, width=41, height=21)
        self.mapNo4.configure(bg="#B9D1EA")

        self.mapNo5 = tk.Label(self.master, text="5", font=("Rockwell", 15, "underline"), anchor="w")
        self.mapNo5.place(x=545, y=168, width=41, height=21)
        self.mapNo5.configure(bg="#B9D1EA")

        self.mapNo6 = tk.Label(self.master, text="6", font=("Rockwell", 15, "underline"), anchor="w")
        self.mapNo6.place(x=545, y=200, width=41, height=21)
        self.mapNo6.configure(bg="#B9D1EA")

        self.entries = []

        for i in range(6):
            y = 40 + (i * 32)

            entry = tk.Entry(self.master, state="readonly")
            entry.place(x=200, y=y, width=257, height=21)

            self.entries.append(entry)

            tk.Button(
                self.master,
                text="Browse",
                command=lambda e=entry: self.browse_file(e)
            ).place(x=472, y=y, width=49, height=25)

        if self.project:
            master_path = rf"C:\ProgramData\EDX\Projects\{self.project}\Data\xPropInfo\GateWay Locations - Master.csv"
            self.entries[0].config(state="normal")
            self.entries[0].delete(0, tk.END)
            if os.path.isfile(master_path):
                self.entries[0].insert(0, master_path)
            self.entries[0].config(state="readonly")

    def browse_file(self, entry):
        filename = filedialog.askopenfilename()

        if filename:
            entry.config(state="normal")
            entry.delete(0, tk.END)
            entry.insert(0, filename)
            entry.config(state="readonly")

    def combine_files(self):
        """Read every imported CSV, tag its rows with the map number (1-6),
        and combine them into a single DataFrame."""
        frames = []
        for i, entry in enumerate(self.entries):
            path = entry.get().strip()
            if not path:
                continue
            try:
                df = pd.read_csv(path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read {path}: {e}")
                return None
            df.insert(0, "Map", i + 1)
            frames.append(df)

        if not frames:
            messagebox.showwarning("No Files", "No CSV files were selected.")
            return None

        combined = pd.concat(frames, ignore_index=True)

        # Where a location exists as both a "3" and "4" variant (e.g. 3Site /
        # 4Site), keep only the "4" version.
        locations = set(combined["Sector"].astype(str))

        def keep_location(loc):
            loc = str(loc)
            return not (loc.startswith("3") and ("4" + loc[1:]) in locations)

        combined = combined[combined["Sector"].apply(keep_location)].reset_index(drop=True)
        return combined

    def build_output(self, df):
        """Transform the combined DataFrame into the collector output layout,
        deriving Collector/Antenna and converting metres to feet."""
        M_TO_FT = 3.28084

        def is_close(value, target):
            try:
                return abs(float(value) - target) < 1e-6
            except (TypeError, ValueError):
                return False

        def collector_antenna(row):
            tpo = row.get("CollectorTPO")
            gain = row.get("AntGain")
            if is_close(tpo, 0):
                return "Tmega", "2xMFB9155"
            if is_close(tpo, -1):
                if is_close(gain, 7.1):
                    return "GPV4", "MFB9155"
                if is_close(gain, 11.1):
                    return "GPV4", "DB589-Y"
                return "GPV4", ""
            if is_close(tpo, 9.8):
                return "450C", "DB636-C"
            if is_close(tpo, 6.5):
                return "450M", "BS450XL3_C"
            return "", ""

        def to_ft(value):
            return pd.to_numeric(value, errors="coerce") * M_TO_FT

        def whole(value):
            return pd.to_numeric(value, errors="coerce").round(0)

        collector = df.apply(lambda r: collector_antenna(r)[0], axis=1)
        antenna = df.apply(lambda r: collector_antenna(r)[1], axis=1)

        output = pd.DataFrame({
            "Map": df["Map"],
            "Location": df["Sector"],
            "Latitude": df["Latitude"],
            "Longitude": df["Longitude"],
            "Collector": collector,
            "Elev(m)": whole(df["Elev(m)"]),
            "AntHgt(m)": whole(df["AntHgt(m)"]),
            "Elev(ft)": to_ft(df["Elev(m)"]).round(0),
            "AntHgt(ft)": to_ft(df["AntHgt(m)"]).round(0),
            "Coax Type": df["Transmission Line Type"],
            "Coax(ft)": to_ft(df["Transmission Line Length (m)"]).round(0),
            "Antenna": antenna,
        })

        # Collapse duplicate collectors (identical in every output column except
        # Map) into a single row whose Map lists each map it appears on, e.g. "1,2".
        data_cols = [c for c in output.columns if c != "Map"]
        output = (
            output.groupby(data_cols, dropna=False, sort=False)["Map"]
            .apply(lambda maps: ",".join(str(m) for m in sorted(maps.unique())))
            .reset_index()
        )
        output = output[["Map"] + data_cols]

        # Sort: multi-map rows first (1,2 before 1,3), single-map rows at the
        # bottom, then by Location. Zero-pad each map number so the string sort
        # matches numeric order (e.g. "001,002" < "001,010").
        map_str = output["Map"].astype(str)
        output["_single"] = map_str.apply(lambda s: 1 if "," not in s else 0)
        output["_mapkey"] = map_str.apply(
            lambda s: ",".join(f"{int(p):03d}" for p in s.split(","))
        )
        output = (
            output.sort_values(by=["_single", "_mapkey", "Location"])
            .drop(columns=["_single", "_mapkey"])
            .reset_index(drop=True)
        )
        return output

    def is_powerpoint_open(self):
        for process in psutil.process_iter(['name']):
            try:
                if process.info['name'] and process.info['name'].upper() == 'POWERPNT.EXE':
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        return False
    
    def find_gateway_slide(self):
        search_text = "Gateway Locations:"
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt.ActivePresentation

        for slide in presentation.Slides:
            for shape in slide.Shapes:
                try:
                    if shape.HasTextFrame and search_text in shape.TextFrame.TextRange.Text:
                        return slide
                except:
                    pass
        return None

    def get_table(self, slide):
        """Return the first table shape on a slide, or None."""
        for shape in slide.Shapes:
            try:
                if shape.HasTable:
                    return shape.Table
            except:
                pass
        return None

    def _cell_text(self, value):
        """Format a DataFrame value for a PowerPoint table cell."""
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except (TypeError, ValueError):
            pass
        if isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return ("%.6f" % value).rstrip("0").rstrip(".")
        return str(value)

    def update_progress(self, current, total):
        """Refresh the progress bar and the 'Checking Line' label."""
        self.progress["maximum"] = total if total else 1
        self.progress["value"] = current
        self.checking_label.config(text=f"Checking Line: {current}/{total}")
        self.master.update_idletasks()

    def fill_table(self, table, rows, done=0, total=None):
        """Write a chunk of rows into the table's data rows (header is row 1).
        `done` is how many rows were written on previous slides and `total` is
        the overall row count, so the progress bar tracks across all slides.
        Returns the new running total of written rows."""
        if total is None:
            total = len(rows)
        for i, row in enumerate(rows):
            for c, value in enumerate(row):
                table.Cell(i + 2, c + 1).Shape.TextFrame.TextRange.Text = self._cell_text(value)
            self.update_progress(done + i + 1, total)
        return done + len(rows)

    def trim_table(self, table, used_rows):
        """Delete any data rows past the ones we filled (header + used_rows kept).
        Deletes bottom-up so row indices stay valid."""
        for r in range(table.Rows.Count, used_rows + 1, -1):
            table.Rows(r).Delete()

    def populate_collector_table(self):
        slide = self.find_gateway_slide()
        if slide is None:
            messagebox.showerror("PowerPoint", "Could not find the 'Gateway Locations:' slide.")
            return

        table = self.get_table(slide)
        if table is None:
            messagebox.showerror("PowerPoint", "Could not find a table on the Gateway Locations slide.")
            return

        # Match the DataFrame column order to the table columns.
        columns = [
            "Map", "Location", "Latitude", "Longitude", "Collector",
            "Elev(m)", "AntHgt(m)", "Elev(ft)", "AntHgt(ft)",
            "Coax Type", "Coax(ft)", "Antenna",
        ]
        records = self.output_df[columns].values.tolist()

        # Data rows available per table (row 1 is the header).
        capacity = table.Rows.Count - 1
        chunks = [records[i:i + capacity] for i in range(0, len(records), capacity)] or [[]]

        total = len(records)
        self.update_progress(0, total)

        # First chunk fills the existing slide; each extra chunk gets a duplicate.
        done = self.fill_table(table, chunks[0], done=0, total=total)
        self.trim_table(table, len(chunks[0]))

        prev_slide = slide
        for chunk in chunks[1:]:
            prev_slide = prev_slide.Duplicate().Item(1)
            new_table = self.get_table(prev_slide)
            done = self.fill_table(new_table, chunk, done=done, total=total)
            self.trim_table(new_table, len(chunk))

    def run_script(self):
        self.combined_df = self.combine_files()
        if self.combined_df is None:
            return

        self.output_df = self.build_output(self.combined_df)

        if self.is_powerpoint_open():
            self.populate_collector_table()
            self.master.destroy()
        else:
            messagebox.showinfo("PowerPoint", "PowerPoint is not open")


if __name__ == "__main__":
    root = tk.Tk()
    app = PresentationApp(root)
    root.mainloop()