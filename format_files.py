import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import sys
import os
import shutil
from datetime import datetime
import pandas as pd
import csv
from pathlib import Path
from win32com.client import DispatchEx
import pythoncom


class FileFormat:
    def __init__(self, root):
        self.root = root
        self.root.title("Select file to format")
        self.root.geometry("380x320")

        # Center everything
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(root, padding=10)
        main_frame.grid(row=0, column=0, sticky="nsew")

        main_frame.columnconfigure(0, weight=1)

        # Style for the radio button text
        style = ttk.Style()
        style.configure("Radio.TRadiobutton", font=("MS Sans Serif", 10))

        # ----------------------------
        # Top Row
        # ----------------------------
        top_frame = ttk.Frame(main_frame)
        top_frame.grid(row=0, column=0, pady=(10, 15))

        ttk.Label(
            top_frame,
            text="Files in:",
            font=("MS Sans Serif", 10)
        ).grid(row=0, column=0, padx=(0, 10))

        self.folder_var = tk.StringVar()

        ttk.Radiobutton(
            top_frame,
            text="Requests Folder",
            variable=self.folder_var,
            value="requests",
            style="Radio.TRadiobutton"
        ).grid(row=0, column=1, padx=10)

        ttk.Radiobutton(
            top_frame,
            text="Projects Folder",
            variable=self.folder_var,
            value="projects",
            style="Radio.TRadiobutton"
        ).grid(row=0, column=2, padx=10)

        # Default selection
        self.folder_var.set("requests")

        # ----------------------------
        # Button Area
        # ----------------------------
        border = tk.Frame(
            main_frame,
            highlightbackground="grey",
            highlightcolor="grey",
            highlightthickness=1,
            padx=100,
            pady=20
        )
        border.grid(row=1, column=0, sticky="ew")

        ttk.Button(
            border,
            text="Export_Output.csv",
            width=28,
            command=self.btnExportOutput_click
        ).grid(row=0, column=0, pady=(0, 12))

        ttk.Button(
            border,
            text="Export_Output.csv (StreetName)",
            width=28,
            command=self.btnExportOutput_click
        ).grid(row=1, column=0, pady=(0, 35))

        ttk.Button(
            border,
            text="meters.csv (ESRI)",
            width=28,
            command=self.meters_click
        ).grid(row=2, column=0, pady=(0, 35))

        ttk.Button(
            border,
            text="GoogleGeo Output",
            width=28,
            command=self.btnGoogleGeo_click
        ).grid(row=3, column=0)

        self.project_name = sys.argv[1] if len(sys.argv) > 1 else "No project name provided"
    
    def get_selected_folder(self):
        return self.folder_var.get()
    
    def btnExportOutput_click(self):
        active_folder = self.get_selected_folder()
        if active_folder == "requests":
            ESRIOutputFile = filedialog.askopenfilename(title="Select ESRI Export_Output csv file.",filetypes=[("CSV files", "*.csv")], initialdir=r"Z:\Requests")
        elif active_folder == "projects":
            if self.project_name == "No project name provided":
                ESRIOutputFile = filedialog.askopenfilename(title="Select ESRI Export_Output csv file.",filetypes=[("CSV files", "*.csv")], initialdir=r"C:\ProgramData\EDX\Projects")
            else:
                ESRIOutputFile = filedialog.askopenfilename(title="Select ESRI Export_Output csv file.",filetypes=[("CSV files", "*.csv")], initialdir=fr"C:\ProgramData\EDX\Projects\{self.project_name}")
        self.ESRIOutputFormat(ESRIOutputFile)

    def meters_click(self):
        selected_folder = self.get_selected_folder()
        if selected_folder == "requests":
            meters_file = filedialog.askopenfilename(title="Select meters.csv file.", filetypes=[("CSV files", "*.csv")], initialdir=r"Z:\Requests")
        elif selected_folder == "projects":
            if self.project_name == "No project name provided":
                meters_file = filedialog.askopenfilename(title="Select meters.csv file.", filetypes=[("CSV files", "*.csv")], initialdir=r"C:\ProgramData\EDX\Projects")
            else:
                meters_file = filedialog.askopenfilename(title="Select meters.csv file.", filetypes=[("CSV files", "*.csv")], initialdir=fr"C:\ProgramData\EDX\Projects\{self.project_name}")
        self.MeterCSVFormat(meters_file)

    def btnGoogleGeo_click(self):
        selected_folder = self.get_selected_folder()
        if selected_folder == "requests":
            google_geo_file = filedialog.askopenfilename(title="Select Google Earth Geocode file.", filetypes=[("KML files", "*.kml"), ("CSV files", "*.csv")], initialdir=r"Z:\Requests")
        elif selected_folder == "projects":
            if self.project_name == "No project name provided":
                google_geo_file = filedialog.askopenfilename(title="Select Google Earth Geocode file.", filetypes=[("KML files", "*.kml"), ("CSV files", "*.csv")], initialdir=r"C:\ProgramData\EDX\Projects")
            else:
                google_geo_file = filedialog.askopenfilename(title="Select Google Earth Geocode file.", filetypes=[("KML files", "*.kml"), ("CSV files", "*.csv")], initialdir=fr"C:\ProgramData\EDX\Projects\{self.project_name}")
        self.GoogleGeoFormat(google_geo_file)

    def ESRIOutputFormat(self, ESRIOutputFile):
        ESRIFolder = os.path.dirname(ESRIOutputFile)
        ESRIZeroStreets = 0

        result = self.esri_output_format(ESRIOutputFile, ESRIFolder, ESRIZeroStreets, show_message=True)   
    
    def esri_output_format(self, esri_output_file, esri_folder, esri_zero_streets=0, show_message=True):
        
        if not esri_output_file:
                return

        if not os.path.isfile(esri_output_file):
            raise FileNotFoundError(f"ESRI output file not found: {esri_output_file}")

        if not os.path.isdir(esri_folder):
            raise NotADirectoryError(f"ESRI folder not found: {esri_folder}")

        # AutoIt format was YYMMDD
        ymd = datetime.now().strftime("%y%m%d")

        backup_file = os.path.join(esri_folder, f"Export_OutputBackup{ymd}.csv")

        # Same as FileCopy()
        shutil.copy2(esri_output_file, backup_file)

        # Read CSV as strings so pandas does not alter IDs, leading zeros, etc.
        df = pd.read_csv(esri_output_file, dtype=str, keep_default_na=False)

        if df.empty:
            raise ValueError("ESRI output file is empty.")

        # AutoIt searched A1:BZ1 for Addr_type.
        # BZ is column 78, but pandas uses zero-based index.
        addr_type_col = self.find_column_by_name(
            df,
            search_name="Addr_type",
            start_index=0,
            end_index=min(78, len(df.columns)),
            partial=True
        )

        # AutoIt searched Z1:BZ1 for X.
        # Z is column 26 in Excel, zero-based index 25 in pandas.
        long_col = self.find_column_by_name(
            df,
            search_name="X",
            start_index=min(25, len(df.columns)),
            end_index=min(78, len(df.columns)),
            partial=True
        )

        long_col_index = list(df.columns).index(long_col)
        lat_col_index = long_col_index + 1

        if lat_col_index >= len(df.columns):
            raise ValueError(f"Latitude column not found after longitude column '{long_col}'.")

        lat_col = df.columns[lat_col_index]

        # Equivalent to your filter criteria
        if esri_zero_streets == 0:
            addr_types_to_zero = [
                "Locality",
                "Postal"
            ]
        else:
            addr_types_to_zero = [
                "Locality",
                "Postal",
                "StreetName",
                "StreetInt",
                "POI"
            ]

        # This replaces the Excel filter + range write behavior
        mask = df[addr_type_col].isin(addr_types_to_zero)

        if mask.any():
            df.loc[mask, [long_col, lat_col]] = "0"

        # Save back to the original CSV
        df.to_csv(esri_output_file, index=False)

        if show_message:
            self.show_auto_close_message(
                title="Done",
                message="ESRI Output File formatted!",
                duration_ms=1000
            )

        return {
            "backup_file": backup_file,
            "rows_updated": int(mask.sum()),
            "addr_type_column": addr_type_col,
            "longitude_column": long_col,
            "latitude_column": lat_col
        }

    def find_column_by_name(self, df, search_name, start_index=0, end_index=None, partial=True):
        columns = list(df.columns)

        if end_index is None:
            end_index = len(columns)

        search_range = columns[start_index:end_index]

        # Prefer exact match first
        for col in search_range:
            if str(col).strip().lower() == search_name.lower():
                return col

        # Then allow partial match if enabled
        if partial:
            for col in search_range:
                if search_name.lower() in str(col).strip().lower():
                    return col

        raise ValueError(f"Could not find column matching '{search_name}'.")

    def show_auto_close_message(self, title, message, duration_ms=1000):
        root = tk.Tk()
        root.withdraw()

        popup = tk.Toplevel(root)
        popup.title(title)
        popup.geometry("300x100")
        popup.resizable(False, False)

        label = tk.Label(popup, text=message, padx=20, pady=20)
        label.pack(expand=True)

        popup.after(duration_ms, root.destroy)
        root.mainloop()

    def MeterCSVFormat(self, meter_csv_file):

        df = pd.read_csv(
                meter_csv_file,
                header=None,
                dtype=str
            )

        # Keep columns after AX
        df = df.iloc[:, 50:]

        # Keep first two remaining columns
        df = df.iloc[:, :2]

        if "R900" in meter_csv_file.upper():
            rf_value = "-118.1"
        elif "R450" in meter_csv_file.upper():
            rf_value = "-113.1"
        else:
            raise ValueError("Cannot determine R450/R900")

        df.insert(2, "RSSI", rf_value)

        df.iloc[0, 0] = "5"
        df.iloc[0, 1] = ""

        df.to_csv(
            meter_csv_file,
            index=False,
            header=False
        )

    def GoogleGeoFormat(self, google_geo_file):

        result = self.google_geocode_format(
            google_geo_file=google_geo_file,

            google_geo_d1="HeaderD",
            google_geo_e1="HeaderE",
            google_geo_f1="HeaderF",
            google_geo_g1="HeaderG",
            google_geo_h1="HeaderH",

            google_geo_d2="=A2",
            google_geo_e2="=B2",
            google_geo_f2="=C2",
            google_geo_g2="",
            google_geo_h2="",

            show_message=True
        )

    def google_geocode_format(self,
        google_geo_file,
        google_geo_d1,
        google_geo_e1,
        google_geo_f1,
        google_geo_g1,
        google_geo_h1,
        google_geo_d2,
        google_geo_e2,
        google_geo_f2,
        google_geo_g2,
        google_geo_h2,
        show_message=True
    ):
        
        XL_CSV = 6
        XL_XML_LOAD_IMPORT_TO_LIST = 2
        XL_PASTE_FORMULAS = -4123
        XL_PASTE_VALUES_AND_NUMBER_FORMATS = 12
        XL_CELL_TYPE_VISIBLE = 12

        if not google_geo_file:
            return False

        google_geo_file = str(Path(google_geo_file))
        file_lower = google_geo_file.lower()

        excel = None
        workbook = None

        try:
            pythoncom.CoInitialize()

            excel = DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            google_geo_csv_file = ""

            # ------------------------------------------------------------
            # If input is KML, open it as XML and save as CSV
            # ------------------------------------------------------------
            if ".kml" in file_lower:
                google_geo_kml_file_name = google_geo_file[:-4]
                google_geo_csv_file = google_geo_kml_file_name + ".csv"

                workbook_kml = excel.Workbooks.OpenXML(
                    Filename=google_geo_file,
                    Stylesheets=None,
                    LoadOption=XL_XML_LOAD_IMPORT_TO_LIST
                )

                workbook_kml.SaveAs(
                    Filename=google_geo_csv_file,
                    FileFormat=XL_CSV
                )

                workbook_kml.Close(SaveChanges=False)

            # ------------------------------------------------------------
            # If input is already CSV
            # ------------------------------------------------------------
            elif ".csv" in file_lower:
                google_geo_csv_file = google_geo_file

            else:
                raise ValueError("Google geocode file must be a .kml or .csv file.")

            if not google_geo_csv_file:
                return False

            google_geo_csv_file_name = google_geo_csv_file[:-4]
            google_geo_csv_file_complete = google_geo_csv_file_name + "Completed.csv"

            # ------------------------------------------------------------
            # Open CSV
            # ------------------------------------------------------------
            workbook = excel.Workbooks.Open(google_geo_csv_file)
            worksheet = workbook.ActiveSheet

            # ------------------------------------------------------------
            # Find "#S_" in M1:AF30
            # ------------------------------------------------------------
            found_cell = self.find_text_in_range(
                worksheet=worksheet,
                text="#S_",
                range_address="M1:AF30"
            )

            found_col_num = found_cell.Column
            found_row_num = found_cell.Row
            found_col_letter = self.column_number_to_letter(found_col_num)

            # ------------------------------------------------------------
            # Delete columns A through found search column
            # Equivalent:
            # _Excel_RangeDelete($GoogleGeoWorkbook.activesheet,"A:" & $GoogleGeoSearchCol)
            # ------------------------------------------------------------
            worksheet.Range(f"A:{found_col_letter}").Delete()

            # ------------------------------------------------------------
            # Delete rows 1 through found_row - 2
            # Equivalent:
            # _Excel_RangeDelete(...,"1:" & $GoogleGeoSearchRow-2)
            # ------------------------------------------------------------
            rows_to_delete_end = found_row_num - 2

            if rows_to_delete_end >= 1:
                worksheet.Rows(f"1:{rows_to_delete_end}").Delete()

            # ------------------------------------------------------------
            # Write D1:H2 values/formulas
            # ------------------------------------------------------------
            worksheet.Range("D1").Value = google_geo_d1
            worksheet.Range("E1").Value = google_geo_e1
            worksheet.Range("F1").Value = google_geo_f1
            worksheet.Range("G1").Value = google_geo_g1
            worksheet.Range("H1").Value = google_geo_h1

            worksheet.Range("D2").Value = google_geo_d2
            worksheet.Range("E2").Value = google_geo_e2
            worksheet.Range("F2").Value = google_geo_f2
            worksheet.Range("G2").Value = google_geo_g2
            worksheet.Range("H2").Value = google_geo_h2

            # ------------------------------------------------------------
            # Save as completed CSV
            # ------------------------------------------------------------
            workbook.SaveAs(
                Filename=google_geo_csv_file_complete,
                FileFormat=XL_CSV
            )

            google_geo_csv_complete_num_lines = self.count_csv_lines(
                google_geo_csv_file_complete
            )

            # ------------------------------------------------------------
            # Copy formulas D2:H4 down to D5:Hlast
            # Equivalent:
            # _Excel_RangeCopyPaste(...,"D2:H4","D5:H" & line_count, formulas)
            # ------------------------------------------------------------
            if google_geo_csv_complete_num_lines >= 5:
                worksheet.Range("D2:H4").Copy()
                worksheet.Range(f"D5:H{google_geo_csv_complete_num_lines}").PasteSpecial(
                    Paste=XL_PASTE_FORMULAS
                )
                excel.CutCopyMode = False

            # ------------------------------------------------------------
            # Force Excel to calculate before pasting values
            # ------------------------------------------------------------
            excel.CalculateFull()

            # ------------------------------------------------------------
            # Paste D2:Hlast as values and number formats
            # Equivalent to your paste-values steps
            # ------------------------------------------------------------
            worksheet.Range(f"D2:H{google_geo_csv_complete_num_lines}").Copy()
            worksheet.Range(f"D2:H{google_geo_csv_complete_num_lines}").PasteSpecial(
                Paste=XL_PASTE_VALUES_AND_NUMBER_FORMATS
            )
            excel.CutCopyMode = False

            # ------------------------------------------------------------
            # Delete A:C
            # ------------------------------------------------------------
            worksheet.Range("A:C").Delete()

            # ------------------------------------------------------------
            # Filter A:E where column A is blank, then delete visible rows
            # Equivalent:
            # _Excel_FilterSet($GoogleGeoWorkbook,Default,"A:E",1,"=")
            # _Excel_RangeDelete(...,"A2:E" & line_count)
            # ------------------------------------------------------------
            worksheet.Range(f"A1:E{google_geo_csv_complete_num_lines}").AutoFilter(
                Field=1,
                Criteria1="="
            )

            try:
                visible_rows = worksheet.Range(
                    f"A2:E{google_geo_csv_complete_num_lines}"
                ).SpecialCells(XL_CELL_TYPE_VISIBLE)

                visible_rows.EntireRow.Delete()

            except pythoncom.com_error:
                # No visible rows matched the blank filter
                pass

            # Clear filter
            if worksheet.AutoFilterMode:
                worksheet.AutoFilterMode = False

            # ------------------------------------------------------------
            # Replace ",0" with blank in A:E
            # Equivalent:
            # _Excel_RangeReplace($GoogleGeoWorkbook,Default,"A:E",",0","")
            # ------------------------------------------------------------
            worksheet.Range("A:E").Replace(
                What=",0",
                Replacement="",
                LookAt=2,       # xlPart
                SearchOrder=1,
                MatchCase=False
            )

            # ------------------------------------------------------------
            # Save and close
            # ------------------------------------------------------------
            workbook.Save()
            workbook.Close(SaveChanges=True)
            workbook = None

            if show_message:
                self.show_auto_close_message(
                    title="Done",
                    message="Google Earth Geocode file formatted!",
                    duration_ms=1000
                )

            return {
                "input_file": google_geo_file,
                "csv_file": google_geo_csv_file,
                "completed_file": google_geo_csv_file_complete,
                "search_column": found_col_letter,
                "search_row": found_row_num,
                "line_count": google_geo_csv_complete_num_lines
            }

        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass

            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass

            pythoncom.CoUninitialize()

    def find_text_in_range(self,worksheet, text, range_address):
        """
        Equivalent to _Excel_RangeFind().
        Searches for text in a range and returns the found cell.

        Returns:
            found_cell
        """
        search_range = worksheet.Range(range_address)

        found_cell = search_range.Find(
            What=text,
            LookIn=-4163,        # xlValues
            LookAt=2,            # xlPart
            SearchOrder=1,       # xlByRows
            SearchDirection=1,   # xlNext
            MatchCase=False
        )

        if found_cell is None:
            raise ValueError(f"Could not find '{text}' in range {range_address}.")

        return found_cell

    def column_number_to_letter(col_num):

        result = ""

        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result

        return result

    def count_csv_lines(csv_file):
        with open(csv_file, "r", newline="", encoding="utf-8", errors="ignore") as f:
            return sum(1 for _ in f)

    def show_auto_close_message(title, message, duration_ms=1000):

        root = tk.Tk()
        root.withdraw()

        popup = tk.Toplevel(root)
        popup.title(title)
        popup.geometry("360x110")
        popup.resizable(False, False)

        label = tk.Label(
            popup,
            text=message,
            padx=20,
            pady=20,
            wraplength=320
        )
        label.pack(expand=True)

        popup.after(duration_ms, root.destroy)
        root.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    app = FileFormat(root)
    root.mainloop()