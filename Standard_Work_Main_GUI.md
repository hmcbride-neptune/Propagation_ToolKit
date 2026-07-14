# Standard Work — Propagation Toolkit (Main GUI)

| | |
|---|---|
| **Document** | Standard Work Instruction — Propagation Toolkit Main Toolbar |
| **Applies to** | `Main_GUI.py` / `Propagation.exe` (build 260714) |
| **Owner** | Propagation / RF Engineering |
| **Last updated** | 2026-07-14 |

---

## 1. Purpose

The Propagation Toolkit is a Windows desktop toolbar that automates common RF
propagation-study tasks around **EDX SignalPro**. It appears as a single
always-on-top row of 24 icon buttons. Each button launches one focused sub-tool
(create a request folder, format files, run terrain/demographic studies, export
transmitters, take screenshots, run the FAA/ASR check, zip & upload, etc.).

This document explains how to start the program and what every button does.

---

## 2. Prerequisites (check before you start)

| # | Requirement | Why |
|---|-------------|-----|
| 1 | **Windows** PC | The app uses Win32 APIs and Windows-only automation. |
| 2 | **EDX SignalPro** installed and, for most tools, **open with a project loaded** | The toolbar reads the active project from the EDX window title, and most tools automate EDX directly. |
| 3 | **Network access to `\\TetonRiver\prop\...`** | Your personal settings file and request/template folders live there. |
| 4 | **`C:\ProgramData\Python\` tools present** | The Meter Combine, Tailored Output, and Best Collector buttons run helper programs from this folder. |
| 5 | **Microsoft Edge** installed | Required only for the FAA/ASR button (Selenium web automation). |
| 6 | A **user settings file** at `\\TetonRiver\prop\template\zAutoit\User Settings\<yourusername> settings.txt` | Controls toolbar transparency, background color, and tool preferences. |

---

## 3. How to start the program

### Option A — Packaged application (normal use)

1. Browse to the build output folder:
   `...\Propagation\dist\Propagation\`
2. Double-click **`Propagation.exe`**.
3. The toolbar appears at the top of the screen, always on top of other windows.

> Tip: create a desktop shortcut to `Propagation.exe` for one-click launch.

### Option B — From source (developers only)

```powershell
# From the project directory
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python Main_GUI.py
```

---

## 4. Understanding the main window

- The toolbar is a **single horizontal row of 24 square icon buttons**.
- **Hover over any button** to see a tooltip describing it.
- **Title bar** shows the active project, e.g. `Propagation - Main | 97035 Lake Oswego Modeling OR R900`.
  - The app checks the open EDX window every ~2 seconds and updates this
    automatically. If EDX is **not** open, the last known project name is kept.
- The window is **always-on-top** and its transparency/background color come
  from your personal settings file.

### The "active project" rule (important)

Many buttons need to know which project you are working on. The project name is
taken from:

1. the **open EDX SignalPro project** (preferred), or
2. the project you just created with **New Request**.

If neither is available, those buttons show a **"No Project Found"** message and
do nothing. Fix it by opening the project in EDX (or running **New Request**),
then click the button again.

### Common behavior

- Tools that drive EDX will **minimize the toolbar** while they run and
  **restore it** when finished — this is normal.
- Closing the toolbar window exits the application.

---

## 5. Button reference (left → right, 1–24)

> Positions are counted left-to-right. Icon file names are in the `Logos\` folder.

| # | Icon file | Name / tooltip | What it does | Needs active project? |
|---|-----------|----------------|--------------|:---:|
| 1 | `New Request.bmp` | **Make new request folder** | Opens the New Request window to create a new project/request folder. Sets the active project name for the other tools. | No — this sets it |
| 2 | `fileformat2.png` | **Format files** (Meters.csv, ESRI export, etc.) | Runs the file-formatting tool on the current project's data files. | Uses it if available |
| 3 | `Meters_combined.bmp` | **Create Meter files from Inbound and Outbound** | Runs `Meter Combine\metersCombined.py` from `C:\ProgramData\Python`. | No |
| 4 | `nEW pROJECT.bmp` | **Moves Data Folder** | Copies the project **Data** folder from `\\TetonRiver\prop\Requests\<project>\Data` to `C:\ProgramData\EDX\Projects\<project>\Data`. If no project is known, prompts you to type the project name. | Prompts if unknown |
| 5 | `addfile.bmp` | **Add Project Folder** | Adds project files. When prompted, **browse to the project's `xReports` subfolder**. | No |
| 6 | `terrain.bmp` | **Create terrain map** | Builds a terrain map based on the max/min elevation in the study area (automates EDX). | **Yes** |
| 7 | `switch group.bmp` | **Switch Group displayed/calculated** | Switches which group EDX is currently displaying or calculating. | **Yes** |
| 8 | `measurement.bmp` | **Measurement Analysis** | Runs the measurement-analysis tool. | No |
| 9 | `demographic.png` | **Demographic Studies** | Runs a demographic study for the current project. | **Yes** |
| 10 | `finds.bmp` | **Re-name Find Locations** | Renames "find" (measurement/site) locations in the project. | **Yes** |
| 11 | `screenshot.png` | **Take screenshots** | Launches the screenshot tool (captures EDX map views). Uses your Windows username for saved settings/paths. | No |
| 12 | `showhide.bmp` | **Show / Hide inactive transmitters** | Shows or hides inactive transmitter locations in EDX. | **Yes** |
| 13 | `hourglass.png` | **Switch Daily / Hourly / Billing** | Toggles the study between daily, hourly, and billing modes. | **Yes** |
| 14 | `movecoverage.bmp` | **Move Coverage layer** | Moves the coverage layer above or below the meters layer in EDX. | **Yes** |
| 15 | `export (1).bmp` | **Export Transmitters** | Exports the project's transmitters. Shows "No Project Found" if none is active. | **Yes** |
| 16 | `presentation.png` | **Presentation / Map indicator** | Either builds the presentation or fills in the map indicator, depending on your preferences. | Uses it if available |
| 17 | `reset.bmp` | **Reset / Restart** | Restarts the toolbar application (relaunches a fresh session). | No |
| 18 | `airplane.png` | **Run FAA / ASR Script** | Runs the FAA/ASR web check (Selenium via Microsoft Edge). | No |
| 19 | `zipup.png` | **Zip and Upload** | Zips project output and uploads it. | No |
| 20 | `question.png` | **Assumptions** | Opens the assumptions tool. | No |
| 21 | `TailoredLogo.bmp` | **Create a tailored output** | Runs the **latest** `Tailored_Output_YYMMDD` script found in `C:\ProgramData\Python`. | No |
| 22 | `calc.bmp` | **By-collector coverage statistics** | Runs `Best_Collector.exe` from `C:\ProgramData\Python\Best_Collector`. | No |
| 23 | `dataImport.bmp` | **Data Import** | Opens the Data Import window. | No |
| 24 | `settings.png` | **Change Settings** | Opens the Settings window (transparency, background color, tool preferences, and shows the current project name/path). | No |

---

## 6. Typical workflow (example)

1. **Open EDX SignalPro** and load the project you're working on (or click
   **New Request** (button 1) to create one).
2. Confirm the toolbar **title bar** shows the correct project name.
3. **Move Data Folder** (button 4) if the project data isn't local yet.
4. **Format files** (button 2) and **Add Project Folder** (button 5) as needed.
5. Run the required study tools — **Terrain** (6), **Demographics** (9),
   **Switch Group** (7), **Daily/Hourly/Billing** (13), etc.
6. **Take screenshots** (button 11) and build the **Presentation** (button 16).
7. **Export Transmitters** (button 15) and generate a **Tailored Output**
   (button 21) if required.
8. **Zip and Upload** (button 19) to deliver the results.

---

## 7. Troubleshooting

| Symptom | Likely cause | Fix |
|---------|--------------|-----|
| **"No Project Found"** message | No EDX project open and none created this session | Open the project in EDX (or run **New Request**), then retry. |
| Title bar doesn't show the project | EDX not open, or window title not recognized | Make sure EDX SignalPro is open with the project loaded; wait a couple seconds for auto-detection. |
| Toolbar disappears when a tool runs | Normal — it minimizes while a study runs | It restores automatically when the tool finishes. |
| Button 3 / 21 / 22 does nothing | Helper program missing from `C:\ProgramData\Python\...` | Confirm Meter Combine / Tailored Output / Best Collector are installed there. |
| FAA/ASR button fails | Microsoft Edge missing or blocked | Install/enable Edge; the tool auto-manages its WebDriver. |
| Settings/colors not applied | Missing personal settings file | Create `\\TetonRiver\prop\template\zAutoit\User Settings\<username> settings.txt`. |

---

*Reference: [Main_GUI.py](Main_GUI.py) (button definitions at lines 180–205).*
