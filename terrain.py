import argparse
import pyautogui
from pywinauto import Application, Desktop, mouse
import time
from edx_open import EDX_OPEN

def _parse_ft(text):
    """Return the float value of a '123 ft' label, or None if not yet available."""
    text = text.replace(" ft", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None

def wait_for_query_results(dlg, timeout=300):
    """Wait for the terrain query to finish, then return its Min/Max values.

    After Run Query is clicked, EDX puts up a modal "Please wait..." dialog
    while it computes; the Query Studies Min/Max labels are only populated once
    that dialog closes. Waiting on that dialog appearing and then disappearing
    is far more reliable than guessing a fixed delay, and scales to large
    projects that take a while to compute.
    """
    please_wait = Desktop(backend="win32").window(title_re=".*[Pp]lease [Ww]ait.*")

    # The dialog may appear a beat after the click. Catch it if we can, but a
    # very fast query can finish before we ever see it, so don't fail if we
    # miss it -- the wait_not below will simply return immediately.
    try:
        please_wait.wait("visible", timeout=10)
    except TimeoutError:
        pass

    # Block until EDX finishes computing (the "Please wait..." dialog closes).
    please_wait.wait_not("visible", timeout=timeout)

    # Labels are populated once the dialog is gone; poll briefly to absorb any
    # small render lag before reading them.
    deadline = time.time() + 10
    while time.time() < deadline:
        min_val = _parse_ft(dlg.child_window(auto_id="_lblMinValue").window_text())
        max_val = _parse_ft(dlg.child_window(auto_id="_lblMaxValue").window_text())
        if min_val is not None and max_val is not None:
            return min_val, max_val
        time.sleep(0.3)
    raise TimeoutError("Query finished but Min/Max values did not appear")

def run_terrain_query():
    pyautogui.hotkey('f10')
    pyautogui.press(['q', 'down', 'enter'], interval=0.2)
    time.sleep(1)
    dlg = Desktop(backend="win32").window(title="Query Studies")

    # Select Terrain radio button
    dlg.child_window(auto_id="_rbTerrain").click_input()

    # Run Query
    dlg.child_window(auto_id="_btnRun").click_input()

    # Wait for the query to actually finish rather than assuming a fixed delay
    min_val, max_val = wait_for_query_results(dlg)

    # Close the dialog and confirm it actually closed. A single click can fire
    # before the dialog is ready to accept it (a human clicking is slower),
    # leaving Query Studies open -- which then swallows set_terrain's F10 menu
    # keys. Wait for the button, then retry until the dialog is gone.
    close_btn = dlg.child_window(auto_id="_btnClose")
    close_btn.wait("exists ready visible enabled", timeout=10)
    for _ in range(3):
        close_btn.click_input()
        try:
            dlg.wait_not("visible", timeout=3)
            break
        except TimeoutError:
            continue
    else:
        raise TimeoutError("Query Studies dialog did not close after clicking Close")

    return min_val, max_val

def get_terrain():
    app = Application(backend="uia").connect(title_re=".*EDX.*")
    dlg = app.top_window()
    dlg.set_focus()
    min_val, max_val = run_terrain_query()
    return min_val, max_val

def set_terrain(min, max):
    pyautogui.hotkey('f10')
    pyautogui.press(['d', 'right', 'down', 'enter'], interval=0.2)
    dlg = Desktop(backend="win32").window(title="Map Layers")
    # Large projects can be slow to open this dialog; wait for it to actually
    # appear rather than assuming a fixed delay is long enough.
    dlg.wait("exists ready visible", timeout=30)
    dlg.set_focus()
    listview = dlg.child_window(class_name="SysListView32").wrapper_object()
    listview.set_focus()
    listview.get_item("terrain").select()
    pyautogui.keyDown('alt')
    pyautogui.press('s')
    pyautogui.keyUp('alt')
    pyautogui.keyDown('alt')
    pyautogui.press('l')
    pyautogui.keyUp('alt')
    dlg2 = Desktop(backend="win32").window(title="Terrain Gradient Color Levels")
    dlg2.wait("exists ready visible", timeout=30)
    max_label = dlg2.child_window(auto_id="label_Max")
    rect = max_label.rectangle()
    x = rect.right + 20
    y = rect.top + ((rect.bottom - rect.top) // 2)
    pyautogui.click(x, y)
    time.sleep(0.2)
    pyautogui.press('end')
    for _ in range(15):   # enough to clear any number
        pyautogui.press('backspace')
    time.sleep(0.1)
    pyautogui.write(str(max), interval=0.05)
    min_label = dlg2.child_window(auto_id="label_Min")
    rect = min_label.rectangle()
    x = rect.right + 20
    y = rect.top + ((rect.bottom - rect.top) // 2)
    pyautogui.click(x, y)
    time.sleep(0.2)
    pyautogui.press('end')
    for _ in range(15):   # enough to clear any number
        pyautogui.press('backspace')
    time.sleep(0.1)
    pyautogui.write(str(min), interval=0.05)
    dlg2.child_window(auto_id="button_SetLevels").click_input()
    dlg2.child_window(auto_id="buttonOk").click_input()
    pyautogui.press('enter')
    edit_layers(dlg)

def edit_layers(dlg):
    listview = dlg.child_window(class_name="SysListView32").wrapper_object()
    terrain_layer = listview.get_item("terrain")
    terrain_layer.select()
    rect = terrain_layer.rectangle()
    map_x = rect.right +315
    legend_x = rect.right + 420
    y = rect.top + 333
    ensure_checkbox_on(map_x, y)
    ensure_checkbox_on(legend_x, y)
    downlink_layer = listview.get_item("Downlink (standard)")
    downlink_layer.select()
    rect = downlink_layer.rectangle()
    map_x = rect.right + 315
    legend_x = rect.right + 420
    y = rect.top + 333
    ensure_checkbox_off(map_x, y)
    ensure_checkbox_off(legend_x, y)
    bing_layer = listview.get_item("Bing(tm) Map Layer")
    bing_layer.select()
    rect = bing_layer.rectangle()
    map_x = rect.right + 315
    legend_x = rect.right + 420
    y = rect.top + 333
    ensure_checkbox_off(map_x, y)
    ensure_checkbox_off(legend_x, y)
    pyautogui.press('enter')
   
def is_checkbox_checked(x, y):
    r, g, b = pyautogui.pixel(x, y)
    brightness = int((r + g + b) / 3)
    if brightness == 255:
        return False
    else:
        return True

def ensure_checkbox_on(x, y):
    if not is_checkbox_checked(x, y):
        pyautogui.click(x, y)
        time.sleep(0.2)

def ensure_checkbox_off(x, y):
    if is_checkbox_checked(x, y):
        pyautogui.click(x, y)
        time.sleep(0.2)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("project_name", nargs="?", default=None)
    args = parser.parse_args()

    edx = EDX_OPEN()
    project = args.project_name or edx.get_project_name_from_edx()
    edx.confirm_edx_signalpro_open()
    min_val, max_val = get_terrain()
    set_terrain(min_val, max_val)

if __name__ == "__main__":

    main()
