import argparse
import pyautogui
from pywinauto import Application, Desktop, mouse
import time
from edx_open import EDX_OPEN

def run_terrain_query():
    pyautogui.hotkey('f10')
    pyautogui.press(['q', 'down', 'enter'], interval=0.2)
    time.sleep(1)
    dlg = Desktop(backend="win32").window(title="Query Studies")

    # Select Terrain radio button
    dlg.child_window(auto_id="_rbTerrain").click_input()

    # Run Query
    dlg.child_window(auto_id="_btnRun").click_input()
    time.sleep(2)

    # Get Min and Max 
    min_text = dlg.child_window(auto_id="_lblMinValue").window_text()
    max_text = dlg.child_window(auto_id="_lblMaxValue").window_text()

    min_val = float(min_text.replace(" ft", "").strip())
    max_val = float(max_text.replace(" ft", "").strip())
    dlg.child_window(auto_id="_btnClose").click_input()
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
    time.sleep(1)
    dlg = Desktop(backend="win32").window(title="Map Layers")
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
