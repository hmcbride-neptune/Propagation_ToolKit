import argparse
import time
from edx_open import EDX_OPEN
from pywinauto.keyboard import send_keys
from pywinauto import Desktop


class EDXAutomation:
    def __init__(self):
         pass
    
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
    



    def get_listview_items(self, listview):
        """
        Returns list of:
        [
            (index, text, item_object),
            ...
        ]
        """
        items = []

        for i, item in enumerate(listview.items()):
            text = item.text().strip()
            items.append((i, text))

        return items


    def find_item_index(self, listview, target_text):
        target_text = target_text.strip().lower()

        for i, text in self.get_listview_items(listview):
            if text.strip().lower() == target_text:
                return i

        raise ValueError(f'Could not find list item: "{target_text}"')



    def move_layer(self):
        self.wait_for_window(".*initial view.*")

        send_keys("%d")
        send_keys("{RIGHT 1}")
        send_keys("{DOWN}")
        send_keys("{ENTER}")

        dlg = Desktop(backend="win32").window(title_re=".*Map Layers.*")
        dlg.wait("exists visible ready", timeout=10)
        dlg.set_focus()

        
        listview = dlg.child_window(class_name="SysListView32").wrapper_object()
        listview.set_focus()
        layer_index = self.find_item_index(listview, "Downlink (standard)")
        anchor_index = self.find_item_index(listview, "measurement data: Meters")
        if layer_index < anchor_index:
            move_down = True
            move_up = False
        else:
            move_down = False
            move_up = True

        while move_down:
            listview.get_item("Downlink (standard)").select()
            time.sleep(0.2)
            send_keys("%d")
            layer_index = self.find_item_index(listview, "Downlink (standard)")
            anchor_index = self.find_item_index(listview, "measurement data: Meters")
            if layer_index > anchor_index:
                move_down = False
                send_keys("%o")
                return

        while move_up:
            # print(f"downlink layer: {layer_index}\nmeter layer: {anchor_index}")
            listview.get_item("Downlink (standard)").select()
            send_keys("%u")
            layer_index = self.find_item_index(listview, "Downlink (standard)")
            anchor_index = self.find_item_index(listview, "measurement data: Meters")
            if anchor_index > layer_index:
                move_up = False
                send_keys("%o")
                return

    def main(self):
        parser = argparse.ArgumentParser()
        parser.add_argument("project_name", nargs="?", default=None)
        args = parser.parse_args()

        edx = EDX_OPEN()
        project = args.project_name or edx.get_project_name_from_edx()
        edx.confirm_edx_signalpro_open()
        self.move_layer()
    

if __name__ == "__main__":
    proj = EDXAutomation()
    proj.main()
