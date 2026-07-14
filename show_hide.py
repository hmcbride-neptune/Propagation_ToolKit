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
    
    def show_hide(self):
        
        try:
            self.wait_for_window(".*initial view.*")
        except:
            raise "Make sure an EDX Project is open."

        send_keys("%d")
        send_keys("{RIGHT 1}")
        send_keys("{DOWN}")
        send_keys("{ENTER}")

        dlg = Desktop(backend="win32").window(title_re=".*Map Layers.*")
        listview_spec = dlg.child_window(class_name="SysListView32")
        listview = listview_spec.wrapper_object()
        listview.set_focus()
        listview.get_item("transmitter/base/hub sites").select()
        time.sleep(0.2)
        send_keys("%s")
        send_keys("{TAB 15}")
        time.sleep(0.2)
        send_keys("{SPACE}")
        send_keys("%o")
        send_keys("%o")

    def main(self):
        self.show_hide()
    

if __name__ == "__main__":
    proj = EDXAutomation()
    proj.main()
