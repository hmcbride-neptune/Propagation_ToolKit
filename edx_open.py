import re
from tkinter import messagebox
from pywinauto import Application, Desktop

class EDX_OPEN:
    def get_edx_main_window(self):
        app = Application(backend="win32").connect(title_re=r".*\[initial view\].*")
        main = app.window(title_re=r".*\[initial view\].*")
        main.wait("visible enabled", timeout=5)
        return app, main

    def confirm_edx_signalpro_open(self):
        _, main = self.get_edx_main_window()
        main.set_focus()

    def get_project_name_from_edx(self):
        try:
            _, main = self.get_edx_main_window()
            title = main.window_text()

            match = re.search(
                r"64-bit\s*-\s*(.*?)\s*-\s*\[initial view\]",
                title,
                re.IGNORECASE
            )

            if not match:
                raise RuntimeError("Could not parse project name")
        except:
            messagebox.showerror("Error","No EDX Project is open.")
        return match.group(1).strip()
