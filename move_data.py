import tkinter as tk
from tkinter import simpledialog, messagebox
import time
import shutil

class MoveData:
    @staticmethod
    def copy_data_folder(project_name: str):
        
        if project_name == None:
            root = tk.Tk()
            root.withdraw()
            try:
                project_name = simpledialog.askstring("City and State", "Enter the Project Name\nie. 12345 Sample City AL R900")
                time.sleep(1)
                source = rf"\\TetonRiver\prop\Requests\{project_name}\Data"
                destination = rf"C:\ProgramData\EDX\Projects\{project_name}\Data"
                shutil.copytree(source, destination)
                messagebox.showinfo("","Data Folder Successfully Moved")
               
            except:
                messagebox.showinfo("Error","Cancelled") 
        else:
            source = rf"\\TetonRiver\prop\Requests\{project_name}\Data"
            destination = rf"C:\ProgramData\EDX\Projects\{project_name}\Data"
            shutil.copytree(source, destination)
