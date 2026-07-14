import csv
import tkinter as tk

ASSUMPTIONS_CSV = r"\\TetonRiver\prop\Template\Propagation Assumptions.csv"


class AssumptionsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Assumptions")
        self.root.geometry("800x500")
        self.root.attributes("-topmost", True)

        self.assumption_final = ""
        self.no = 1

        self.assump_titles = {}
        self.assumptions = {}

        self.get_assumptions()
        self.build_gui()

    def build_gui(self):
        left_frame = tk.Frame(self.root)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        right_frame = tk.Frame(self.root)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Text box (equivalent to $Input31)
        self.text_box = tk.Text(right_frame, wrap="word")
        self.text_box.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Button mapping (same as AutoIt)
        button_mapping = [
            (1, 2), (2, 3), (3, 4),
            (4, 5), (5, 6), (6, 7), (7, 8),
            (8, 9)
        ]

        for i, idx in button_mapping:
            title = self.assump_titles.get(idx, f"Assumption {idx}")

            tk.Button(
                left_frame,
                text=title,
                anchor="w",
                width=35,
                command=lambda x=idx: self.add_assumption(x)
            ).pack(fill=tk.X, pady=2)

        # Clear button (Button311)
        tk.Button(
            left_frame,
            text="Clear",
            width=35,
            command=self.clear
        ).pack(pady=10)
        tk.Label(left_frame, 
                 text="Click on the button to copy Assumption to ClipBoard. Close this window after Assumptions have been selected:",
                 justify="left",
                 wraplength=250,
                 anchor="w"
                 ).pack(pady=(5, 5), fill=tk.X)

    def add_assumption(self, idx):
        text = self.assumptions.get(idx, "")

        if self.assumption_final == "":
            self.assumption_final = f"Assumptions\n{self.no}. {text}"
        else:
            self.assumption_final += f"\n{self.no}. {text}"

        self.no += 1

        # Clipboard (ClipPut)
        self.root.clipboard_clear()
        self.root.clipboard_append(self.assumption_final)
        self.root.update()

        # Update text box (GUICtrlSetData)
        self.text_box.delete("1.0", tk.END)
        self.text_box.insert(tk.END, self.assumption_final)

    def clear(self):
        self.assumption_final = ""
        self.no = 1

        self.root.clipboard_clear()
        self.text_box.delete("1.0", tk.END)

    def get_assumptions(self):
        with open(ASSUMPTIONS_CSV, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            index = 1

            for row in reader:
                if not row or len(row) < 2:
                    continue

                self.assump_titles[index] = row[0].strip()
                self.assumptions[index] = ",".join(row[1:]).strip()

                index += 1


if __name__ == "__main__":
    root = tk.Tk()
    app = AssumptionsGUI(root)
    root.mainloop()
    
