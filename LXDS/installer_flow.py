import os
import sys
import tkinter as tk
from tkinter import messagebox

EULA_TEXT = """This application is intended to perform system-level operations and structural modifications on your PC. It requires standard local file system access, Python runtime execution privileges, and spreadsheet engine bindings.

BRANDING & OWNERSHIP:
This project is an initiative by AXC (Alex Creations Lab) under the LX Programs initiative. All rights, architectural patterns (LXComm), and software suite specifications are managed under AXC guidelines.

THANKSGIVING & FEEDBACK:
Thank you for using LX Programs. Please provide feedback, bug reports, or feature requests directly to the developer/distribution source.

(Scroll to the bottom to accept)
""" + "\n" * 20 + "End of Agreement."

def check_dependencies():
    REQUIRED_MODULES = ["xlsxwriter", "openpyxl"]
    missing = []
    
    for mod in REQUIRED_MODULES:
        try:
            __import__(mod)
        except ImportError:
            missing.append(mod)
            
    if missing:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Missing Dependencies",
            f"The following required Python packages are missing:\n\n{', '.join(missing)}\n\nPlease run this command in your terminal:\npip install {' '.join(missing)}"
        )
        root.destroy()
        sys.exit(1)

class LicenseAgreementModal:
    def __init__(self, root, on_accept):
        self.top = tk.Toplevel(root)
        self.top.title("LX Programs - End User License Agreement")
        self.top.geometry("500x400")
        self.top.protocol("WM_DELETE_WINDOW", self.on_decline)
        self.top.transient(root)
        self.top.grab_set()
        
        self.on_accept_callback = on_accept
        
        tk.Label(self.top, text="End User License Agreement", font=("Helvetica", 12, "bold")).pack(pady=10)
        
        self.text_area = tk.Text(self.top, wrap="word", width=55, height=15)
        self.text_area.pack(padx=10, pady=5)
        self.text_area.insert("1.0", EULA_TEXT)
        self.text_area.config(state="disabled")
        
        # Scrollbar
        self.scrollbar = tk.Scrollbar(self.text_area)
        self.scrollbar.pack(side="right", fill="y")
        self.text_area.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.text_area.yview)
        
        self.text_area.bind("<MouseWheel>", self.check_scroll)
        self.text_area.bind("<B1-Motion>", self.check_scroll)
        self.scrollbar.bind("<B1-Motion>", self.check_scroll)
        
        btn_frame = tk.Frame(self.top)
        btn_frame.pack(pady=10)
        
        self.btn_accept = tk.Button(btn_frame, text="Accept & Continue", command=self.on_accept, state=tk.DISABLED)
        self.btn_accept.pack(side="left", padx=10)
        
        tk.Button(btn_frame, text="Decline", command=self.on_decline).pack(side="right", padx=10)
        
    def check_scroll(self, event=None):
        # Allow a little bit of time for scrollbar to update
        self.top.after(100, self._check_scroll_pos)
        
    def _check_scroll_pos(self):
        # scrollbar.get() returns (top, bottom) fractions
        try:
            if self.scrollbar.get()[1] >= 0.99:
                self.btn_accept.config(state=tk.NORMAL)
        except Exception:
            pass

    def on_accept(self):
        with open(".license_accepted", "w") as f:
            f.write("Accepted")
        self.top.destroy()
        self.on_accept_callback()
        
    def on_decline(self):
        sys.exit(0)

def enforce_lifecycle(root, main_app_launcher):
    check_dependencies()
    if not os.path.exists(".license_accepted"):
        root.withdraw() # Hide root until accepted
        LicenseAgreementModal(root, lambda: (root.deiconify(), main_app_launcher()))
    else:
        main_app_launcher()
