import tkinter as tk
from datetime import datetime
import pytz
import json
import os
import sys
import tkinter.font as tkfont
from tkinter import messagebox

# Windows API for Desktop Integration
try:
    import win32gui
    import win32con
    import win32api
    import win32com.client
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

CONFIG_FILE = "config.json"
DEFAULT_POS = "+100+100"

class DesktopClock:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Horloge Bureau Deluxe")
        
        # Load position
        pos = self.load_position()
        try:
            self.root.geometry(pos)
        except:
            self.root.geometry(DEFAULT_POS)
            
        # UI Configuration
        self.root.overrideredirect(True)
        self.root.configure(bg="#000001") # Near-black for transparency key
        self.root.wm_attributes("-transparentcolor", "#000001")
        
        if HAS_WIN32:
            self.setup_windows_style()
            self.root.after(200, self.stick_to_desktop)

        # Theme Definitions
        self.themes = {
            "Bénin": {
                "tz": "Africa/Porto-Novo",
                "stripes": ["#008751", "#FCD116", "#E8112D"],
                "cities": "COTONOU • PORTO-NOVO • PARAKOU • DJOUGOU"
            },
            "France": {
                "tz": "Europe/Paris",
                "stripes": ["#002395", "#FFFFFF", "#ED2939"],
                "cities": "PARIS • LYON • MARSEILLE • BORDEAUX"
            },
            "Espagne": {
                "tz": "Europe/Madrid",
                "stripes": ["#AA151B", "#F1BF00", "#AA151B"],
                "cities": "MADRID • BARCELONE • VALENCE • SÉVILLE"
            }
        }

        # Fonts
        self.time_font = tkfont.Font(family="Segoe UI Variable Display", size=62, weight="bold")
        self.city_font = tkfont.Font(family="Segoe UI Variable Display", size=10, weight="normal")
        self.title_font = tkfont.Font(family="Segoe UI Variable Display", size=15, weight="bold")

        # Main Container
        self.main_frame = tk.Frame(self.root, bg="#000001")
        self.main_frame.pack(padx=30, pady=30)

        self.clock_widgets = []
        for name, data in self.themes.items():
            self.create_clock_card(name, data)

        # Right-click menu
        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label="Créer raccourci sur le Bureau", command=self.create_desktop_shortcut)
        self.menu.add_command(label="Mettre au démarrage", command=self.toggle_startup)
        self.menu.add_separator()
        self.menu.add_command(label="Quitter", command=self.quit_app)
        self.root.bind("<Button-3>", self.show_menu)
        
        # Draggable
        self.root.bind("<Button-1>", self.start_move)
        self.root.bind("<B1-Motion>", self.do_move)

        self.update_clocks()
        self.root.mainloop()

    def setup_windows_style(self):
        """Ensure the window doesn't steal focus or taskbar space."""
        hwnd = win32gui.GetParent(self.root.winfo_id())
        
        # WS_EX_NOACTIVATE (0x08000000) prevents taking focus
        # WS_EX_TOOLWINDOW (0x00000080) hides from taskbar
        # WS_EX_LAYERED (0x00080000) for transparency
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        ex_style |= win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOOLWINDOW | win32con.WS_EX_LAYERED
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style)

    def stick_to_desktop(self):
        """Parent to WorkerW to stay behind all icons and windows."""
        if not HAS_WIN32: return
        progman = win32gui.FindWindow("Progman", "Program Manager")
        win32gui.SendMessageTimeout(progman, 0x052C, 0, 0, win32con.SMTO_NORMAL, 1000)
        
        def enum_handler(hwnd, ctx):
            if win32gui.FindWindowEx(hwnd, 0, "SHELLDLL_DefView", None):
                ctx['workerw'] = win32gui.FindWindowEx(0, hwnd, "WorkerW", None)
        
        ctx = {'workerw': None}
        win32gui.EnumWindows(enum_handler, ctx)
        
        if ctx['workerw']:
            hwnd = self.root.winfo_id()
            win32gui.SetParent(hwnd, ctx['workerw'])

    def create_clock_card(self, name, data):
        card = tk.Frame(self.main_frame, bg="#000001", pady=20)
        card.pack(fill="x")

        # Header with Vertical Flag Strip
        header = tk.Frame(card, bg="#000001")
        header.pack(fill="x")
        
        flag_bar = tk.Frame(header, bg="#000001", width=5)
        flag_bar.pack(side="left", fill="y", padx=(0, 20))
        for color in data["stripes"]:
            tk.Frame(flag_bar, bg=color, height=10, width=5).pack(fill="x")

        tk.Label(header, text=name.upper(), font=self.title_font, fg="#888888", bg="#000001").pack(side="left")

        # Time and Cities
        time_lbl = tk.Label(card, text="00:00", font=self.time_font, fg="#ffffff", bg="#000001")
        time_lbl.pack(anchor="w", padx=(25, 0))

        tk.Label(card, text=data["cities"], font=self.city_font, fg="#555555", bg="#000001").pack(anchor="w", padx=(25, 0))
        
        self.clock_widgets.append({"label": time_lbl, "tz": pytz.timezone(data["tz"])})

    def update_clocks(self):
        for cw in self.clock_widgets:
            now = datetime.now(cw["tz"])
            cw["label"].config(text=now.strftime("%H:%M"))
        self.root.after(1000, self.update_clocks)

    def start_move(self, event):
        self.start_x = event.x
        self.start_y = event.y

    def do_move(self, event):
        x = self.root.winfo_x() + (event.x - self.start_x)
        y = self.root.winfo_y() + (event.y - self.start_y)
        self.root.geometry(f"+{x}+{y}")
        self.save_position(f"+{x}+{y}")

    def show_menu(self, event):
        self.menu.post(event.x_root, event.y_root)

    def load_position(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    return json.load(f).get("position", DEFAULT_POS)
            except: return DEFAULT_POS
        return DEFAULT_POS

    def save_position(self, pos):
        with open(CONFIG_FILE, "w") as f:
            json.dump({"position": pos}, f)

    def create_desktop_shortcut(self):
        if not HAS_WIN32: return
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        lnk = os.path.join(desktop, "Horloge Bureau.lnk")
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk)
        # Use pythonw to avoid console
        pythonw = sys.executable.replace("python.exe", "pythonw.exe")
        shortcut.Targetpath = pythonw
        shortcut.Arguments = f'"{os.path.abspath(__file__)}"'
        shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(__file__))
        shortcut.save()
        messagebox.showinfo("Raccourci", "Raccourci créé sur le Bureau. Glissez-le dans la barre des tâches !")

    def toggle_startup(self):
        if not HAS_WIN32: return
        startup = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
        lnk = os.path.join(startup, "HorlogeBureau.lnk")
        if os.path.exists(lnk):
            os.remove(lnk)
            messagebox.showinfo("Démarrage", "Retiré du démarrage.")
        else:
            self.create_desktop_shortcut() # reuse the logic
            # Move the link to startup
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            src = os.path.join(desktop, "Horloge Bureau.lnk")
            if os.path.exists(src):
                os.rename(src, lnk)
            messagebox.showinfo("Démarrage", "Ajouté au démarrage.")

    def quit_app(self):
        self.root.destroy()
        sys.exit()

if __name__ == "__main__":
    DesktopClock()
