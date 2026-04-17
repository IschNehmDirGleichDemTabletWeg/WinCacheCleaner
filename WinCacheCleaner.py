import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import glob
import shutil
import subprocess
import ctypes
import winreg
import time


# -------------------------------------------------------
# Logging
# -------------------------------------------------------
def get_log_path():
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "WinCacheCleaner.log")


def write_log(msg):
    from datetime import datetime
    try:
        with open(get_log_path(), "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    except Exception:
        pass


# -------------------------------------------------------
# UAC self-elevation
# -------------------------------------------------------
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def elevate():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )
    sys.exit()


# -------------------------------------------------------
# Helper
# -------------------------------------------------------
def expand(path):
    return os.path.expandvars(path)


def set_status(msg, color="white"):
    status_var.set(msg)
    status_label.config(fg=color)
    root.update_idletasks()


def add_tooltip(widget, text):
    tip_window = []

    def show(event):
        x = widget.winfo_rootx() + 20
        y = widget.winfo_rooty() + widget.winfo_height() + 4
        tw = tk.Toplevel(widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(
            tw,
            text=text,
            justify="left",
            background="#FFFFE0",
            foreground="#333333",
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 9),
            wraplength=320,
            padx=6,
            pady=4,
        )
        label.pack()
        tip_window.append(tw)

    def hide(event):
        for tw in tip_window:
            tw.destroy()
        tip_window.clear()

    widget.bind("<Enter>", show)
    widget.bind("<Leave>", hide)


# -------------------------------------------------------
# Cache-Cleaner Funktionen
# -------------------------------------------------------

def clear_recent_files():
    path = expand(r"%APPDATA%\Microsoft\Windows\Recent")
    if not os.path.exists(path):
        msg = "Recent Files: Ordner nicht vorhanden - nichts zu tun"
        set_status(msg, "lightyellow")
        write_log(msg)
        return
    count = 0
    errors = 0
    try:
        for item in os.listdir(path):
            full = os.path.join(path, item)
            try:
                if os.path.isfile(full):
                    os.remove(full)
                    count += 1
                elif os.path.isdir(full):
                    shutil.rmtree(full)
                    count += 1
            except Exception as e:
                errors += 1
                write_log(f"Recent Files: Fehler bei '{full}' - {e}")
        msg = f"Recent Files: {count} Eintraege geloescht"
        if errors:
            msg += f" ({errors} Fehler)"
        set_status(msg, "lightgreen")
        write_log(msg)
    except Exception as e:
        msg = f"Recent Files: Fehler - {e}"
        set_status(msg, "salmon")
        write_log(msg)


def clear_automatic_destinations():
    path = expand(r"%APPDATA%\Microsoft\Windows\Recent\AutomaticDestinations")
    if not os.path.exists(path):
        msg = "Jump Lists: Ordner nicht vorhanden - nichts zu tun"
        set_status(msg, "lightyellow")
        write_log(msg)
        return
    count = 0
    errors = 0
    try:
        for item in os.listdir(path):
            full = os.path.join(path, item)
            try:
                os.remove(full)
                count += 1
            except Exception as e:
                errors += 1
                write_log(f"Jump Lists: Fehler bei '{full}' - {e}")
        msg = f"Jump Lists: {count} Eintraege geloescht"
        if errors:
            msg += f" ({errors} Fehler)"
        set_status(msg, "lightgreen")
        write_log(msg)
    except Exception as e:
        msg = f"Jump Lists: Fehler - {e}"
        set_status(msg, "salmon")
        write_log(msg)


def clear_thumbnail_cache():
    path = expand(r"%LOCALAPPDATA%\Microsoft\Windows\Explorer")
    if not os.path.exists(path):
        msg = "Thumbnail Cache: Ordner nicht vorhanden - nichts zu tun"
        set_status(msg, "lightyellow")
        write_log(msg)
        return
    pattern = os.path.join(path, "thumbcache*")
    count = 0
    errors = 0
    try:
        subprocess.run(["taskkill", "/f", "/im", "explorer.exe"],
                       capture_output=True)
        time.sleep(1)
        for f in glob.glob(pattern):
            try:
                os.remove(f)
                count += 1
            except Exception as e:
                errors += 1
                write_log(f"Thumbnail Cache: Fehler bei '{f}' - {e}")
        subprocess.Popen(["explorer.exe"])
        msg = f"Thumbnail Cache: {count} Dateien geloescht"
        if errors:
            msg += f" ({errors} Fehler)"
        set_status(msg, "lightgreen")
        write_log(msg)
    except Exception as e:
        msg = f"Thumbnail Cache: Fehler - {e}"
        set_status(msg, "salmon")
        write_log(msg)


def clear_icon_cache():
    path = expand(r"%LOCALAPPDATA%\Microsoft\Windows\Explorer")
    if not os.path.exists(path):
        msg = "Icon Cache: Ordner nicht vorhanden - nichts zu tun"
        set_status(msg, "lightyellow")
        write_log(msg)
        return
    pattern_1 = os.path.join(path, "iconcache*")
    db_path = expand(r"%LOCALAPPDATA%\IconCache.db")
    count = 0
    errors = 0
    try:
        subprocess.run(["taskkill", "/f", "/im", "explorer.exe"],
                       capture_output=True)
        time.sleep(1)
        for f in glob.glob(pattern_1):
            try:
                os.remove(f)
                count += 1
            except Exception as e:
                errors += 1
                write_log(f"Icon Cache: Fehler bei '{f}' - {e}")
        try:
            if os.path.exists(db_path):
                os.remove(db_path)
                count += 1
        except Exception as e:
            errors += 1
            write_log(f"Icon Cache: Fehler bei IconCache.db - {e}")
        subprocess.Popen(["explorer.exe"])
        msg = f"Icon Cache: {count} Dateien geloescht"
        if errors:
            msg += f" ({errors} Fehler)"
        set_status(msg, "lightgreen")
        write_log(msg)
    except Exception as e:
        msg = f"Icon Cache: Fehler - {e}"
        set_status(msg, "salmon")
        write_log(msg)


def clear_prefetch():
    if not is_admin():
        msg = "Prefetch: Administrator-Rechte benoetigt!"
        set_status(msg, "orange")
        write_log(msg)
        return
    path = r"C:\Windows\Prefetch"
    if not os.path.exists(path):
        msg = "Prefetch: Ordner nicht vorhanden - nichts zu tun"
        set_status(msg, "lightyellow")
        write_log(msg)
        return
    count = 0
    errors = 0
    try:
        for f in glob.glob(os.path.join(path, "*.pf")):
            try:
                os.remove(f)
                count += 1
            except Exception as e:
                errors += 1
                write_log(f"Prefetch: Fehler bei '{f}' - {e}")
        msg = f"Prefetch: {count} Dateien geloescht"
        if errors:
            msg += f" ({errors} Fehler)"
        set_status(msg, "lightgreen")
        write_log(msg)
    except Exception as e:
        msg = f"Prefetch: Fehler - {e}"
        set_status(msg, "salmon")
        write_log(msg)


def clear_mui_cache():
    key_paths = [
        r"Software\Microsoft\Windows\ShellNoRoam\MUICache",
        r"Software\Classes\Local Settings\MuiCache",
    ]
    total = 0
    found = 0
    for key_path in key_paths:
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path,
                                 0, winreg.KEY_ALL_ACCESS)
            found += 1
            count = 0
            while True:
                try:
                    name, _, _ = winreg.EnumValue(key, 0)
                    winreg.DeleteValue(key, name)
                    count += 1
                except OSError:
                    break
            winreg.CloseKey(key)
            total += count
            write_log(f"MUI Cache: '{key_path}' - {count} Eintraege geloescht")
        except FileNotFoundError:
            write_log(f"MUI Cache: '{key_path}' - nicht vorhanden, uebersprungen")
        except Exception as e:
            write_log(f"MUI Cache: '{key_path}' - Fehler: {e}")
    if found == 0:
        msg = "MUI Cache: Keine Schlussel gefunden (bereits leer oder nicht vorhanden)"
        set_status(msg, "lightyellow")
    else:
        msg = f"MUI Cache: {total} Eintraege geloescht ({found} Schlussel)"
        set_status(msg, "lightgreen")
    write_log(msg)


def clear_shellbags():
    keys = [
        r"Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags",
        r"Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU",
        r"Software\Microsoft\Windows\Shell\Bags",
        r"Software\Microsoft\Windows\Shell\BagMRU",
    ]
    count = 0
    errors = 0
    for key_path in keys:
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
            count += 1
        except FileNotFoundError:
            pass
        except Exception:
            errors += 1

    def delete_tree(hive, path):
        try:
            key = winreg.OpenKey(hive, path, 0, winreg.KEY_ALL_ACCESS)
            while True:
                try:
                    subkey = winreg.EnumKey(key, 0)
                    delete_tree(key, subkey)
                except OSError:
                    break
            winreg.CloseKey(key)
            winreg.DeleteKey(hive, path)
        except Exception:
            pass

    for key_path in keys:
        delete_tree(winreg.HKEY_CURRENT_USER, key_path)

    msg = f"ShellBags: Eintraege bereinigt (Fehler: {errors}) - Abmelden empfohlen"
    set_status(msg, "lightgreen")
    write_log(msg)


def clear_runmru():
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path,
                             0, winreg.KEY_ALL_ACCESS)
        count = 0
        while True:
            try:
                name, _, _ = winreg.EnumValue(key, 0)
                winreg.DeleteValue(key, name)
                count += 1
            except OSError:
                break
        winreg.CloseKey(key)
        msg = f"RunMRU: {count} Eintraege geloescht"
        set_status(msg, "lightgreen")
        write_log(msg)
    except FileNotFoundError:
        msg = "RunMRU: Schlussel nicht gefunden (bereits leer?)"
        set_status(msg, "lightyellow")
        write_log(msg)
    except Exception as e:
        msg = f"RunMRU: Fehler - {e}"
        set_status(msg, "salmon")
        write_log(msg)


def clear_dns_cache():
    try:
        result = subprocess.run(
            ["ipconfig", "/flushdns"],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            msg = "DNS Cache: erfolgreich geleert"
            set_status(msg, "lightgreen")
            write_log(msg)
        else:
            msg = f"DNS Cache: Fehler - {result.stderr.strip()}"
            set_status(msg, "salmon")
            write_log(msg)
    except Exception as e:
        msg = f"DNS Cache: Fehler - {e}"
        set_status(msg, "salmon")
        write_log(msg)


def clear_store_cache():
    try:
        subprocess.Popen(["wsreset.exe"])
        msg = "Store Cache: wsreset.exe gestartet (laeuft im Hintergrund)"
        set_status(msg, "lightgreen")
        write_log(msg)
    except Exception as e:
        msg = f"Store Cache: Fehler - {e}"
        set_status(msg, "salmon")
        write_log(msg)


# -------------------------------------------------------
# UI Build
# -------------------------------------------------------
BG = "#1e1e2e"
FG = "#cdd6f4"
BTN_BG = "#313244"
BTN_FG = "#cdd6f4"
BTN_ACTIVE = "#45475a"
ACCENT = "#89b4fa"
GROUP_FG = "#89dceb"
FONT_MAIN = ("Segoe UI", 10)
FONT_TITLE = ("Segoe UI Semibold", 10)
FONT_GROUP = ("Segoe UI Semibold", 9)


def make_button(parent, text, command, tooltip):
    btn = tk.Button(
        parent,
        text=text,
        command=command,
        bg=BTN_BG,
        fg=BTN_FG,
        activebackground=BTN_ACTIVE,
        activeforeground=FG,
        relief="flat",
        font=FONT_MAIN,
        cursor="hand2",
        padx=10,
        pady=6,
        width=28,
    )
    btn.bind("<Enter>", lambda e: btn.config(bg=BTN_ACTIVE))
    btn.bind("<Leave>", lambda e: btn.config(bg=BTN_BG))
    add_tooltip(btn, tooltip)
    return btn


def make_group(parent, title):
    frame = tk.LabelFrame(
        parent,
        text=f"  {title}  ",
        bg=BG,
        fg=GROUP_FG,
        font=FONT_GROUP,
        bd=1,
        relief="groove",
        padx=10,
        pady=8,
        labelanchor="n",
    )
    return frame


root = tk.Tk()
root.title("Windows Cache Cleaner")
root.configure(bg=BG)
root.resizable(False, False)

# Title bar
title_bar = tk.Frame(root, bg=BG)
title_bar.pack(fill="x", padx=16, pady=(14, 4))
tk.Label(
    title_bar,
    text="Windows Cache Cleaner",
    bg=BG,
    fg=ACCENT,
    font=("Segoe UI Semibold", 14),
).pack(side="left")

if is_admin():
    tk.Label(title_bar, text="[Admin]", bg=BG, fg="#a6e3a1",
             font=("Segoe UI", 9)).pack(side="right", padx=4)
else:
    tk.Label(title_bar, text="[kein Admin - Prefetch deaktiviert]",
             bg=BG, fg="#f38ba8", font=("Segoe UI", 9)).pack(side="right", padx=4)

main = tk.Frame(root, bg=BG)
main.pack(padx=16, pady=6)

# --- Group 1: Dateisystem-Caches ---
g1 = make_group(main, "Dateisystem-Caches")
g1.grid(row=0, column=0, padx=8, pady=6, sticky="n")

buttons_fs = [
    ("Clear Recent Files", clear_recent_files,
     "Loescht den Inhalt von:\n%APPDATA%\\Microsoft\\Windows\\Recent\n\nEntfernt alle LNK-Verknuepfungen zu zuletzt geoeffneten Dateien."),
    ("Clear Jump Lists", clear_automatic_destinations,
     "Loescht den Inhalt von:\n%APPDATA%\\...\\Recent\\AutomaticDestinations\n\nEntfernt die Eintraege die beim Rechtsklick auf Taskleisten-Icons erscheinen."),
    ("Clear Thumbnail Cache", clear_thumbnail_cache,
     "Loescht thumbcache_*.db im Explorer-Ordner.\nExplorer wird kurz beendet und neu gestartet.\n\nSpeichert Dateivorschau-Bilder (Fotos, Videos, Dokumente)."),
    ("Clear Icon Cache", clear_icon_cache,
     "Loescht iconcache_*.db und IconCache.db.\nExplorer wird kurz beendet und neu gestartet.\n\nSpeichert App-Icons. Loesen: falsche oder fehlende Icons."),
    ("Clear Prefetch", clear_prefetch,
     "Loescht *.pf-Dateien aus C:\\Windows\\Prefetch.\nBENOETIGT ADMINISTRATOR-RECHTE.\n\nWindows-Vorladeoptimierung. Unbedenklich zu loeschen,\nWindows baut den Cache selbst neu auf."),
]

for i, (txt, cmd, tip) in enumerate(buttons_fs):
    btn = make_button(g1, txt, cmd, tip)
    btn.grid(row=i, column=0, pady=3, sticky="ew")

# --- Group 2: Registry-Caches ---
g2 = make_group(main, "Registry-Caches")
g2.grid(row=0, column=1, padx=8, pady=6, sticky="n")

buttons_reg = [
    ("Clear MUI Cache", clear_mui_cache,
     "Leert: HKCU\\Software\\Microsoft\\Windows\\ShellNoRoam\\MUICache\n\nSpeichert Anzeigenamen von ausgefuehrten Programmen.\nKann veraltete/geloeschte Eintraege enthalten."),
    ("Clear ShellBags", clear_shellbags,
     "Loescht ShellBag-Eintraege in HKCU\\...\\Shell\\Bags und BagMRU.\n\nWindows merkt sich Position und Ansicht JEDES je geoeffneten Ordners -\nauch geloeschter Ordner, USB-Sticks und Netzlaufwerke.\nAbmelden nach dem Loeschen empfohlen."),
    ("Clear RunMRU", clear_runmru,
     "Leert: HKCU\\...\\Explorer\\RunMRU\n\nSpeichert alle Befehle die du je im Win+R Dialog eingegeben hast."),
]

for i, (txt, cmd, tip) in enumerate(buttons_reg):
    btn = make_button(g2, txt, cmd, tip)
    btn.grid(row=i, column=0, pady=3, sticky="ew")

# --- Group 3: System ---
g3 = make_group(main, "System")
g3.grid(row=0, column=2, padx=8, pady=6, sticky="n")

buttons_sys = [
    ("Clear DNS Cache", clear_dns_cache,
     "Fuehrt 'ipconfig /flushdns' aus.\n\nLoescht den Windows-DNS-Cache.\nHilft bei: Seiten nicht erreichbar, alte IP-Adressen,\nnach DNS-Aenderungen."),
    ("Clear Store Cache", clear_store_cache,
     "Startet wsreset.exe im Hintergrund.\n\nLoescht den Microsoft Store Cache.\nHilft bei: Store startet nicht, Apps laden nicht,\nInstallationsfehler."),
]

for i, (txt, cmd, tip) in enumerate(buttons_sys):
    btn = make_button(g3, txt, cmd, tip)
    btn.grid(row=i, column=0, pady=3, sticky="ew")

# --- Status Bar ---
status_frame = tk.Frame(root, bg="#11111b", pady=6)
status_frame.pack(fill="x", padx=0, pady=(6, 0))

status_var = tk.StringVar(value="Bereit.")
status_label = tk.Label(
    status_frame,
    textvariable=status_var,
    bg="#11111b",
    fg="white",
    font=("Segoe UI", 9),
    anchor="w",
    padx=16,
)
status_label.pack(fill="x")

# --- Elevation hint ---
if not is_admin():
    hint = tk.Label(
        root,
        text="Tipp: Als Administrator ausfuehren um alle Funktionen zu nutzen (Rechtsklick -> Als Administrator ausfuehren)",
        bg=BG,
        fg="#fab387",
        font=("Segoe UI", 8),
        anchor="w",
        padx=16,
    )
    hint.pack(fill="x", pady=(0, 8))

root.mainloop()
