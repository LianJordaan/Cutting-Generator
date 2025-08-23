import tkinter as tk
from tkinter import messagebox
from config_utils import save_config, load_config

def setup():
    """Open GUI to input config"""

    def on_submit():
        ip = ip_entry.get().strip()
        port = port_entry.get().strip()
        username = username_entry.get().strip()
        password = password_entry.get().strip()
        filepath = filepath_entry.get().strip() or "C:/ZAWare/DB/CutMan/CUTMAN.FDB"
        charset = charset_entry.get().strip() or "UTF8"

        if not ip or not username or not password:
            messagebox.showerror("Error", "IP, username and password are required!")
            return

        config_data = {
            "ip": ip,
            "port": port,
            "username": username,
            "password": password,
            "filepath": filepath,
            "charset": charset,
        }

        save_config(config_data)
        messagebox.showinfo("Success", "Configuration saved successfully!")
        root.destroy()

    root = tk.Tk()
    root.title("Setup Firebird DB Connection")
    root.geometry("400x300")

    # Load existing config if present
    existing_config = load_config()
    
    tk.Label(root, text="IP Address *").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    ip_entry = tk.Entry(root, width=30)
    ip_entry.grid(row=0, column=1)
    if existing_config:
        ip_entry.insert(0, existing_config.get("ip", ""))

    tk.Label(root, text="Port (optional)").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    port_entry = tk.Entry(root, width=30)
    port_entry.grid(row=1, column=1)
    if existing_config:
        port_entry.insert(0, existing_config.get("port", ""))

    tk.Label(root, text="Username *").grid(row=2, column=0, sticky="w", padx=5, pady=5)
    username_entry = tk.Entry(root, width=30)
    username_entry.grid(row=2, column=1)
    if existing_config:
        username_entry.insert(0, existing_config.get("username", ""))

    tk.Label(root, text="Password *").grid(row=3, column=0, sticky="w", padx=5, pady=5)
    password_entry = tk.Entry(root, width=30, show="*")
    password_entry.grid(row=3, column=1)
    if existing_config:
        password_entry.insert(0, existing_config.get("password", ""))

    tk.Label(root, text="File Path (optional)").grid(row=4, column=0, sticky="w", padx=5, pady=5)
    filepath_entry = tk.Entry(root, width=30)
    filepath_entry.grid(row=4, column=1)
    if existing_config:
        filepath_entry.insert(0, existing_config.get("filepath", "C:/ZAWare/DB/CutMan/CUTMAN.FDB"))
    else:
        filepath_entry.insert(0, "C:/ZAWare/DB/CutMan/CUTMAN.FDB")

    tk.Label(root, text="Charset (optional)").grid(row=5, column=0, sticky="w", padx=5, pady=5)
    charset_entry = tk.Entry(root, width=30)
    charset_entry.grid(row=5, column=1)
    if existing_config:
        charset_entry.insert(0, existing_config.get("charset", "UTF8"))
    else:
        charset_entry.insert(0, "UTF8")

    submit_btn = tk.Button(root, text="Save", command=on_submit)
    submit_btn.grid(row=6, column=0, columnspan=2, pady=15)

    root.mainloop()

def get_setup_info() -> dict:
    """Return the decrypted config dict, or None if missing"""
    from config_utils import load_config
    return load_config()
