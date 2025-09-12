import tkinter as tk
from tkinter import messagebox
from config_utils import save_config, load_config
from helpers import open_license_link
import datetime

def setup():
    """Open GUI to input config"""

    def on_submit():
        ip = ip_entry.get().strip()
        port = port_entry.get().strip()
        username = username_entry.get().strip()
        password = password_entry.get().strip()
        filepath = filepath_entry.get().strip() or "C:/ZAWare/DB/CutMan/CUTMAN.FDB"
        charset = charset_entry.get().strip() or "UTF8"
        agree_terms = agree_terms_entry.get()

        if not ip or not username or not password:
            messagebox.showerror("Error", "IP, username and password are required!")
            return
        
        agree_time = None
        if existing_config and "agree_time" in existing_config:
            agree_time = existing_config["agree_time"]
        else:
            if agree_terms:  # only set time if they agreed
                agree_time = datetime.now().isoformat()

        config_data = {
            "ip": ip,
            "port": port,
            "username": username,
            "password": password,
            "filepath": filepath,
            "charset": charset,
            "agree_terms": agree_terms,
            "agree_time": agree_time
        }

        save_config(config_data)
        messagebox.showinfo("Success", "Configuration saved successfully!")
        root.destroy()

    root = tk.Tk()
    root.title("Setup Firebird DB Connection (Same as Cutting Manager)")
    root.geometry("500x350")

    # Load existing config if present
    global existing_config
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

    # Checkbox variable
    agree_terms_entry = tk.IntVar(value=1 if existing_config and existing_config.get("agree_terms") else 0)

    # Checkbox
    agree_terms_check = tk.Checkbutton(
        root,
        text="I agree to the Terms of Service and that I am using this software at my own risk",
        variable=agree_terms_entry,
        wraplength=450,
        justify="left"
    )
    agree_terms_check.grid(row=6, column=0, columnspan=2, pady=10)

    # Optional: clickable license link
    license_link = tk.Label(root, text="View License", fg="blue", cursor="hand2")
    license_link.grid(row=7, column=0, columnspan=2, pady=(0, 10))
    license_link.bind("<Button-1>", lambda e: open_license_link())

    # Save button below everything
    submit_btn = tk.Button(root, text="Save", command=on_submit, width=15)
    submit_btn.grid(row=8, column=0, columnspan=2, pady=10)

    root.mainloop()
def get_setup_info() -> dict:
    """Return the decrypted config dict, or None if missing"""
    from config_utils import load_config
    return load_config()
