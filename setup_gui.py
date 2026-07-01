import datetime
import tkinter as tk

from tkinter import messagebox

from config_utils import load_config, save_config
from device_auth import get_machine_guid
from helpers import open_license_link

OWNERSHIP_NOTICE = "Cutting Generator was sold to ByteBuilders Hosting and is now owned by ByteBuilders Hosting."


def copy_text_to_clipboard(window, value: str, success_message: str):
    window.clipboard_clear()
    window.clipboard_append(value)
    window.update()
    messagebox.showinfo("Copied", success_message)


def show_license_issue_window(machine_guid: str, message: str):
    dialog = tk.Tk()
    dialog.title("Cutting Generator Licence Required")
    dialog.geometry("620x220")
    dialog.resizable(False, False)

    tk.Label(
        dialog,
        text=message,
        wraplength=560,
        justify="left",
        fg="red",
    ).pack(anchor="w", padx=20, pady=(20, 10))

    tk.Label(dialog, text="Device ID (Machine GUID)").pack(anchor="w", padx=20)

    guid_frame = tk.Frame(dialog)
    guid_frame.pack(fill="x", padx=20, pady=(4, 8))

    guid_entry = tk.Entry(guid_frame, width=58)
    guid_entry.insert(0, machine_guid)
    guid_entry.configure(state="readonly")
    guid_entry.pack(side="left", fill="x", expand=True)

    tk.Button(
        guid_frame,
        text="Copy",
        width=10,
        command=lambda: copy_text_to_clipboard(dialog, machine_guid, "Device ID copied to clipboard."),
    ).pack(side="left", padx=(8, 0))

    tk.Label(
        dialog,
        text="Use this Device ID when adding or renewing this machine in the GitHub licence file.",
        wraplength=560,
        justify="left",
    ).pack(anchor="w", padx=20)

    tk.Label(
        dialog,
        text=OWNERSHIP_NOTICE,
        wraplength=560,
        justify="left",
    ).pack(anchor="w", padx=20, pady=(10, 0))

    tk.Button(dialog, text="Close", width=14, command=dialog.destroy).pack(pady=16)
    dialog.mainloop()


def setup():
    """Open GUI to input config"""

    def on_submit():
        ip = ip_entry.get().strip()
        port = port_entry.get().strip()
        username = username_entry.get().strip()
        password = password_entry.get().strip()
        developer_password = developer_password_entry.get().strip()
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
            if agree_terms:
                agree_time = datetime.datetime.now().isoformat()

        config_data = {
            "ip": ip,
            "port": port,
            "username": username,
            "password": password,
            "developer_password": developer_password,
            "filepath": filepath,
            "charset": charset,
            "agree_terms": agree_terms,
            "agree_time": agree_time,
        }

        save_config(config_data)
        messagebox.showinfo("Success", "Configuration saved successfully!")
        root.destroy()

    root = tk.Tk()
    root.title("Setup Firebird DB Connection (Same as Cutting Manager)")
    root.geometry("660x500")

    global existing_config
    existing_config = load_config()
    machine_guid = get_machine_guid()

    tk.Label(
        root,
        text=OWNERSHIP_NOTICE,
        wraplength=610,
        justify="left",
    ).grid(row=0, column=0, columnspan=3, sticky="w", padx=5, pady=(10, 12))

    tk.Label(root, text="IP Address *").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    ip_entry = tk.Entry(root, width=30)
    ip_entry.grid(row=1, column=1, sticky="w")
    if existing_config:
        ip_entry.insert(0, existing_config.get("ip", ""))

    tk.Label(root, text="Port (optional)").grid(row=2, column=0, sticky="w", padx=5, pady=5)
    port_entry = tk.Entry(root, width=30)
    port_entry.grid(row=2, column=1, sticky="w")
    if existing_config:
        port_entry.insert(0, existing_config.get("port", ""))

    tk.Label(root, text="Username *").grid(row=3, column=0, sticky="w", padx=5, pady=5)
    username_entry = tk.Entry(root, width=30)
    username_entry.grid(row=3, column=1, sticky="w")
    if existing_config:
        username_entry.insert(0, existing_config.get("username", ""))

    tk.Label(root, text="Password *").grid(row=4, column=0, sticky="w", padx=5, pady=5)
    password_entry = tk.Entry(root, width=30, show="*")
    password_entry.grid(row=4, column=1, sticky="w")
    if existing_config:
        password_entry.insert(0, existing_config.get("password", ""))

    tk.Label(root, text="Developer Password").grid(row=5, column=0, sticky="w", padx=5, pady=5)
    developer_password_entry = tk.Entry(root, width=30, show="*")
    developer_password_entry.grid(row=5, column=1, sticky="w")
    if existing_config:
        developer_password_entry.insert(0, existing_config.get("developer_password", ""))

    tk.Label(root, text="File Path (optional)").grid(row=6, column=0, sticky="w", padx=5, pady=5)
    filepath_entry = tk.Entry(root, width=30)
    filepath_entry.grid(row=6, column=1, sticky="w")
    if existing_config:
        filepath_entry.insert(0, existing_config.get("filepath", "C:/ZAWare/DB/CutMan/CUTMAN.FDB"))
    else:
        filepath_entry.insert(0, "C:/ZAWare/DB/CutMan/CUTMAN.FDB")

    tk.Label(root, text="Charset (optional)").grid(row=7, column=0, sticky="w", padx=5, pady=5)
    charset_entry = tk.Entry(root, width=30)
    charset_entry.grid(row=7, column=1, sticky="w")
    if existing_config:
        charset_entry.insert(0, existing_config.get("charset", "UTF8"))
    else:
        charset_entry.insert(0, "UTF8")

    tk.Label(root, text="Device ID (Machine GUID)").grid(row=8, column=0, sticky="w", padx=5, pady=5)
    machine_guid_entry = tk.Entry(root, width=48)
    machine_guid_entry.insert(0, machine_guid)
    machine_guid_entry.configure(state="readonly")
    machine_guid_entry.grid(row=8, column=1, sticky="w")
    tk.Button(
        root,
        text="Copy",
        width=10,
        command=lambda: copy_text_to_clipboard(root, machine_guid, "Device ID copied to clipboard."),
    ).grid(row=8, column=2, sticky="w", padx=5)

    tk.Label(
        root,
        text="Use this Device ID when creating or renewing the licence entry for this machine on GitHub.",
        wraplength=610,
        justify="left",
    ).grid(row=9, column=0, columnspan=3, sticky="w", padx=5, pady=(0, 10))

    agree_terms_entry = tk.IntVar(value=1 if existing_config and existing_config.get("agree_terms") else 0)

    agree_terms_check = tk.Checkbutton(
        root,
        text="I agree to the Terms of Service and that I am using this software at my own risk",
        variable=agree_terms_entry,
        wraplength=450,
        justify="left",
    )
    agree_terms_check.grid(row=10, column=0, columnspan=3, pady=10)

    license_link = tk.Label(root, text="View License", fg="blue", cursor="hand2")
    license_link.grid(row=11, column=0, columnspan=3, pady=(0, 10))
    license_link.bind("<Button-1>", lambda e: open_license_link())

    submit_btn = tk.Button(root, text="Save", command=on_submit, width=15)
    submit_btn.grid(row=12, column=0, columnspan=3, pady=10)

    root.mainloop()


def get_setup_info() -> dict:
    """Return the decrypted config dict, or None if missing"""
    from config_utils import load_config

    return load_config()
