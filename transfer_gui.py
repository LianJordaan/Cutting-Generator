"""Review, backup and explicit approval for a staged extra-operation transfer."""

from pathlib import Path
from queue import Empty, Queue
from threading import Thread
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from operation_transfer import (CommitOutcomeUnknown, apply_transfer, check_target,
                                create_verified_backup, read_manifest, OPERATIONS, digest)


class TransferWindow:
    def __init__(self, parent, manifest_path, config):
        self.manifest = read_manifest(Path(manifest_path), config)
        self.window = tk.Toplevel(parent)
        self.window.title("Transfer extra operations - Cutting Generator")
        self.window.geometry("1050x760")
        self.window.minsize(900, 680)
        self.window.protocol("WM_DELETE_WINDOW", self.close)
        self.path = Path(manifest_path)
        self.config = config
        self.plan = None
        self.receipt = None
        self.busy = False
        self.queue = Queue()
        self.quote = tk.StringVar(self.window)
        self.closed = tk.BooleanVar(self.window, False)
        self.reviewed = tk.BooleanVar(self.window, False)
        self.difference_summary = tk.StringVar(self.window)
        self.status = tk.StringVar(self.window, "Import the merged workbook into a NEW quote first, then enter that quote number.")
        self.target = tk.StringVar(self.window, "No target quote checked yet. No database changes have been made.")
        self.backup_status = tk.StringVar(self.window, "A backup restored successfully to a temporary test database is required.")
        body = ttk.Frame(self.window, padding=22)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(5, weight=1)
        ttk.Label(body, text="Finish importing your extra operations", style="Heading.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(body, text="1. Import the merged .xlsx into a new Cutting Manager quote.\n"
                  "2. Save and close Cutting Manager to avoid cached data replacing the additions.\n"
                  "3. Check the new quote, create a verified backup, then review and approve the additions.",
                  wraplength=950, justify="left").grid(row=1, column=0, sticky="ew", pady=(12, 15))
        ttk.Label(body, text=f"Workbook: {self.manifest['workbook']}", style="Muted.TLabel").grid(row=2, column=0, sticky="w")
        bar = ttk.Frame(body)
        bar.grid(row=3, column=0, sticky="ew", pady=14)
        ttk.Label(bar, text="New quote number").pack(side="left")
        self.entry = ttk.Entry(bar, textvariable=self.quote, width=14)
        self.entry.pack(side="left", padx=10)
        self.check_button = ttk.Button(bar, text="Check quote (read only)", command=self.check)
        self.check_button.pack(side="left")
        ttk.Label(body, textvariable=self.target, wraplength=950, justify="left").grid(row=4, column=0, sticky="ew", pady=(0, 12))
        table_frame = ttk.Frame(body)
        table_frame.grid(row=5, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        self.table = ttk.Treeview(table_frame, columns=("type", "uid", "cutlist", "details", "source"), show="headings")
        for field, label, width in [("type", "Operation", 135), ("uid", "Piece ID", 75),
                                     ("cutlist", "Cutlist", 65), ("details", "Exact stored values", 430),
                                     ("source", "Source quote", 100)]:
            self.table.heading(field, text=label)
            self.table.column(field, width=width)
        self.table.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(table_frame, command=self.table.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.table.configure(yscrollcommand=scroll.set)
        horizontal = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        self.table.configure(xscrollcommand=horizontal.set)
        backup = ttk.Frame(body)
        backup.grid(row=6, column=0, sticky="ew", pady=(14, 6))
        self.backup_button = ttk.Button(backup, text="Create & verify backup...", command=self.backup)
        self.backup_button.pack(side="left")
        ttk.Label(backup, textvariable=self.backup_status, style="Muted.TLabel", wraplength=650).pack(side="left", padx=12)
        review = ttk.Frame(body)
        review.grid(row=7, column=0, sticky="ew", pady=10)
        self.difference_label = ttk.Frame(review)
        self.difference_label.columnconfigure(0, weight=1)
        self.difference_text = tk.Text(self.difference_label, height=4, width=105, wrap="word",
                                       font=("Segoe UI", 9), background="white", relief="flat",
                                       padx=8, pady=6, state="disabled")
        self.difference_text.grid(row=0, column=0, sticky="ew")
        difference_scroll = ttk.Scrollbar(self.difference_label, command=self.difference_text.yview)
        difference_scroll.grid(row=0, column=1, sticky="ns")
        self.difference_text.configure(yscrollcommand=difference_scroll.set)
        self.difference_summary.trace_add("write", self.render_differences)
        self.reviewed_check = ttk.Checkbutton(review, text="I accept the displayed import differences for this quote.",
                                             variable=self.reviewed, command=self.update_actions)
        self.closed_check = ttk.Checkbutton(review, text="I have saved and closed Cutting Manager, and checked the target customer/job above.",
                                            variable=self.closed, command=self.update_actions)
        self.closed_check.pack(anchor="w", pady=(6, 0))
        footer = ttk.Frame(body)
        footer.grid(row=8, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, textvariable=self.status, style="Muted.TLabel", wraplength=650).grid(row=0, column=0, sticky="w")
        self.apply_button = ttk.Button(footer, text="Apply reviewed additions", style="Accent.TButton", command=self.apply)
        self.apply_button.grid(row=0, column=1, padx=(12, 0))
        self.progress = ttk.Progressbar(body, mode="indeterminate")
        self.progress.grid(row=9, column=0, sticky="ew", pady=(10, 0))
        self.progress.grid_remove()
        self.quote.trace_add("write", self.invalidate)
        self.timer = self.window.after(80, self.poll)
        self.update_actions()

    def render_differences(self, *args):
        self.difference_text.configure(state="normal")
        self.difference_text.delete("1.0", "end")
        self.difference_text.insert("1.0", self.difference_summary.get())
        self.difference_text.configure(state="disabled")

    def invalidate(self, *args):
        self.plan = None
        self.receipt = None
        self.closed.set(False)
        self.reviewed.set(False)
        self.difference_summary.set("")
        self.difference_label.pack_forget()
        self.reviewed_check.pack_forget()
        self.target.set("Quote number changed. Check the target again before continuing.")
        self.backup_status.set("A newly verified backup is required for this target.")
        self.table.delete(*self.table.get_children())
        self.update_actions()

    def update_actions(self):
        checked = self.plan is not None and not self.plan["already_applied"]
        for widget in [self.entry, self.check_button, self.closed_check, self.reviewed_check]:
            widget.configure(state="disabled" if self.busy else "normal")
        self.backup_button.configure(state="normal" if checked and not self.busy else "disabled")
        differences_accepted = not self.plan or not self.plan.get("review_differences") or self.reviewed.get()
        self.apply_button.configure(state="normal" if checked and self.receipt and self.closed.get() and differences_accepted and not self.busy else "disabled")

    def run(self, action, worker, message):
        if self.busy:
            return
        self.busy = True
        self.status.set(message)
        self.progress.grid()
        self.progress.start(12)
        self.update_actions()

        def background():
            try:
                self.queue.put((action, worker(), None))
            except Exception as exc:
                self.queue.put((action, None, exc))
        Thread(target=background, daemon=True).start()

    def check(self):
        try:
            number = int(self.quote.get())
            if number <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Quote number required", "Enter the new quote number shown by Cutting Manager.", parent=self.window)
            return
        self.plan = self.receipt = None
        self.closed.set(False)
        self.reviewed.set(False)
        self.difference_summary.set("")
        self.difference_label.pack_forget()
        self.reviewed_check.pack_forget()
        self.table.delete(*self.table.get_children())
        self.run("checked", lambda: check_target(self.path, number, self.config), "Checking the target and rechecking all source pieces (read only)...")

    def backup(self):
        folder = filedialog.askdirectory(parent=self.window, title="Choose a folder for the verified database backup")
        if not folder:
            return
        self.receipt = None
        self.run("backup", lambda: create_verified_backup(self.path, self.plan, self.config, folder),
                 "Creating a full logical backup and restoring it to an isolated temporary database. Please wait...")

    def apply(self):
        if not self.plan or not self.receipt or not self.closed.get() or self.busy:
            return
        differences = self.plan.get("review_differences", [])
        if differences and not self.reviewed.get():
            return
        review_hash = digest(differences) if differences else None
        number = self.plan["quote"]
        if not messagebox.askyesno("Approve extra-operation additions",
                                  f"Add {len(self.plan['inserts'])} operations to quote {number}?\n\n"
                                  f"Customer: {self.plan['header'][2]}\nJob: {self.plan['header'][1]}\n\n"
                                  "The database will be rechecked under locks. Existing records will not be replaced.",
                                  parent=self.window, default="no"):
            return
        self.run("applied", lambda: apply_transfer(self.path, self.plan, self.receipt, self.config,
                                                    confirmed_quote=number, cutting_manager_closed=True,
                                                    confirmed_review_hash=review_hash),
                 "Rechecking the database and applying the approved additions in one transaction...")

    def poll(self):
        try:
            while True:
                action, value, error = self.queue.get_nowait()
                self.busy = False
                self.progress.stop()
                self.progress.grid_remove()
                if error is not None:
                    self.receipt = None
                    if action in ("checked", "applied"):
                        self.plan = None
                    message = str(error)
                    self.status.set(message)
                    messagebox.showerror("Transfer stopped" if not isinstance(error, CommitOutcomeUnknown) else "Check the commit outcome",
                                         message, parent=self.window)
                elif action == "checked":
                    self.plan = value
                    differences = value.get("review_differences", [])
                    if differences:
                        lines = []
                        for difference in differences:
                            label = difference["field"] + (f" / CutList{difference['cutlist']}" if "cutlist" in difference else "")
                            original = "blank" if difference["source"] is None else str(difference["source"])
                            lines.append(f"{label}: {original} (source) -> {difference['target']} (imported)")
                        self.difference_summary.set("Import differences to review (piece data matches):\n" + "\n".join(lines))
                        self.difference_label.pack(fill="x", before=self.closed_check)
                        self.reviewed_check.pack(anchor="w", before=self.closed_check, pady=(5, 0))
                    self.target.set(f"Quote {value['quote']}  |  Customer: {value['header'][2]}  |  Job: {value['header'][1]}\n"
                                    f"{len(value['inserts'])} extra-operation records match the imported pieces.")
                    for item in value["inserts"]:
                        self.table.insert("", "end", values=(item["table"], item["uid"], item["cutlist"],
                                                               "; ".join(f"{field}={value!r}" for field, value in zip(OPERATIONS[item["table"]], item["values"])), item["source_quote"]))
                    self.status.set("All expected operations are already present; nothing needs to be added." if value["already_applied"] else
                                    "Target verified. Create a verified backup before approving additions.")
                elif action == "backup":
                    self.receipt = value
                    self.backup_status.set("Backup created and successfully restored for verification. Receipt: " + str(value))
                    self.status.set("Review the additions and confirm Cutting Manager is closed, then apply.")
                elif action == "applied":
                    self.plan = self.receipt = None
                    if value == "committed_audit_pending":
                        self.status.set("Operations committed, but the final audit receipt could not be saved. Check the quote before continuing.")
                        messagebox.showwarning("Operations committed - audit receipt pending", self.status.get(), parent=self.window)
                    else:
                        self.status.set("Transfer completed. Reopen Cutting Manager to review the added operations.")
                        messagebox.showinfo("Transfer complete", "The extra operations are present. No existing piece or source records were changed.\n\n"
                                            "Reopen Cutting Manager to review the target quote.", parent=self.window)
                self.update_actions()
        except Empty:
            pass
        if self.window.winfo_exists():
            self.timer = self.window.after(80, self.poll)

    def close(self):
        if self.busy:
            messagebox.showinfo("Operation in progress", "Please wait for the database/backup operation to finish.", parent=self.window)
        else:
            self.window.after_cancel(self.timer)
            self.window.destroy()


def show_transfer_window(parent, manifest_path, config):
    existing = getattr(parent, "_operation_transfer_window", None)
    if existing is not None and existing.window.winfo_exists():
        existing.window.lift()
        return existing
    try:
        window = TransferWindow(parent, manifest_path, config)
        parent._operation_transfer_window = window
        return window
    except Exception as exc:
        messagebox.showerror("Cannot open transfer record", str(exc), parent=parent)
        return None


def resume_transfer(manifest_path, config):
    root = tk.Tk()
    root.withdraw()
    # Reuse the merge window's shared style without creating a second workflow.
    from merge_gui import MergeWindow
    shell = MergeWindow.__new__(MergeWindow)
    shell.root = root
    shell._style()
    window = show_transfer_window(root, manifest_path, config)
    try:
        if window is not None:
            window.window.protocol("WM_DELETE_WINDOW", lambda: window.close() if window.busy else root.quit())
            root.mainloop()
    finally:
        root.destroy()
