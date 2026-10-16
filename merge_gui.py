"""A two-step, keyboard-accessible Cut Manager merge window."""

import os
from pathlib import Path
from queue import Empty, Queue
import re
from threading import Thread
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from merge_processor import group_cutlists, merge_cutlists, move_items, read_merge_file, text


BG = "#f3f6fa"
INK = "#14283f"
MUTED = "#52657b"
ACCENT = "#087e8b"


class OrderTable(ttk.Frame):
    """Shared file/cutlist ordering control with buttons, keys and row dragging."""

    def __init__(self, parent, columns, on_change):
        super().__init__(parent)
        self.items = []
        self.describe = None
        self.on_change = on_change
        self.dragging = None
        self.enabled = True
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        names = [name for name, label, width in columns]
        self.tree = ttk.Treeview(self, columns=names, show="headings", selectmode="extended", height=10)
        for name, label, width in columns:
            self.tree.heading(name, text=label)
            self.tree.column(name, width=width, minwidth=45, stretch=name not in {"index", "counts"})
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        horizontal = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=horizontal.set)
        self.tree.tag_configure("even", background="#f4f8fb")
        controls = ttk.Frame(self)
        controls.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        self.up = ttk.Button(controls, text="Move up", command=lambda: self.move(-1))
        self.up.pack(side="left")
        self.down = ttk.Button(controls, text="Move down", command=lambda: self.move(1))
        self.down.pack(side="left", padx=(8, 0))
        ttk.Label(controls, text="Drag rows to reorder  |  Alt + Up / Down", style="Muted.TLabel").pack(side="right")
        self.tree.bind("<<TreeviewSelect>>", lambda event: self.update_buttons())
        self.tree.bind("<Alt-Up>", lambda event: self.move(-1))
        self.tree.bind("<Alt-Down>", lambda event: self.move(1))
        self.tree.bind("<ButtonPress-1>", self.drag_start, add=True)
        self.tree.bind("<B1-Motion>", self.drag_motion)
        self.tree.bind("<ButtonRelease-1>", self.drag_end, add=True)
        self.update_buttons()

    def set_items(self, items, describe, selected=()):
        self.items = list(items)
        self.describe = describe
        self.tree.delete(*self.tree.get_children())
        for index, item in enumerate(self.items):
            self.tree.insert("", "end", iid=str(index), values=(index + 1, *describe(item)),
                             tags=("even",) if index % 2 == 0 else ())
        if selected:
            self.tree.selection_set([str(i) for i in selected])
            self.tree.focus(str(selected[0]))
            self.tree.see(str(selected[0]))
        self.update_buttons()

    def selected(self):
        return sorted(int(i) for i in self.tree.selection())

    def update_buttons(self):
        selected = self.selected()
        self.up.configure(state="normal" if self.enabled and any(i > 0 and i - 1 not in selected for i in selected) else "disabled")
        self.down.configure(state="normal" if self.enabled and any(i < len(self.items) - 1 and i + 1 not in selected for i in selected) else "disabled")

    def set_enabled(self, enabled):
        self.enabled = enabled
        self.update_buttons()

    def move(self, direction):
        if self.enabled:
            result, selected = move_items(self.items, self.selected(), direction)
            self.set_items(result, self.describe, selected)
            self.on_change()
        return "break"

    def drag_start(self, event):
        self.dragging = self.tree.identify_row(event.y) if self.tree.identify_region(event.x, event.y) == "cell" else None

    def drag_motion(self, event):
        if self.enabled and self.dragging:
            self.tree.configure(cursor="hand2")
            if event.y < 28:
                self.tree.yview_scroll(-1, "units")
            elif event.y > self.tree.winfo_height() - 24:
                self.tree.yview_scroll(1, "units")
            return "break"

    def drag_end(self, event):
        start, self.dragging = self.dragging, None
        self.tree.configure(cursor="")
        target = self.tree.identify_row(event.y)
        if not self.enabled or not start or not target or start == target:
            return
        selected = self.selected() or [int(start)]
        if int(target) in selected:
            return
        moving = [self.items[i] for i in selected]
        remaining = [item for i, item in enumerate(self.items) if i not in selected]
        position = int(target) - sum(i < int(target) for i in selected)
        if int(target) > int(start):
            position += 1
        result = remaining[:position] + moving + remaining[position:]
        self.set_items(result, self.describe, list(range(position, position + len(moving))))
        self.on_change()


class MergeWindow:
    def __init__(self, root, paths=()):
        self.root = root
        self.result = None
        self.busy = False
        self.page = 1
        self.queue = Queue()
        self.preview_signature = None
        self.customer = tk.StringVar(root)
        self.job = tk.StringVar(root)
        self.contact = tk.StringVar(root)
        self.status = tk.StringVar(root, "Add at least two Cut Manager .xlsx exports to begin.")
        self.summary = tk.StringVar(root)
        self.default_details_set = False
        self._style()
        root.title("Merge cutting lists - Cutting Generator")
        root.geometry("1120x780")
        root.minsize(960, 700)
        root.configure(background=BG)
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.bind("<Escape>", lambda event: self.close())
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)
        self._header()
        self.content = ttk.Frame(root, padding=(26, 18))
        self.content.grid(row=1, column=0, sticky="nsew")
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(0, weight=1)
        self.files_page = ttk.Frame(self.content)
        self.preview_page = ttk.Frame(self.content)
        self._files_page()
        self._preview_page()
        self._footer()
        self.show_files()
        root.after(80, self.poll)
        if paths:
            root.after(120, lambda: self.load_paths(paths))

    def _style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 10), background=BG, foreground=INK)
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG)
        style.configure("Muted.TLabel", foreground=MUTED)
        style.configure("Heading.TLabel", font=("Segoe UI Semibold", 19))
        style.configure("TButton", padding=(14, 9), background="#e4ebf2", borderwidth=0)
        style.map("TButton", background=[("active", "#d6e2ec")], foreground=[("disabled", "#8c9aaa")])
        style.configure("Accent.TButton", background=ACCENT, foreground="white", font=("Segoe UI Semibold", 10))
        style.map("Accent.TButton", background=[("disabled", "#c4d3d9"), ("active", "#066671")], foreground=[("disabled", "#52657b")])
        style.configure("TEntry", fieldbackground="white", padding=8)
        style.configure("Treeview", background="white", fieldbackground="white", rowheight=38, borderwidth=0)
        style.configure("Treeview.Heading", background="#e4ebf2", foreground=INK,
                        font=("Segoe UI Semibold", 10), padding=(8, 10))
        style.map("Treeview", background=[("selected", "#d0edf0")], foreground=[("selected", "#123e46")])

    def _header(self):
        header = tk.Frame(self.root, background=INK, padx=26, pady=18)
        header.grid(row=0, column=0, sticky="ew")
        tk.Label(header, text="Merge cutting lists", font=("Segoe UI Semibold", 23),
                 background=INK, foreground="white").pack(anchor="w")
        tk.Label(header, text="Arrange your files. Combine matching boards. Set your cutting order.",
                 font=("Segoe UI", 11), background=INK, foreground="#c7d6e5").pack(anchor="w", pady=(5, 12))
        self.steps = tk.Label(header, font=("Segoe UI Semibold", 11), background=INK, foreground="#7bdee1")
        self.steps.pack(anchor="w")

    def _files_page(self):
        page = self.files_page
        page.columnconfigure(0, weight=1)
        page.rowconfigure(3, weight=1)
        ttk.Label(page, text="Choose the order of your files", style="Heading.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(page, text="This sets the order of pieces within matching boards. Original files stay untouched.",
                  style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 14))
        toolbar = ttk.Frame(page)
        toolbar.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        self.add_button = ttk.Button(toolbar, text="+ Add files", command=self.add_files)
        self.add_button.pack(side="left")
        self.remove_button = ttk.Button(toolbar, text="Remove selected", command=self.remove_files)
        self.remove_button.pack(side="left", padx=8)
        ttk.Label(toolbar, text="Cut Manager exports (.xlsx)", style="Muted.TLabel").pack(side="right")
        self.file_table = OrderTable(page, [("index", "Order", 55), ("file", "File", 400),
                                            ("job", "Source job", 240), ("counts", "Cutlists / pieces", 140)], self.files_changed)
        self.file_table.grid(row=3, column=0, sticky="nsew")
        self.file_table.set_items([], self.describe_file)
        ttk.Label(page, textvariable=self.summary, style="Muted.TLabel").grid(row=4, column=0, sticky="w", pady=(14, 0))

    def _preview_page(self):
        page = self.preview_page
        page.columnconfigure(0, weight=1)
        page.rowconfigure(3, weight=1)
        ttk.Label(page, text="Set the final cutting order", style="Heading.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(page, text="Matching boards are combined. Top to bottom becomes CutList1, CutList2, and so on.",
                  style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(5, 14))
        details = ttk.Frame(page)
        details.grid(row=2, column=0, sticky="ew", pady=(0, 16))
        self.detail_entries = []
        for column, (label, variable) in enumerate([("Customer *", self.customer), ("Merged job name *", self.job), ("Contact (optional)", self.contact)]):
            details.columnconfigure(column, weight=1)
            ttk.Label(details, text=label).grid(row=0, column=column, sticky="w", padx=(0, 14))
            entry = ttk.Entry(details, textvariable=variable)
            entry.grid(row=1, column=column, sticky="ew", padx=(0, 14), pady=(5, 0))
            self.detail_entries.append(entry)
            variable.trace_add("write", lambda *args: self.update_actions())
        self.group_table = OrderTable(page, [("index", "Index", 55), ("board", "Board", 265),
                                             ("category", "Category / thickness", 200), ("sources", "Sources", 100),
                                             ("counts", "Rows / pieces", 120)], self.groups_changed)
        self.group_table.grid(row=3, column=0, sticky="nsew")
        self.group_table.tree.bind("<<TreeviewSelect>>", lambda event: self.show_group_details(), add=True)
        self.source_details = tk.StringVar(self.root, "Select a board to see its source cutlists.")
        source_frame = ttk.Frame(page)
        source_frame.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        source_frame.columnconfigure(0, weight=1)
        self.source_text = tk.Text(source_frame, height=3, wrap="word", font=("Segoe UI", 9),
                                   background=BG, foreground=MUTED, relief="flat", highlightthickness=0)
        self.source_text.grid(row=0, column=0, sticky="ew")
        source_scroll = ttk.Scrollbar(source_frame, command=self.source_text.yview)
        source_scroll.grid(row=0, column=1, sticky="ns")
        self.source_text.configure(yscrollcommand=source_scroll.set)
        self.source_details.trace_add("write", self.render_sources)
        self.render_sources()
        ttk.Label(page, text="Different thicknesses or cutting settings stay separate. Edging, holes, remarks and unit IDs are preserved.",
                  style="Muted.TLabel", wraplength=970).grid(row=5, column=0, sticky="w", pady=(8, 0))

    def _footer(self):
        footer = ttk.Frame(self.root, padding=(26, 14))
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, textvariable=self.status, style="Muted.TLabel", wraplength=600).grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(footer, mode="indeterminate", length=160)
        self.progress.grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.progress.grid_remove()
        self.back_button = ttk.Button(footer, text="Back", command=self.show_files)
        self.back_button.grid(row=0, column=1, padx=(10, 8))
        self.cancel_button = ttk.Button(footer, text="Cancel", command=self.close)
        self.cancel_button.grid(row=0, column=2, padx=(0, 8))
        self.next_button = ttk.Button(footer, text="Review cutlists", style="Accent.TButton", command=self.next_step)
        self.next_button.grid(row=0, column=3)

    @staticmethod
    def describe_file(file):
        return file.path.name, file.job, f"{len(file.cutlists)} / {sum(s.quantity for s in file.cutlists)}"

    @staticmethod
    def describe_group(group):
        first = group.sources[0].sheet
        return group.board, f"{text(first['C4'].value)} / {text(first['J4'].value)} mm", len(group.sources), f"{group.row_count} / {group.quantity}"

    def show_files(self):
        if self.busy:
            return
        self.page = 1
        self.preview_page.grid_remove()
        self.files_page.grid(row=0, column=0, sticky="nsew")
        self.steps.configure(text="1  FILE ORDER     /     2  CUTLIST ORDER & SAVE")
        self.back_button.grid_remove()
        self.next_button.configure(text="Review cutlists")
        self.files_changed()

    def files_changed(self):
        files = self.file_table.items
        cutlists = sum(len(f.cutlists) for f in files)
        pieces = sum(s.quantity for f in files for s in f.cutlists)
        self.summary.set(f"{len(files)} files  |  {cutlists} source cutlists  |  {pieces} pieces")
        self.status.set("Ready to review the combined boards." if len(files) >= 2 else "Add at least two Cut Manager .xlsx exports to begin.")
        self.update_actions()

    def groups_changed(self):
        self.status.set(f"{len(self.group_table.items)} final cutlists. Indices update automatically when you move a row.")
        self.show_group_details()

    def show_group_details(self):
        selected = self.group_table.selected()
        self.group_table.update_buttons()
        if not selected:
            self.source_details.set("Select a board to see its source cutlists.")
            return
        group = self.group_table.items[selected[0]]
        sources = [f"{s.path.name} / {s.sheet.title} ({len(s.rows)} rows)" for s in group.sources]
        self.source_details.set("Piece order:\n" + "\n".join(f"{index}. {source}" for index, source in enumerate(sources, 1)))

    def render_sources(self, *args):
        self.source_text.configure(state="normal")
        self.source_text.delete("1.0", "end")
        self.source_text.insert("1.0", self.source_details.get())
        self.source_text.configure(state="disabled")

    def update_actions(self):
        if not hasattr(self, "next_button"):
            return
        ready = len(self.file_table.items) >= 2 if self.page == 1 else bool(self.group_table.items and text(self.customer.get()) and text(self.job.get()))
        self.next_button.configure(state="normal" if ready and not self.busy else "disabled")
        for button in [self.add_button, self.remove_button, self.back_button]:
            button.configure(state="disabled" if self.busy else "normal")
        self.file_table.set_enabled(not self.busy)
        self.group_table.set_enabled(not self.busy)
        for entry in self.detail_entries:
            entry.configure(state="disabled" if self.busy else "normal")

    def add_files(self):
        paths = filedialog.askopenfilenames(parent=self.root, title="Add Cut Manager exports",
                                          filetypes=[("Cut Manager workbooks", "*.xlsx")])
        if paths:
            self.load_paths(paths)

    def remove_files(self):
        if self.busy:
            return
        selected = self.file_table.selected()
        self.file_table.set_items([f for i, f in enumerate(self.file_table.items) if i not in selected], self.describe_file)
        self.files_changed()

    def set_busy(self, busy, message=""):
        self.busy = busy
        if busy:
            self.status.set(message)
            self.progress.grid()
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.grid_remove()
        self.update_actions()

    def load_paths(self, paths):
        if self.busy:
            return
        known = {f.path for f in self.file_table.items}
        fresh = []
        for path in paths:
            path = Path(path).resolve()
            if path not in known:
                fresh.append(path)
                known.add(path)
        if not fresh:
            self.status.set("These files are already in the list; no duplicates were added.")
            return
        self.set_busy(True, "Reading and checking workbooks...")

        def worker():
            loaded, errors = [], []
            for path in fresh:
                try:
                    loaded.append(read_merge_file(path))
                except Exception as exc:
                    errors.append(f"{path.name}: {exc}")
            self.queue.put(("loaded", loaded, errors))

        Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            while True:
                action, value, errors = self.queue.get_nowait()
                self.set_busy(False)
                if action == "loaded":
                    self.file_table.set_items(self.file_table.items + value, self.describe_file)
                    self.files_changed()
                    if errors:
                        messagebox.showerror("Some files could not be added", "\n\n".join(errors), parent=self.root)
                elif action == "saved":
                    self.saved(value)
                elif action == "failed":
                    self.status.set("The merge could not be saved. Your input files have not changed.")
                    messagebox.showerror("Could not save merged workbook", str(value), parent=self.root)
        except Empty:
            pass
        if self.root.winfo_exists():
            self.root.after(80, self.poll)

    def next_step(self):
        if self.busy:
            return
        if self.page == 2:
            self.save()
            return
        files = self.file_table.items
        if len(files) < 2:
            return
        signature = tuple((f.path, id(f)) for f in files)
        if self.preview_signature != signature:
            self.group_table.set_items(group_cutlists(files), self.describe_group)
            self.preview_signature = signature
        if not self.default_details_set:
            first = files[0].cutlists[0].sheet
            self.customer.set(text(first["C2"].value))
            self.job.set("MERGED - " + text(first["C3"].value))
            self.contact.set(text(first["J2"].value))
            self.default_details_set = True
        self.page = 2
        self.files_page.grid_remove()
        self.preview_page.grid(row=0, column=0, sticky="nsew")
        self.steps.configure(text="1  FILE ORDER     /     2  CUTLIST ORDER & SAVE  •  CURRENT STEP")
        self.back_button.grid()
        self.next_button.configure(text="Save merged workbook")
        self.groups_changed()
        self.update_actions()

    def save(self):
        customer, job, contact = self.customer.get().strip(), self.job.get().strip(), self.contact.get().strip()
        if not customer or not job:
            messagebox.showerror("Details required", "Enter a customer and merged job name.", parent=self.root)
            return
        filename = re.sub(r'[\\/:*?"<>|]', "-", job).strip(" .")[:120] or "Merged cutlists"
        path = filedialog.asksaveasfilename(parent=self.root, title="Save merged Cut Manager workbook",
                                          initialdir=self.file_table.items[0].path.parent,
                                          initialfile=filename + ".xlsx", defaultextension=".xlsx",
                                          filetypes=[("Cut Manager workbook", "*.xlsx")], confirmoverwrite=True)
        if not path:
            return
        if Path(path).resolve() in {f.path for f in self.file_table.items}:
            messagebox.showerror("Choose a new filename", "The output cannot replace an input file.", parent=self.root)
            return
        groups = list(self.group_table.items)
        overwrite = Path(path).exists()
        self.set_busy(True, "Saving merged workbook...")

        def worker():
            try:
                output = merge_cutlists(groups, path, customer=customer, job=job, contact=contact, overwrite=overwrite)
                self.queue.put(("saved", output, None))
            except Exception as exc:
                self.queue.put(("failed", exc, None))

        Thread(target=worker, daemon=True).start()

    def saved(self, path):
        self.result = path
        self.page = 3
        self.files_page.grid_remove()
        self.preview_page.grid_remove()
        self.steps.configure(text="MERGE COMPLETE")
        success = ttk.Frame(self.content, padding=(24, 34))
        success.grid(row=0, column=0, sticky="nsew")
        ttk.Label(success, text="Your merged workbook is ready", style="Heading.TLabel").pack(anchor="w")
        pieces = sum(g.quantity for g in self.group_table.items)
        ttk.Label(success, text=f"{len(self.file_table.items)} files combined into {len(self.group_table.items)} cutlists, with {pieces} pieces.",
                  style="Muted.TLabel").pack(anchor="w", pady=(12, 24))
        ttk.Label(success, text="Saved workbook", font=("Segoe UI Semibold", 11)).pack(anchor="w")
        ttk.Label(success, text=str(path), wraplength=900).pack(anchor="w", pady=(8, 24))
        ttk.Label(success, text="Import this .xlsx workbook into Cutting Manager. Your original files are unchanged.",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 20))
        buttons = ttk.Frame(success)
        buttons.pack(anchor="w")
        ttk.Button(buttons, text="Open containing folder", command=lambda: os.startfile(str(path.parent))).pack(side="left")
        ttk.Button(buttons, text="Copy file path", command=lambda: self.copy_path(path)).pack(side="left", padx=8)
        self.back_button.grid_remove()
        self.cancel_button.grid_remove()
        self.next_button.configure(text="Done", command=self.root.destroy, state="normal")
        self.status.set("Saved successfully. Ready to import into Cutting Manager.")

    def copy_path(self, path):
        self.root.clipboard_clear()
        self.root.clipboard_append(str(path))
        self.root.update()
        self.status.set("File path copied to the clipboard.")

    def close(self):
        if self.busy:
            messagebox.showinfo("Please wait", "A workbook operation is in progress. Please wait for it to finish.", parent=self.root)
            return
        self.root.destroy()


def show_merge_window(paths):
    root = tk.Tk()
    window = MergeWindow(root, paths)
    root.mainloop()
    return window.result
