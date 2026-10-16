from pathlib import Path
import sys
from tempfile import TemporaryDirectory
import time
import tkinter as tk
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from merge_gui import MergeWindow
from merge_processor import read_merge_file
from test_merge_processor import make_export


@unittest.skipUnless(sys.platform == "win32", "Requires a Windows Tk display")
class MergeGuiTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.first = self.folder / "First.xlsx"
        self.second = self.folder / "Second.xlsx"
        make_export(self.first, board="White")
        make_export(self.second, job="Second", board="Grey")
        self.root = tk.Tk()
        self.root.withdraw()
        self.addCleanup(self.destroy)
        self.window = MergeWindow(self.root)
        files = [read_merge_file(self.first), read_merge_file(self.second)]
        self.window.file_table.set_items(files, self.window.describe_file)
        self.window.files_changed()

    def destroy(self):
        if self.root.winfo_exists():
            self.root.update_idletasks()
            for identifier in self.root.tk.call("after", "info"):
                self.root.after_cancel(identifier)
            self.root.destroy()

    def wait_for_worker(self):
        deadline = time.monotonic() + 10
        while self.window.busy and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertFalse(self.window.busy, "Worker did not finish")

    def test_file_order_and_final_index_order_are_independent(self):
        window = self.window
        window.file_table.tree.selection_set("1")
        window.file_table.move(-1)
        window.next_step()
        self.assertEqual([g.board for g in window.group_table.items], ["Grey", "White"])
        window.group_table.tree.selection_set("1")
        window.group_table.move(-1)
        self.assertEqual([g.board for g in window.group_table.items], ["White", "Grey"])
        self.assertEqual(window.group_table.tree.item("0", "values")[0], "1")
        self.assertEqual(window.file_table.items[0].path, self.second)

    def test_back_preserves_final_order_and_user_details_when_files_unchanged(self):
        window = self.window
        window.next_step()
        window.group_table.tree.selection_set("1")
        window.group_table.move(-1)
        window.customer.set("Edited Customer")
        window.job.set("Edited Job")
        window.show_files()
        window.next_step()
        self.assertEqual([g.board for g in window.group_table.items], ["Grey", "White"])
        self.assertEqual((window.customer.get(), window.job.get()), ("Edited Customer", "Edited Job"))

    def test_duplicates_are_ignored(self):
        window = self.window
        window.load_paths([self.first, self.second])
        self.assertFalse(window.busy)
        self.assertEqual(len(window.file_table.items), 2)
        self.assertIn("no duplicates", window.status.get())

    def test_dragging_a_row_changes_order(self):
        table = self.window.file_table
        table.tree.selection_set("0")
        with patch.object(table.tree, "identify_region", return_value="cell"), patch.object(table.tree, "identify_row", side_effect=["0", "1"]):
            table.drag_start(SimpleNamespace(x=20, y=40))
            table.drag_end(SimpleNamespace(x=20, y=80))
        self.assertEqual([f.path for f in table.items], [self.second, self.first])

    def test_missing_details_disable_save(self):
        self.window.next_step()
        self.window.customer.set("")
        self.assertIn("disabled", self.window.next_button.state())
        self.window.customer.set("Customer")
        self.assertNotIn("disabled", self.window.next_button.state())

    def test_cancelling_save_does_not_create_output(self):
        self.window.next_step()
        with patch("merge_gui.filedialog.asksaveasfilename", return_value=""):
            self.window.save()
        self.assertFalse(self.window.busy)
        self.assertEqual(len(list(self.folder.iterdir())), 2)

    def test_cannot_save_over_input_file(self):
        self.window.next_step()
        with patch("merge_gui.filedialog.asksaveasfilename", return_value=str(self.first)), patch("merge_gui.messagebox.showerror") as error:
            self.window.save()
        error.assert_called_once()
        self.assertFalse(self.window.busy)

    def test_save_worker_creates_selected_order_and_success_page(self):
        window = self.window
        window.next_step()
        window.group_table.tree.selection_set("1")
        window.group_table.move(-1)
        window.job.set("Customer chosen job")
        path = self.folder / "Saved.xlsx"
        with patch("merge_gui.filedialog.asksaveasfilename", return_value=str(path)):
            window.save()
            self.wait_for_worker()
        self.assertEqual(window.result, path)
        self.assertEqual(window.page, 3)
        workbook = load_workbook(path)
        try:
            self.assertEqual(workbook.active["E4"].value, "Grey")
            self.assertEqual(workbook.active["C3"].value, "Customer chosen job")
        finally:
            workbook.close()

    def test_invalid_added_file_keeps_existing_files_and_shows_error(self):
        with patch("merge_gui.messagebox.showerror") as error:
            self.window.load_paths([self.folder / "bad.xls"])
            self.wait_for_worker()
        error.assert_called_once()
        self.assertEqual(len(self.window.file_table.items), 2)

    def enable_discovery(self, operations=True, errors=None):
        self.window.config = {"test_only": True}
        result = {"sources": [{"mapping": [{"operations": {"CROSSCUTS": [[0,185,0,0,0,"0"]] if operations else []}}]}],
                  "errors": errors or {}}
        self.window.inspection_finished(result, None)
        return result

    def test_detected_operations_offer_transfer_and_unique_ids(self):
        self.enable_discovery()
        with patch("merge_gui.messagebox.askyesno", return_value=True) as ask:
            self.window.next_step()
        ask.assert_called_once()
        self.assertTrue(self.window.include_operations.get())
        self.assertEqual(self.window.page, 2)

    def test_unknown_sources_require_explicit_workbook_only_choice(self):
        self.enable_discovery(errors={str(self.first): "quote missing"})
        with patch("merge_gui.messagebox.askyesno", return_value=False):
            self.window.next_step()
        self.assertEqual(self.window.page, 1)
        self.assertIn("disabled", self.window.operations_check.state())
        with patch("merge_gui.messagebox.askyesno", return_value=True):
            self.window.next_step()
        self.assertEqual(self.window.page, 2)
        self.assertFalse(self.window.include_operations.get())

    def test_readonly_discovery_runs_in_worker_and_reports_offline(self):
        self.window.config = {"test_only": True}
        with patch("merge_gui.inspect_sources", side_effect=OSError("offline")) as inspect:
            self.window.inspect()
            self.wait_for_worker()
        inspect.assert_called_once()
        self.assertIn("could not be verified", self.window.operation_status.get())
        self.assertIn("disabled", self.window.operations_check.state())

    def test_transfer_save_cannot_replace_existing_workbook(self):
        self.enable_discovery()
        with patch("merge_gui.messagebox.askyesno", return_value=True):
            self.window.next_step()
        path = self.folder / "existing.xlsx"
        path.write_bytes(b"keep")
        with patch("merge_gui.filedialog.asksaveasfilename", return_value=str(path)), patch("merge_gui.messagebox.showerror") as error:
            self.window.save()
        error.assert_called_once()
        self.assertEqual(path.read_bytes(), b"keep")

    def test_main_window_cannot_close_during_transfer(self):
        self.root._operation_transfer_window = SimpleNamespace(busy=True)
        with patch("merge_gui.messagebox.showinfo") as info:
            self.window.close()
        info.assert_called_once()
        self.assertTrue(self.root.winfo_exists())


if __name__ == "__main__":
    unittest.main()
