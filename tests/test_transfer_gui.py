"""GUI safety-state tests; every database/backup call is mocked."""

from pathlib import Path
import sys
import time
import tkinter as tk
import unittest
from unittest.mock import patch

from transfer_gui import TransferWindow
from operation_transfer import CommitOutcomeUnknown, digest


@unittest.skipUnless(sys.platform == "win32", "Requires Windows Tk display")
class TransferGuiTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.addCleanup(self.destroy)
        with patch("transfer_gui.read_manifest", return_value={"workbook": "Merged.xlsx"}):
            self.window = TransferWindow(self.root, Path("fixture.transfer.json"), {"test_only": True})
        self.window.window.withdraw()
        self.plan = {"quote": 102, "header": [102, "Merged", "Customer", 1], "already_applied": False,
                     "inserts": [{"table": "CROSSCUTS", "uid": 1, "cutlist": 1,
                                  "values": [0, 185, 0, 2, 0, "0"], "source_quote": 100}]}

    def destroy(self):
        self.root.update_idletasks()
        self.window.window.after_cancel(self.window.timer)
        for identifier in self.root.tk.call("after", "info"):
            self.root.after_cancel(identifier)
        self.root.destroy()

    def wait(self):
        deadline = time.monotonic() + 5
        while self.window.busy and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertFalse(self.window.busy)

    def checked(self):
        self.window.quote.set("102")
        with patch("transfer_gui.check_target", return_value=self.plan):
            self.window.check()
            self.wait()

    def ready(self):
        self.checked()
        self.window.receipt = "fixture-receipt.json"
        self.window.closed.set(True)
        self.window.update_actions()

    def test_apply_requires_review_backup_and_closed_confirmation(self):
        self.assertIn("disabled", self.window.apply_button.state())
        self.checked()
        self.assertIn("disabled", self.window.apply_button.state())
        self.window.receipt = "fixture-receipt.json"
        self.window.update_actions()
        self.assertIn("disabled", self.window.apply_button.state())
        self.window.closed.set(True)
        self.window.update_actions()
        self.assertNotIn("disabled", self.window.apply_button.state())

    def test_quote_edit_invalidates_review_and_backup(self):
        self.ready()
        self.window.quote.set("103")
        self.assertIsNone(self.window.plan)
        self.assertIsNone(self.window.receipt)
        self.assertFalse(self.window.closed.get())
        self.assertFalse(self.window.reviewed.get())
        self.assertEqual(self.window.table.get_children(), ())

    def test_declined_confirmation_never_calls_apply(self):
        self.ready()
        with patch("transfer_gui.messagebox.askyesno", return_value=False), patch("transfer_gui.apply_transfer") as apply:
            self.window.apply()
        apply.assert_not_called()

    def test_unknown_commit_clears_approval_and_requires_recheck(self):
        self.ready()
        with patch("transfer_gui.messagebox.askyesno", return_value=True), patch("transfer_gui.messagebox.showerror") as error, \
                patch("transfer_gui.apply_transfer", side_effect=CommitOutcomeUnknown("Check the outcome")) as apply:
            self.window.apply()
            self.wait()
        apply.assert_called_once()
        self.assertEqual(error.call_args.args[0], "Check the commit outcome")
        self.assertIsNone(self.window.plan)
        self.assertIsNone(self.window.receipt)
        self.assertIn("disabled", self.window.apply_button.state())

    def test_audit_failure_is_not_reported_as_rollback(self):
        self.ready()
        with patch("transfer_gui.messagebox.askyesno", return_value=True), patch("transfer_gui.messagebox.showwarning") as warning, \
                patch("transfer_gui.apply_transfer", return_value="committed_audit_pending"):
            self.window.apply()
            self.wait()
        warning.assert_called_once()
        self.assertIn("committed", self.window.status.get())

    def test_busy_transfer_cannot_be_closed(self):
        self.window.busy = True
        with patch("transfer_gui.messagebox.showinfo"):
            self.window.close()
        self.assertTrue(self.window.window.winfo_exists())

    def test_import_differences_require_separate_confirmation(self):
        self.plan["review_differences"] = [{"field": "Job name", "source": "Merged", "target": "TEST - MERGE"}]
        self.ready()
        self.assertIn("TEST - MERGE", self.window.difference_summary.get())
        self.assertIn("disabled", self.window.apply_button.state())
        with patch("transfer_gui.apply_transfer") as apply:
            self.window.apply()
        apply.assert_not_called()
        self.window.reviewed.set(True)
        self.window.update_actions()
        self.assertNotIn("disabled", self.window.apply_button.state())
        with patch("transfer_gui.apply_transfer", return_value="committed") as apply, \
                patch("transfer_gui.messagebox.askyesno", return_value=True), patch("transfer_gui.messagebox.showinfo"):
            self.window.apply()
            self.wait()
        self.assertEqual(apply.call_args.kwargs["confirmed_review_hash"],
                         digest(self.plan["review_differences"]))


if __name__ == "__main__":
    unittest.main()
