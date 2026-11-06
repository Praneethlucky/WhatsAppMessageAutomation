import csv
import logging
import threading
import queue
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import pywhatkit

APP_TITLE = "WhatsApp Bulk Sender"
BROWSER_WAIT_SECONDS = 20  # Browser-loading wait required by PyWhatKit; not an inter-message delay.

class WhatsAppBulkSender:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("760x570")
        self.root.minsize(680, 500)

        self.file_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Choose an Excel workbook to begin.")
        self.progress_text = tk.StringVar(value="0 / 0")
        self.stop_event = threading.Event()
        self.worker = None
        self.events = queue.Queue()
        self.log_path = None

        self._build_ui()
        self.root.after(150, self._process_events)

    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text=APP_TITLE, font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="Reads Contacts and Messages sheets from an Excel workbook. "
                 "Only contacts marked OptIn = Yes are eligible.",
            wraplength=700
        ).pack(anchor="w", pady=(5, 14))

        file_row = ttk.Frame(outer)
        file_row.pack(fill="x", pady=4)
        ttk.Entry(file_row, textvariable=self.file_path).pack(side="left", fill="x", expand=True)
        ttk.Button(file_row, text="Browse…", command=self._browse).pack(side="left", padx=(8, 0))
        ttk.Button(file_row, text="Create sample Excel", command=self._create_template).pack(side="left", padx=(8, 0))

        ttk.Separator(outer).pack(fill="x", pady=14)

        stats = ttk.Frame(outer)
        stats.pack(fill="x")
        ttk.Label(stats, text="Progress", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Label(stats, textvariable=self.progress_text).pack(side="right")

        self.progress = ttk.Progressbar(outer, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(6, 12))

        controls = ttk.Frame(outer)
        controls.pack(fill="x")
        self.start_btn = ttk.Button(controls, text="Start sending", command=self._start)
        self.start_btn.pack(side="left")
        self.stop_btn = ttk.Button(controls, text="Stop after current attempt", command=self._stop, state="disabled")
        self.stop_btn.pack(side="left", padx=8)
        ttk.Button(controls, text="Open log folder", command=self._open_log_folder).pack(side="right")

        ttk.Label(outer, textvariable=self.status_text, wraplength=700).pack(anchor="w", pady=(12, 8))

        ttk.Label(outer, text="Activity", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.activity = tk.Text(outer, height=15, wrap="word", state="disabled", font=("Consolas", 9))
        self.activity.pack(fill="both", expand=True, pady=(5, 0))

        note = (
            "Important: Use only for recipients who explicitly opted in. "
            "Follow WhatsApp policies and honor opt-outs. PyWhatKit automates a browser; "
            "a logged 'SUBMITTED' result is not proof of delivery or reading. "
            "Keep WhatsApp Web logged in. Do not use this to send unsolicited bulk messages."
        )
        ttk.Label(outer, text=note, wraplength=700, foreground="#8a4b08").pack(anchor="w", pady=(10, 0))

    def _browse(self):
        selected = filedialog.askopenfilename(
            title="Select WhatsApp Excel workbook",
            filetypes=[("Excel workbook", "*.xlsx")]
        )
        if selected:
            self.file_path.set(selected)

    def _create_template(self):
        path = filedialog.asksaveasfilename(
            title="Save sample workbook",
            defaultextension=".xlsx",
            initialfile="WhatsApp_Bulk_Messages.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")]
        )
        if not path:
            return
        from openpyxl import Workbook
        wb = Workbook()
        contacts = wb.active
        contacts.title = "Contacts"
        contacts.append(["Phone", "Name", "OptIn"])
        contacts.append(["+919876543210", "Example Recipient", "No"])
        messages = wb.create_sheet("Messages")
        messages.append(["Message"])
        messages.append(["Hello {name}, this is an example message."])
        wb.save(path)
        self.file_path.set(path)
        self._append_activity(f"Created workbook: {path}")
        messagebox.showinfo(
            "Workbook created",
            "Template created. Replace the example contact and set OptIn to Yes only "
            "for people who agreed to receive WhatsApp messages."
        )

    def _append_activity(self, text):
        self.activity.configure(state="normal")
        self.activity.insert("end", f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {text}\n")
        self.activity.see("end")
        self.activity.configure(state="disabled")

    def _start(self):
        path = self.file_path.get().strip()
        if not path or not Path(path).is_file():
            messagebox.showerror("Excel file required", "Choose a valid .xlsx workbook first.")
            return
        try:
            contacts = pd.read_excel(path, sheet_name="Contacts")
            messages = pd.read_excel(path, sheet_name="Messages")
            required = {"Phone", "Name", "OptIn"}
            if not required.issubset(set(contacts.columns)):
                raise ValueError("Contacts sheet must contain columns: Phone, Name, OptIn.")
            if "Message" not in messages.columns or messages.empty:
                raise ValueError("Messages sheet must contain a non-empty Message column.")
            eligible = contacts[
                contacts["Phone"].notna()
                & contacts["OptIn"].astype(str).str.strip().str.lower().isin(["yes", "true", "1"])
            ]
            if eligible.empty:
                raise ValueError("No eligible contacts. Add recipients who opted in and set OptIn to Yes.")
            message = str(messages.iloc[0]["Message"]).strip()
            if not message or message.lower() == "nan":
                raise ValueError("The Message cell is empty.")
        except Exception as exc:
            messagebox.showerror("Workbook validation failed", str(exc))
            return

        if not messagebox.askyesno(
            "Confirm campaign",
            f"{len(eligible)} opted-in contacts are eligible.\n\n"
            "PyWhatKit will use WhatsApp Web in your browser. "
            "Continue only if these recipients agreed to receive this message.\n\nStart?"
        ):
            return

        self.stop_event.clear()
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.progress.configure(value=0, maximum=len(eligible))
        self.progress_text.set(f"0 / {len(eligible)}")
        self.status_text.set("Starting… Keep the browser session available.")
        self.worker = threading.Thread(
            target=self._send_worker,
            args=(path,),
            daemon=True
        )
        self.worker.start()

    def _send_worker(self, path):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = Path(path).resolve().parent / "logs"
        log_dir.mkdir(exist_ok=True)
        self.log_path = log_dir / f"whatsapp_log_{timestamp}.csv"

        try:
            contacts = pd.read_excel(path, sheet_name="Contacts")
            messages = pd.read_excel(path, sheet_name="Messages")
            message_template = str(messages.iloc[0]["Message"]).strip()

            eligible = contacts[
                contacts["Phone"].notna()
                & contacts["OptIn"].astype(str).str.strip().str.lower().isin(["yes", "true", "1"])
            ]

            total = len(eligible)
            processed = 0
            with open(self.log_path, "w", newline="", encoding="utf-8-sig") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Timestamp", "Phone", "Name", "Status", "Details"])

                for _, row in eligible.iterrows():
                    if self.stop_event.is_set():
                        break

                    phone = str(row["Phone"]).strip()
                    # Avoid Excel converting phone numbers to floats/scientific notation.
                    if phone.endswith(".0"):
                        phone = phone[:-2]
                    name = "" if pd.isna(row["Name"]) else str(row["Name"]).strip()
                    personalized_message = message_template.replace("{name}", name or "there")
                    stamp = datetime.now().astimezone().isoformat(timespec="seconds")

                    try:
                        if not phone.startswith("+") or not phone[1:].isdigit():
                            raise ValueError("Phone must include country code, e.g. +919876543210")

                        pywhatkit.sendwhatmsg_instantly(
                            phone,
                            personalized_message,
                            wait_time=BROWSER_WAIT_SECONDS,
                            tab_close=True
                        )
                        status = "SUBMITTED"
                        details = "PyWhatKit call returned; delivery is not confirmed."
                    except Exception as exc:
                        status = "FAILED"
                        details = str(exc).replace("\n", " ")[:1000]

                    writer.writerow([stamp, phone, name, status, details])
                    csvfile.flush()
                    processed += 1
                    self.events.put(("progress", processed, total, phone, status))

            final = "Stopped by user." if self.stop_event.is_set() else "Processing completed."
            self.events.put(("finished", final, str(self.log_path), processed, total))
        except Exception as exc:
            self.events.put(("fatal", str(exc)))

    def _stop(self):
        self.stop_event.set()
        self.status_text.set("Stop requested. The current browser attempt will finish first.")
        self.stop_btn.configure(state="disabled")

    def _process_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    _, done, total, phone, status = event
                    self.progress.configure(value=done)
                    self.progress_text.set(f"{done} / {total}")
                    self.status_text.set(f"{status}: {phone}")
                    self._append_activity(f"{status}: {phone}")
                elif event[0] == "finished":
                    _, final, log_path, processed, total = event
                    self.start_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    self.status_text.set(f"{final} {processed}/{total} processed. Log: {log_path}")
                    self._append_activity(final)
                    messagebox.showinfo("Finished", f"{final}\nProcessed: {processed}/{total}\n\nLog:\n{log_path}")
                elif event[0] == "fatal":
                    self.start_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    self.status_text.set("Application error.")
                    self._append_activity(f"ERROR: {event[1]}")
                    messagebox.showerror("Application error", event[1])
        except queue.Empty:
            pass
        self.root.after(150, self._process_events)

    def _open_log_folder(self):
        import os
        import sys
        folder = Path(self.file_path.get()).resolve().parent / "logs" if self.file_path.get() else Path.cwd() / "logs"
        folder.mkdir(exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)
            else:
                messagebox.showinfo("Log folder", str(folder))
        except Exception as exc:
            messagebox.showerror("Cannot open folder", str(exc))

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppBulkSender(root)
    root.mainloop()
