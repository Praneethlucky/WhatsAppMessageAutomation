
import csv, threading, queue, os
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pywhatkit
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

WAIT_TIME = 30  # Allow WhatsApp Web extra time to load; not an inter-recipient delay.

class App:
    def __init__(self, root):
        self.root = root
        root.title("WhatsApp Bulk Sender")
        root.geometry("780x650")
        root.minsize(680, 560)
        self.excel = tk.StringVar()
        self.image = tk.StringVar()
        self.status = tk.StringVar(value="Select a contacts workbook.")
        self.events = queue.Queue()
        self.stop_event = threading.Event()
        self.build()

    def build(self):
        f = ttk.Frame(self.root, padding=16); f.pack(fill="both", expand=True)
        ttk.Label(f, text="WhatsApp Bulk Sender", font=("Segoe UI", 19, "bold")).pack(anchor="w")
        ttk.Label(f, text="Contacts come from Excel. Compose one shared message and optionally choose one shared image here.", wraplength=720).pack(anchor="w", pady=(4,12))

        ttk.Label(f, text="Contacts Excel (.xlsx)").pack(anchor="w")
        r = ttk.Frame(f); r.pack(fill="x", pady=4)
        ttk.Entry(r, textvariable=self.excel).pack(side="left", fill="x", expand=True)
        ttk.Button(r, text="Browse", command=self.pick_excel).pack(side="left", padx=5)
        ttk.Button(r, text="Create sample", command=self.sample).pack(side="left")

        ttk.Label(f, text="Message").pack(anchor="w", pady=(10,3))
        self.msg = tk.Text(f, height=5, wrap="word")
        self.msg.pack(fill="x")
        self.msg.insert("1.0", "Hello {name}, this is a message for you.")
        ttk.Label(f, text="Use {name} to personalize the message. Leave blank if sending an image only.", foreground="#555").pack(anchor="w", pady=(3,10))

        ttk.Label(f, text="Optional image (same image for all contacts)").pack(anchor="w")
        r = ttk.Frame(f); r.pack(fill="x", pady=4)
        ttk.Entry(r, textvariable=self.image).pack(side="left", fill="x", expand=True)
        ttk.Button(r, text="Choose image", command=self.pick_image).pack(side="left", padx=5)
        ttk.Button(r, text="Clear", command=lambda:self.image.set("")).pack(side="left")

        ttk.Separator(f).pack(fill="x", pady=12)
        r = ttk.Frame(f); r.pack(fill="x")
        self.progress = ttk.Progressbar(r, maximum=100); self.progress.pack(side="left", fill="x", expand=True)
        self.progress_label = ttk.Label(r, text="0 / 0"); self.progress_label.pack(side="left", padx=10)
        r = ttk.Frame(f); r.pack(fill="x", pady=8)
        self.start_btn = ttk.Button(r, text="Start sending", command=self.start); self.start_btn.pack(side="left")
        self.stop_btn = ttk.Button(r, text="Stop after current attempt", command=self.stop, state="disabled"); self.stop_btn.pack(side="left", padx=8)
        ttk.Button(r, text="Open logs", command=self.open_logs).pack(side="right")
        ttk.Label(f, textvariable=self.status, wraplength=720).pack(anchor="w", pady=5)
        ttk.Label(f, text="Activity").pack(anchor="w")
        self.log = tk.Text(f, height=12, state="disabled", wrap="word", font=("Consolas",9)); self.log.pack(fill="both", expand=True)
        ttk.Label(f, text="Only send to people who explicitly opted in. ATTEMPTED means the browser automation call returned; it does not confirm delivery or reading. The WhatsApp tab is closed after each attempt to prevent tab buildup.", wraplength=720, foreground="#8a4b08").pack(anchor="w", pady=(8,0))
        self.root.after(150, self.poll)

    def add_log(self, s):
        self.log.configure(state="normal")
        self.log.insert("end", f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {s}\\n")
        self.log.see("end"); self.log.configure(state="disabled")

    def pick_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel workbook","*.xlsx")])
        if p: self.excel.set(p)

    def pick_image(self):
        p = filedialog.askopenfilename(filetypes=[("Images","*.png *.jpg *.jpeg *.bmp *.gif"),("All files","*.*")])
        if p: self.image.set(p)

    def sample(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="WhatsApp_Contacts.xlsx", filetypes=[("Excel workbook","*.xlsx")])
        if not p: return
        wb = Workbook(); ws = wb.active; ws.title = "Contacts"
        ws.append(["Phone","Name","OptIn"]); ws.append(["+919876543210","Example Recipient","No"])
        ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
        for c in ws[1]:
            c.font = Font(bold=True,color="FFFFFF"); c.fill = PatternFill("solid",fgColor="217346")
        for col in ws.columns:
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(16,max(len(str(c.value or "")) for c in col)+3)
        wb.save(p); self.excel.set(p)
        messagebox.showinfo("Sample created","Add opted-in recipients to Contacts. Set OptIn to Yes only when they agreed to receive WhatsApp messages.")

    @staticmethod
    def phone(v):
        s = str(v).strip()
        if s.endswith(".0"): s = s[:-2]
        s = s.replace(" ","").replace("-","").replace("(","").replace(")","")
        if not s.startswith("+"): s = "+" + s
        if not s[1:].isdigit(): raise ValueError("Use international format, e.g. +919876543210")
        return s

    def start(self):
        x = self.excel.get().strip(); msg = self.msg.get("1.0","end").strip(); img = self.image.get().strip()
        if not x or not Path(x).is_file():
            messagebox.showerror("Excel required","Select a valid contacts workbook."); return
        if not msg and not img:
            messagebox.showerror("Content required","Enter a message, choose an image, or both."); return
        if img and not Path(img).is_file():
            messagebox.showerror("Image missing","Choose an existing image."); return
        try:
            df = pd.read_excel(x, sheet_name="Contacts", dtype={"Phone":str})
            if not {"Phone","Name","OptIn"}.issubset(df.columns):
                raise ValueError("Contacts sheet needs Phone, Name and OptIn columns.")
            df = df[df.Phone.notna() & df.OptIn.astype(str).str.strip().str.lower().isin(["yes","true","1"])]
            if df.empty: raise ValueError("No opted-in contacts found (OptIn must be Yes).")
        except Exception as e:
            messagebox.showerror("Workbook error",str(e)); return
        kind = "image + caption" if img and msg else ("image only" if img else "text")
        if not messagebox.askyesno("Confirm",f"{len(df)} opted-in contacts.\\nContent: {kind}\\n\\nSame message/image will be used for everyone. Continue only if recipients opted in. Start?"):
            return
        self.stop_event.clear(); self.start_btn.configure(state="disabled"); self.stop_btn.configure(state="normal")
        self.progress.configure(maximum=len(df),value=0); self.progress_label.configure(text=f"0 / {len(df)}")
        self.status.set("Sending started. Keep WhatsApp Web logged in.")
        threading.Thread(target=self.worker,args=(x,msg,img),daemon=True).start()

    def worker(self,x,msg,img):
        try:
            df = pd.read_excel(x,sheet_name="Contacts",dtype={"Phone":str})
            df = df[df.Phone.notna() & df.OptIn.astype(str).str.strip().str.lower().isin(["yes","true","1"])]
            folder = Path(x).resolve().parent / "logs"; folder.mkdir(exist_ok=True)
            logpath = folder / f"whatsapp_log_{datetime.now():%Y%m%d_%H%M%S}.csv"
            done=0; total=len(df)
            with open(logpath,"w",newline="",encoding="utf-8-sig") as fp:
                w=csv.writer(fp); w.writerow(["Timestamp","Phone","Name","Type","Status","Details"])
                for _,row in df.iterrows():
                    if self.stop_event.is_set(): break
                    name = "" if pd.isna(row.Name) else str(row.Name).strip()
                    try:
                        phone=self.phone(row.Phone); body=msg.replace("{name}",name or "there")
                        if img:
                            args={"receiver":phone,"img_path":img,"wait_time":WAIT_TIME,"tab_close":True}
                            if body: args["caption"]=body
                            pywhatkit.sendwhats_image(**args); typ="Image + caption" if body else "Image"
                        else:
                            pywhatkit.sendwhatmsg_instantly(phone,body,wait_time=WAIT_TIME,tab_close=True); typ="Text"
                        status="ATTEMPTED"; details="PyWhatKit returned. Verify the message appears as a sent bubble; composer text alone is not success."
                    except Exception as e:
                        phone=str(row.Phone); typ="Image" if img else "Text"; status="FAILED"; details=str(e).replace("\\n"," ")[:1000]
                    w.writerow([datetime.now().astimezone().isoformat(timespec="seconds"),phone,name,typ,status,details]); fp.flush()
                    done+=1; self.events.put(("progress",done,total,phone,status))
            self.events.put(("done",str(logpath),done,total,self.stop_event.is_set()))
        except Exception as e: self.events.put(("error",str(e)))

    def stop(self):
        self.stop_event.set(); self.stop_btn.configure(state="disabled")
        self.status.set("Stop requested; current browser attempt will finish first.")

    def poll(self):
        try:
            while True:
                e=self.events.get_nowait()
                if e[0]=="progress":
                    _,d,t,p,s=e; self.progress.configure(value=d); self.progress_label.configure(text=f"{d} / {t}")
                    self.status.set(f"{s}: {p}"); self.add_log(f"{s}: {p}")
                elif e[0]=="done":
                    _,logpath,d,t,stopped=e; self.start_btn.configure(state="normal"); self.stop_btn.configure(state="disabled")
                    text=("Stopped" if stopped else "Completed")+f": {d}/{t}. Log: {logpath}"
                    self.status.set(text); self.add_log(text); messagebox.showinfo("Finished",text)
                elif e[0]=="error":
                    self.start_btn.configure(state="normal"); self.stop_btn.configure(state="disabled")
                    self.status.set("Application error"); messagebox.showerror("Error",e[1])
        except queue.Empty: pass
        self.root.after(150,self.poll)

    def open_logs(self):
        folder=Path(self.excel.get()).resolve().parent/"logs" if self.excel.get() else Path.cwd()/"logs"
        folder.mkdir(exist_ok=True)
        try: os.startfile(folder)
        except Exception as e: messagebox.showerror("Cannot open logs",str(e))

if __name__=="__main__":
    root=tk.Tk(); App(root); root.mainloop()
