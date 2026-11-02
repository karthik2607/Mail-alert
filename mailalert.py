"""
MailAlert - one repeating popup for every mail ID in the NEW Outlook app.

How it works
- New Outlook shows a Windows notification for each new mail. Windows keeps a copy
  of every notification in a local database on your PC. MailAlert reads those copies
  - no Client ID, no admin approval, no classic Outlook, no passwords.
- New mail -> one popup at the bottom-right of the screen, always on top.
- Click "Later" and it comes back every 5 minutes until you tick "Read".
- Red alert (also repeating) if new Outlook is closed, or if Windows notifications
  can't be read - so you know when mail is NOT being watched.
- Mail notifications that arrived while the script was stopped pop up when you start it
  (as long as you didn't clear them from the Windows notification center).

Run:       python mailalert.py           (or double-click start_mailalert.bat)
Check:     python mailalert.py --diagnose   (shows what MailAlert can see)
Stop:      tray icon (blue envelope) > Quit
"""

import json
import logging
import os
import queue
import sqlite3
import subprocess
import sys
import threading
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone

# ---------------------------------------------------------------- paths/config
APP_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "MailAlert")
os.makedirs(APP_DIR, exist_ok=True)
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
DB_PATH = os.path.join(APP_DIR, "mailalert_notify.db")
LOG_PATH = os.path.join(APP_DIR, "mailalert.log")
WIN_NOTIF_DB = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "Microsoft", "Windows", "Notifications", "wpndatabase.db",
)

NEW_OUTLOOK_AUMID = "Microsoft.OutlookForWindows_8wekyb3d8bbwe!Microsoft.OutlookforWindows"

DEFAULT_CONFIG = {
    "poll_seconds": 5,                 # how often Windows notifications are read
    "remind_minutes": 5,               # popup repeats this often until ticked
    # notifications from apps whose ID contains any of these texts are treated as mail
    "app_ids": ["OutlookForWindows", "OUTLOOK.EXE"],
    "ignore_reminders": True,          # skip calendar/meeting reminder notifications
    "ignore_subject_contains": [],     # e.g. ["Reminder:"] to skip by text
    "check_outlook_running": True,     # red alert if new Outlook is closed
    "outlook_process": "olk.exe",
}

logging.basicConfig(filename=LOG_PATH, level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("mailalert")

ui_queue: "queue.Queue[tuple]" = queue.Queue()
stop_event = threading.Event()
poll_now = threading.Event()

# problems shown as red lines; written by the worker thread, read by the UI
health = {"outlook_closed": False, "notif_error": ""}


def load_config():
    if not os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, indent=2)
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            return {**DEFAULT_CONFIG, **json.load(f)}
    except (OSError, ValueError):
        log.exception("Bad config, using defaults")
        return dict(DEFAULT_CONFIG)


# ---------------------------------------------------------------- time helpers
EPOCH_1601 = datetime(1601, 1, 1, tzinfo=timezone.utc)


def filetime_to_local(ft):
    return (EPOCH_1601 + timedelta(microseconds=int(ft) // 10)).astimezone().replace(tzinfo=None)


def now_filetime():
    return int((datetime.now(timezone.utc) - EPOCH_1601).total_seconds() * 10_000_000)


# ---------------------------------------------------------------- toast parsing
def parse_toast(payload):
    """Returns dict(scenario, texts, attribution) or None if unreadable."""
    if payload is None:
        return None
    if isinstance(payload, (bytes, bytearray)):
        payload = bytes(payload).decode("utf-8", "ignore")
    payload = payload.lstrip("﻿").strip("\x00").strip()
    try:
        root = ET.fromstring(payload)
    except ET.ParseError:
        return None
    texts, attribution = [], []
    for el in root.iter():
        if el.tag.split("}")[-1] != "text":
            continue
        val = "".join(el.itertext()).strip()
        if not val:
            continue
        (attribution if el.get("placement") == "attribution" else texts).append(val)
    return {
        "scenario": (root.get("scenario") or "").lower(),
        "texts": texts,
        "attribution": " ".join(attribution),
    }


# ---------------------------------------------------------------- our storage
class Store:
    def __init__(self, path):
        self.lock = threading.Lock()
        self.db = sqlite3.connect(path, check_same_thread=False)
        self.db.executescript(
            """
            CREATE TABLE IF NOT EXISTS messages(
                key TEXT PRIMARY KEY, sender TEXT, subject TEXT, preview TEXT,
                account TEXT, received TEXT, acked INTEGER DEFAULT 0);
            CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT);
            """
        )
        self.db.commit()

    def _q(self, sql, args=(), commit=False):
        with self.lock:
            cur = self.db.execute(sql, args)
            if commit:
                self.db.commit()
            return cur.fetchall()

    def get_meta(self, k, default=None):
        r = self._q("SELECT v FROM meta WHERE k=?", (k,))
        return r[0][0] if r else default

    def set_meta(self, k, v):
        self._q("INSERT OR REPLACE INTO meta(k, v) VALUES(?,?)", (k, str(v)), commit=True)

    def insert(self, key, sender, subject, preview, account, received):
        """Returns True if this mail was new."""
        before = self._q("SELECT 1 FROM messages WHERE key=?", (key,))
        if before:
            return False
        self._q("INSERT INTO messages(key, sender, subject, preview, account, received) "
                "VALUES(?,?,?,?,?,?)", (key, sender, subject, preview, account, received),
                commit=True)
        return True

    def pending(self):
        rows = self._q("SELECT key, sender, subject, preview, account, received "
                       "FROM messages WHERE acked=0 ORDER BY received DESC")
        keys = ("key", "sender", "subject", "preview", "account", "received")
        return [dict(zip(keys, r)) for r in rows]

    def ack(self, key):
        self._q("UPDATE messages SET acked=1 WHERE key=?", (key,), commit=True)

    def ack_all(self):
        self._q("UPDATE messages SET acked=1 WHERE acked=0", commit=True)


# ---------------------------------------------------------------- reading Windows notifications
def read_windows_toasts(since_ft, db_path=WIN_NOTIF_DB):
    """Rows (id, arrival_filetime, payload, app_id) of toasts newer than since_ft."""
    uri = "file:" + db_path.replace("\\", "/") + "?mode=ro"
    con = sqlite3.connect(uri, uri=True, timeout=5)
    try:
        return con.execute(
            "SELECT n.Id, n.ArrivalTime, n.Payload, h.PrimaryId "
            "FROM Notification n JOIN NotificationHandler h ON h.RecordId = n.HandlerId "
            "WHERE n.Type = 'toast' AND n.ArrivalTime > ? ORDER BY n.ArrivalTime",
            (since_ft,),
        ).fetchall()
    finally:
        con.close()


def outlook_is_running(process_name):
    if os.name != "nt":
        return True
    try:
        out = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {process_name}", "/NH"],
            capture_output=True, text=True, timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        ).stdout
        return process_name.lower() in out.lower()
    except Exception:
        log.exception("tasklist failed")
        return True  # don't raise a false alarm if we can't tell


def open_outlook():
    try:
        subprocess.Popen(["explorer.exe", f"shell:AppsFolder\\{NEW_OUTLOOK_AUMID}"])
    except Exception:
        log.exception("Could not open Outlook")


class NotifWorker:
    OVERLAP = 60 * 10_000_000  # re-read the last 60 s each time; duplicates are skipped

    def __init__(self, store, cfg, db_path=WIN_NOTIF_DB):
        self.store, self.cfg, self.db_path = store, cfg, db_path
        self.fail_count = 0
        if self.store.get_meta("last_ft") is None:
            # first ever run: start from now, so old notifications don't flood you
            self.store.set_meta("last_ft", now_filetime())

    def is_mail_app(self, app_id):
        app_id = (app_id or "").lower()
        return any(p.lower() in app_id for p in self.cfg["app_ids"])

    def check_once(self):
        """One pass. Returns number of new mails."""
        last_ft = int(self.store.get_meta("last_ft"))
        rows = read_windows_toasts(last_ft - self.OVERLAP, self.db_path)
        new, max_ft = 0, last_ft
        for nid, ft, payload, app_id in rows:
            max_ft = max(max_ft, int(ft))
            if not self.is_mail_app(app_id):
                continue
            t = parse_toast(payload)
            if not t or not t["texts"]:
                continue
            if self.cfg["ignore_reminders"] and t["scenario"] in ("reminder", "alarm", "incomingcall"):
                continue
            joined = " ".join(t["texts"]).lower()
            if any(s.lower() in joined for s in self.cfg["ignore_subject_contains"]):
                continue
            sender = t["texts"][0]
            subject = t["texts"][1] if len(t["texts"]) > 1 else ""
            preview = " ".join(t["texts"][2:])
            received = filetime_to_local(ft).isoformat()
            if self.store.insert(f"{nid}:{ft}", sender, subject, preview, t["attribution"], received):
                new += 1
        self.store.set_meta("last_ft", max_ft)
        return new

    def run(self):
        while not stop_event.is_set():
            try:
                new = self.check_once()
                self.fail_count = 0
                self.set_error("")
                if new:
                    log.info("%d new mail notification(s)", new)
                    ui_queue.put(("alert",))
            except Exception as e:
                self.fail_count += 1
                log.exception("Reading Windows notifications failed")
                if self.fail_count >= 3:  # ignore brief 'database is locked' blips
                    self.set_error(f"Can't read Windows notifications ({e}) - "
                                   "your mail is NOT being monitored")

            if self.cfg["check_outlook_running"]:
                closed = not outlook_is_running(self.cfg["outlook_process"])
                if closed != health["outlook_closed"]:
                    health["outlook_closed"] = closed
                    log.warning("New Outlook %s", "closed" if closed else "running again")
                    ui_queue.put(("alert",) if closed else ("refresh",))

            poll_now.wait(self.cfg["poll_seconds"])
            poll_now.clear()

    def set_error(self, msg):
        if msg != health["notif_error"]:
            health["notif_error"] = msg
            ui_queue.put(("alert",) if msg else ("refresh",))


# ---------------------------------------------------------------- popup UI
BG, FG, MUTED, ACCENT, WARN = "#1f2430", "#f2f4f8", "#a3a9b8", "#4c8dff", "#ff5c5c"
ROW, WARN_ROW = "#272d3b", "#3a1f24"
WIDTH = 460
MAX_ROWS = 8


def build_ui_classes(tk):
    class AlertUI:
        def __init__(self, root, store, cfg):
            self.root, self.store, self.cfg = root, store, cfg
            self.win = None
            self.last_dismiss = 0.0
            root.after(1000, self.tick)

        def problems(self):
            p = []
            if health["outlook_closed"]:
                p.append(("New Outlook is closed - you won't get mail notifications",
                          "Open Outlook", open_outlook))
            if health["notif_error"]:
                p.append((health["notif_error"], "Retry", poll_now.set))
            return p

        def tick(self):
            try:
                while True:
                    kind, *args = ui_queue.get_nowait()
                    if kind == "alert":
                        self.show()
                    elif kind == "refresh":
                        if self.win is not None:
                            self.show()
                    elif kind == "show":
                        self.show(force=True)
                    elif kind == "info":
                        self.toast(args[0])
                    elif kind == "quit":
                        self.root.quit()
                        return
            except queue.Empty:
                pass
            attention = self.store.pending() or self.problems()
            due = time.time() - self.last_dismiss >= self.cfg["remind_minutes"] * 60
            if attention and self.win is None and due:
                self.show()
            self.root.after(1000, self.tick)

        def close(self, dismissed=True):
            if self.win is not None:
                self.win.destroy()
                self.win = None
            if dismissed:
                self.last_dismiss = time.time()

        def show(self, force=False):
            msgs, probs = self.store.pending(), self.problems()
            if not msgs and not probs:
                self.close(dismissed=False)
                if force:
                    self.toast("All caught up - no pending mail.")
                return
            self.close(dismissed=False)

            w = tk.Toplevel(self.root, bg=BG)
            w.overrideredirect(True)
            w.attributes("-topmost", True)
            self.win = w

            head = tk.Frame(w, bg=BG)
            head.pack(fill="x", padx=12, pady=(10, 4))
            title = (f"{len(msgs)} new mail" + ("" if len(msgs) == 1 else "s")) if msgs else "MailAlert"
            tk.Label(head, text="✉  " + title, bg=BG, fg=FG,
                     font=("Segoe UI", 11, "bold")).pack(side="left")
            tk.Button(head, text=f"Later ({self.cfg['remind_minutes']} min)", command=self.close,
                      bg=BG, fg=MUTED, bd=0, activebackground=BG, cursor="hand2").pack(side="right")

            for text, btn, cmd in probs:
                row = tk.Frame(w, bg=WARN_ROW)
                row.pack(fill="x", padx=12, pady=3)
                tk.Label(row, text="⚠  " + text, bg=WARN_ROW, fg=WARN, font=("Segoe UI", 9, "bold"),
                         wraplength=320, justify="left").pack(side="left", padx=8, pady=6)
                tk.Button(row, text=btn, bg=WARN, fg="white", bd=0, cursor="hand2",
                          command=cmd).pack(side="right", padx=8)

            for m in msgs[:MAX_ROWS]:
                row = tk.Frame(w, bg=ROW)
                row.pack(fill="x", padx=12, pady=3)
                tk.Checkbutton(row, text="Read", bg=ROW, fg=FG, selectcolor=ROW,
                               activebackground=ROW, activeforeground=FG,
                               command=lambda k=m["key"]: self.ack(k)).pack(side="left", padx=4)
                txt = tk.Frame(row, bg=ROW)
                txt.pack(side="left", fill="x", expand=True, pady=4)
                when = m["received"][11:16]
                day = "" if m["received"][:10] == datetime.now().date().isoformat() else m["received"][5:10] + " "
                meta = "  •  ".join(x for x in (m["account"], m["sender"], day + when) if x)
                tk.Label(txt, text=meta, bg=ROW, fg=MUTED, font=("Segoe UI", 8), anchor="w",
                         wraplength=300, justify="left").pack(fill="x")
                tk.Label(txt, text=m["subject"] or m["preview"] or "(no subject)", bg=ROW, fg=FG,
                         font=("Segoe UI", 10, "bold"), anchor="w", wraplength=300,
                         justify="left").pack(fill="x")
                tk.Button(row, text="Open", bg=ACCENT, fg="white", bd=0, cursor="hand2",
                          command=open_outlook).pack(side="right", padx=6)

            if len(msgs) > MAX_ROWS:
                tk.Label(w, text=f"+ {len(msgs) - MAX_ROWS} more (tick some to see the rest)",
                         bg=BG, fg=MUTED).pack(pady=2)

            foot = tk.Frame(w, bg=BG)
            foot.pack(fill="x", padx=12, pady=(4, 10))
            if msgs:
                tk.Button(foot, text="Mark all read", command=self.ack_all, bg=BG, fg=ACCENT,
                          bd=0, cursor="hand2").pack(side="right")

            w.update_idletasks()
            h = w.winfo_reqheight()
            sw, sh = w.winfo_screenwidth(), w.winfo_screenheight()
            w.geometry(f"{WIDTH}x{h}+{sw - WIDTH - 16}+{sh - h - 60}")
            w.bell()

        def after_ack(self):
            if self.store.pending() or self.problems():
                self.show()
            else:
                self.close()

        def ack(self, key):
            self.store.ack(key)
            self.after_ack()

        def ack_all(self):
            self.store.ack_all()
            self.after_ack()

        def toast(self, text, ms=4000):
            t = tk.Toplevel(self.root, bg=BG)
            t.overrideredirect(True)
            t.attributes("-topmost", True)
            tk.Label(t, text=text, bg=BG, fg=FG, font=("Segoe UI", 10),
                     wraplength=WIDTH - 30).pack(padx=14, pady=12)
            t.update_idletasks()
            sw, sh = t.winfo_screenwidth(), t.winfo_screenheight()
            h = t.winfo_reqheight()
            offset = (self.win.winfo_height() + 8) if self.win is not None else 0
            t.geometry(f"{WIDTH}x{h}+{sw - WIDTH - 16}+{sh - h - 60 - offset}")
            t.after(ms, t.destroy)

    return AlertUI


# ---------------------------------------------------------------- tray icon
def start_tray():
    import pystray
    from PIL import Image, ImageDraw

    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((4, 12, 60, 52), radius=8, fill=(76, 141, 255))
    d.line((8, 16, 32, 36, 56, 16), fill="white", width=4)

    def quit_app(icon, _item):
        stop_event.set()
        poll_now.set()
        icon.stop()
        ui_queue.put(("quit",))

    menu = pystray.Menu(
        pystray.MenuItem("Show pending mail", lambda i, it: ui_queue.put(("show",)), default=True),
        pystray.MenuItem("Check now", lambda i, it: poll_now.set()),
        pystray.MenuItem("Open Outlook", lambda i, it: open_outlook()),
        pystray.MenuItem("Quit (stop monitoring)", quit_app),
    )
    icon = pystray.Icon("MailAlert", img, "MailAlert - watching your mail", menu)
    threading.Thread(target=icon.run, daemon=True).start()
    return icon


# ---------------------------------------------------------------- diagnose
def diagnose(cfg):
    print("Windows notification database:", WIN_NOTIF_DB)
    print("  exists:", os.path.exists(WIN_NOTIF_DB))
    try:
        rows = read_windows_toasts(now_filetime() - 7 * 24 * 3600 * 10_000_000)
    except Exception as e:
        print("  ERROR reading it:", e)
        return
    print(f"  toasts in the last 7 days: {len(rows)}")
    apps = {}
    for _, _, _, app in rows:
        apps[app] = apps.get(app, 0) + 1
    print("\nApps that sent notifications (count):")
    for app, n in sorted(apps.items(), key=lambda x: -x[1]):
        mark = "  <-- treated as MAIL" if any(p.lower() in (app or "").lower() for p in cfg["app_ids"]) else ""
        print(f"  {n:4}  {app}{mark}")
    mail = [r for r in rows if any(p.lower() in (r[3] or "").lower() for p in cfg["app_ids"])]
    print(f"\nLast {min(10, len(mail))} mail-app notifications:")
    for nid, ft, payload, app in mail[-10:]:
        t = parse_toast(payload) or {}
        print(f"  {filetime_to_local(ft):%Y-%m-%d %H:%M}  scenario={t.get('scenario') or '-'}"
              f"  attribution={t.get('attribution') or '-'}")
        for i, x in enumerate(t.get("texts", [])):
            print(f"      text{i + 1}: {x[:90]}")
    print("\nOutlook running:", outlook_is_running(cfg["outlook_process"]))
    if not mail:
        print("\nNo Outlook mail notifications found. Check the Outlook and Windows "
              "notification settings in README.md.")


# ---------------------------------------------------------------- main
def main():
    cfg = load_config()
    if "--diagnose" in sys.argv:
        diagnose(cfg)
        return

    import tkinter as tk

    store = Store(DB_PATH)
    worker = NotifWorker(store, cfg)

    root = tk.Tk()
    root.withdraw()
    build_ui_classes(tk)(root, store, cfg)
    icon = start_tray()
    threading.Thread(target=worker.run, daemon=True).start()

    ui_queue.put(("info", "MailAlert is running - watching your Outlook mail notifications."))
    log.info("MailAlert started")
    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass
    stop_event.set()
    try:
        icon.stop()
    except Exception:
        pass
    log.info("MailAlert stopped")


if __name__ == "__main__":
    main()