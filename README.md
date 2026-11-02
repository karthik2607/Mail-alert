# MailAlert – never miss a mail in new Outlook

A small Python script for Windows that works with the **new Outlook** app.

New Outlook shows a Windows notification for every new mail, and Windows keeps a copy
of each notification on your PC. MailAlert reads those copies and turns them into
**one popup at the bottom-right that keeps coming back every 5 minutes** until you tick
**Read**.

- All mail IDs in new Outlook, in **one popup**, always on top.
- **Later** → comes back in 5 minutes. **Read** → gone for good. **Mark all read** → clears all.
- **Red alert** (also every 5 minutes) if new Outlook is closed or the notifications
  can't be read – so you know when mail is *not* being watched.
- Meeting / calendar reminders are skipped.
- Notifications that arrived while the script was stopped pop up when you start it.

No .exe, no classic Outlook, no Client ID, no IT approval, no passwords.

**Limits (because new Outlook doesn't let other programs read mail directly):**
ticking **Read** clears it in MailAlert only (not in Outlook); **Open** brings up Outlook
but not the exact mail; MailAlert can't tell if an account is signed out.

---

## One-time setup (≈10 minutes)

### 1. Turn on mail notifications in new Outlook
**Settings (gear icon) → General → Notifications**
- Turn **Mail** notifications **on**.
- Choose to be notified about **all** email, not only Focused Inbox.
- If it lists accounts, make sure **every** mail ID is switched on.
- Keep **"Notify when Outlook is closed / run in background"** on if your version has it.

### 2. Make sure Windows keeps Outlook's notifications
**Windows Settings → System → Notifications → Outlook (new)**
- **Notifications:** On
- **Show notifications in notification center:** On (MailAlert needs this)

Windows *Do Not Disturb* is fine – notifications still go to the notification center.

### 3. Install the 2 packages
In Command Prompt / Git Bash inside the MailAlert folder:
```
pip install -r requirements.txt
```

### 4. Check that MailAlert can see your mail
Send yourself a test mail (or wait for one), then run:
```
python mailalert.py --diagnose
```
You should see **Outlook** in the list marked `<-- treated as MAIL`, and your recent mails
underneath. If not, re-check steps 1–2, and send the output to Claude.

## Every day

| | |
|---|---|
| **Start** (when you switch on the PC) | Double-click `start_mailalert.bat` – or run `python mailalert.py` |
| **Stop** (when you leave office) | Blue envelope tray icon (bottom-right, maybe under **^**) → **Quit** |

Tip: don't click **Clear all** in the Windows notification center while MailAlert is stopped –
those cleared notifications can't be picked up later. While it's running, clearing is fine.

## Settings (`%APPDATA%\MailAlert\config.json`, created on first run)

| Setting | Default | Meaning |
|---|---|---|
| `poll_seconds` | 5 | How often notifications are read |
| `remind_minutes` | 5 | Popup repeats this often until you tick Read |
| `ignore_reminders` | true | Skip meeting / calendar reminders |
| `ignore_subject_contains` | [] | Skip mails containing these words, e.g. `["Newsletter"]` |
| `check_outlook_running` | true | Red alert if new Outlook is closed |
| `app_ids` | Outlook | Which apps count as mail. Add more (e.g. a Gmail app's ID from `--diagnose`) to watch them too |

Restart the script after changing settings. Log file: `%APPDATA%\MailAlert\mailalert.log`.
