# outlook-dav-sync

Two-way sync between a **CalDAV / CardDAV server** and **Legacy Outlook for Mac**.

Outlook for Mac cannot connect to CalDAV or CardDAV servers, even though Apple Calendar and Contacts can. This tool runs in the background and keeps a calendar and an address book under Outlook's **On My Computer** in sync with your server. Typical servers are cPanel, Nextcloud, Radicale and Baïkal.

*Română: [README.ro.md](README.ro.md)*

> Not affiliated with or endorsed by Microsoft. "Outlook" is a trademark of Microsoft Corporation.

> **No support — use as is.** This tool was built for my own setup, where it has been tested and is in daily use. It is shared **as is**, without support or warranty. Everyone who installs it uses it at their own risk and is responsible for their own data. Keep your own backups. Issues and pull requests are welcome, but there is no promise of a reply or a fix.

## Requirements

| Requirement | Detail |
|---|---|
| macOS | Tested on macOS 27. |
| Outlook | **Legacy Outlook** for Mac, tested with 16.115. "New Outlook" cannot be scripted; switch with **Outlook → Legacy Outlook**. |
| Python | 3.9 or newer. Only the standard library is used, so there is nothing to `pip install`. |
| Server | Any CalDAV/CardDAV server reachable over HTTPS with username and password. Tested with cPanel (`cpdavd`, port 2080). |

## Install

```bash
git clone <repository-url> outlook-dav-sync
cd outlook-dav-sync
./install.sh
```

The installer:

1. Checks the requirements.
2. Copies the program to `~/Library/Application Support/outlook-dav-sync/bin`. After that you can delete the cloned folder.
3. Asks for your server URL and username.
   - Any URL on the server works, for example `https://mail.example.com:2080/`.
   - The password goes straight into the **macOS Keychain**. It is never stored in a file.
4. Finds your calendars and address books and lets you pick which ones to sync.
5. Runs a **dry run** that shows what would change without changing anything, then asks before the first real sync.
6. Installs a background job (launchd) that syncs every 3 minutes while Outlook is open.

macOS will ask once whether Python may **control Microsoft Outlook**. Click **OK**. If you click "Don't Allow", re-enable it in **System Settings → Privacy & Security → Automation**.

Useful options:

```bash
./install.sh --reconfigure     # change server, account or collections
./install.sh --interval 300    # sync every 5 minutes
```

Re-running `./install.sh` from a newer checkout upgrades the program and keeps your settings and sync state.

## Use

Work in the two Outlook folders the installer created. Their default names are **CalDAV Calendar** and **CardDAV Contacts**. Changes travel both ways within a few minutes.

```bash
APP="$HOME/Library/Application Support/outlook-dav-sync/bin/davsync.py"
python3 "$APP" run              # sync now
python3 "$APP" run --dry-run    # show what would change
python3 "$APP" discover         # list collections on the server
python3 "$APP" status           # number of tracked items
tail -f ~/Library/Logs/outlook-dav-sync.log
```

If your server password changes, replace it in Keychain:

```bash
security add-generic-password -s outlook-dav-sync -a you@example.com -U -w
```

## How it syncs

**Detecting changes**
- Server changes are detected by ETag.
- Outlook changes are detected by comparing **content**, not just the modification date. Outlook touches items on its own, for example it clears reminders of past events, and that does not count as an edit.

**What is sent from Outlook to the server**
- Events: title, location, description, start and end, all-day flag, recurrence and its exceptions.
- Contacts: name, company, job title, phones, e-mails, addresses, web pages, notes and birthday.
- Everything else on the server copy is kept as it is. This includes attendees, organizer, alarms, contact photos and other apps' extensions.

**Conflicts.** If an item is edited on both sides between two runs, the **server version wins**. The Outlook version is saved in `backups/`.

**Deletions**
- Every deleted item is backed up first.
- A run that would delete more than 20% of the items stops and asks you to confirm with `--allow-mass-delete`.

**First run.** All server items are backed up to `backups/initial-…` before anything else happens.

**Reminders** are not synced. Each side keeps its own.

**Only while Outlook is open.** The job does nothing when Outlook is closed and catches up when it reopens.

## Files

| What | Where |
|---|---|
| Program | `~/Library/Application Support/outlook-dav-sync/bin/` |
| Settings | `~/Library/Application Support/outlook-dav-sync/config.json` (example: [config.example.json](config.example.json)) |
| Sync state | `~/Library/Application Support/outlook-dav-sync/state.json` |
| Backups | `~/Library/Application Support/outlook-dav-sync/backups/` |
| Log | `~/Library/Logs/outlook-dav-sync.log` |
| Background job | `~/Library/LaunchAgents/local.outlook-dav-sync.plist` |
| Password | Keychain, service `outlook-dav-sync` |

## Uninstall

```bash
./uninstall.sh
```

The uninstaller stops the background job. It then asks separately whether to delete the Outlook folders, the Keychain password, and the settings and backups. Nothing on the server is touched.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `Outlook bridge failed … not allowed` | Allow Python under **System Settings → Privacy & Security → Automation → Microsoft Outlook**. |
| `No password in Keychain` | Run the `security add-generic-password` command above. |
| `HTTP 401` | Wrong username or password. Run `./install.sh --reconfigure`. |
| `refusing to delete N items` | Check that it is intended, then run once with `--allow-mass-delete`. |
| Duplicate folders in a Microsoft 365 account | The tool only uses folders under **On My Computer**. Delete stray copies in the cloud account. |

## Development

```bash
python3 -m unittest discover tests   # offline tests for the iCalendar/vCard conversions
```

The code has two parts:
- `davsync.py` contains the WebDAV client, the conversions and the sync engine.
- `outlook.js` is the JavaScript for Automation bridge to Outlook's scripting dictionary.

The sync engine tracks each item as server path + ETag ↔ Outlook id + content fingerprint, in `state.json`.

## License

[MIT](LICENSE)
