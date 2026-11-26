# Changelog

## v1.0.0 — 2026-11-26

First public release: two-way CalDAV/CardDAV sync for **Legacy Outlook for Mac**.

### What it does
- Keeps an Outlook calendar and an address book (under *On My Computer*) in two-way sync with a CalDAV/CardDAV server.
- Covers events, all-day events, recurring series with their exceptions, and contacts.
- Runs in the background every 3 minutes while Outlook is open, using launchd.

### Install
```bash
git clone https://github.com/allexkey/outlook-dav-sync.git
cd outlook-dav-sync && ./install.sh
```

### Safety
- The password is stored in the macOS Keychain and is never written to a file.
- The first run backs up every item on the server.
- Every deleted item is backed up first.
- A run that would delete too much stops on its own.
- When an item is edited on both sides, the server version wins and the Outlook version is saved as a backup.
- Outlook edits are merged onto the server copy, so attendees, organizer, alarms and contact photos are kept.

### Tested with
- cPanel CalDAV/CardDAV (`cpdavd`)
- Legacy Outlook 16.115
- macOS 27
- Python 3.14

Reports for other servers (Nextcloud, Radicale, Baïkal, SOGo…) are welcome.

### Known limitations
- New Outlook for Mac is not supported, because it cannot be scripted.
- Reminders and tasks are not synced.
