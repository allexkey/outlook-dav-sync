#!/usr/bin/env python3
"""outlook-dav-sync: two-way sync between a CalDAV/CardDAV server and Legacy Outlook for Mac.

Outlook side: a local ("On My Computer") calendar and address book, driven through
Outlook's AppleScript dictionary (outlook.js). Server side: plain WebDAV/CalDAV/CardDAV.

Commands:
  davsync.py setup               interactive: server, user, password (Keychain), collections
  davsync.py discover            list calendars / address books on the server
  davsync.py run [--dry-run]     one sync pass (what launchd runs every few minutes)
  davsync.py status              show what is tracked

Only the Python standard library is used.
"""
__version__ = "1.0.0"
import argparse
import base64
import datetime
import fcntl
import hashlib
import json
import logging
import logging.handlers
import os
import re
import ssl
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
import uuid
import xml.etree.ElementTree as ET
from urllib.parse import quote, unquote, urljoin, urlparse
from xml.sax.saxutils import escape as xml_escape

HERE = os.path.dirname(os.path.abspath(__file__))
APP_DIR = os.environ.get("DAVSYNC_HOME") or os.path.expanduser("~/Library/Application Support/outlook-dav-sync")
CONFIG_FILE = os.path.join(APP_DIR, "config.json")
STATE_FILE = os.path.join(APP_DIR, "state.json")
BACKUP_DIR = os.path.join(APP_DIR, "backups")
LOG_FILE = os.path.expanduser("~/Library/Logs/outlook-dav-sync.log")
KEYCHAIN_SERVICE = "outlook-dav-sync"

NS = {"d": "DAV:", "c": "urn:ietf:params:xml:ns:caldav", "a": "urn:ietf:params:xml:ns:carddav"}
DEFAULT_CONFIG = {
    "username": "",
    "calendar_url": "",
    "addressbook_url": "",
    "outlook_calendar": "CalDAV Calendar",
    "outlook_addressbook": "CardDAV Contacts",
    "sync_calendar": True,
    "sync_contacts": True,
    # Refuse a run that would delete more than this share of items (and more than 5).
    "max_delete_fraction": 0.2,
}

log = logging.getLogger("davsync")


class SyncError(Exception):
    pass


class DavHTTPError(SyncError):
    def __init__(self, code, msg):
        super().__init__(msg)
        self.code = code


# --------------------------------------------------------------------------- helpers

def now_tag():
    return datetime.datetime.now().strftime("%Y%m%d-%H%M%S")


def safe_name(s):
    s = re.sub(r"[^A-Za-z0-9._-]", "_", s or "")
    return s[:80] or uuid.uuid4().hex


def unfold(text):
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    out = []
    for ln in text.split("\n"):
        if ln[:1] in (" ", "\t") and out:
            out[-1] += ln[1:]
        else:
            out.append(ln)
    return [ln for ln in out if ln != ""]


def fold(line):
    b = line.encode("utf-8")
    if len(b) <= 75:
        return line
    parts, cur, size = [], "", 0
    for ch in line:
        n = len(ch.encode("utf-8"))
        if size + n > (75 if not parts else 74):
            parts.append(cur)
            cur, size = "", 0
        cur += ch
        size += n
    parts.append(cur)
    return "\r\n ".join(parts)


def join_lines(lines):
    return "\r\n".join(fold(ln) for ln in lines) + "\r\n"


def split_prop(line):
    """Split 'NAME;PARAM="a:b":value' at the first colon outside quotes."""
    q = False
    for i, ch in enumerate(line):
        if ch == '"':
            q = not q
        elif ch == ":" and not q:
            return line[:i], line[i + 1:]
    return line, ""


def get_uid(text):
    for ln in unfold(text):
        name, val = split_prop(ln)
        if name.split(";")[0].upper() == "UID":
            return val.strip()
    return None


VOLATILE = ("DTSTAMP", "LAST-MODIFIED", "SEQUENCE", "PRODID", "REV", "X-ENTOURAGE", "UID",
            "X-MICROSOFT-CDO-INSTTYPE", "METHOD", "X-MICROSOFT-EXCHANGE")
# Outlook-internal identifiers that must not be stored on the server.
STRIP = ("X-ENTOURAGE", "X-MICROSOFT-EXCHANGE")


def fingerprint(text):
    """Content hash that ignores timestamps, ids and reminders, which Outlook rewrites on its own
    (e.g. it drops the VALARM of past events shortly after import)."""
    keep, in_alarm = [], False
    for ln in unfold(text):
        up = ln.upper()
        if up == "BEGIN:VALARM":
            in_alarm = True
        if in_alarm:
            in_alarm = up != "END:VALARM"
            continue
        if prop_name(ln).startswith(VOLATILE):
            continue
        keep.append(ln)
    return hashlib.sha256("\n".join(sorted(keep)).encode("utf-8")).hexdigest()


def prop_name(line):
    return split_prop(line)[0].split(";")[0].upper()


def bare_name(line):
    """Property name without a vCard group prefix (item1.EMAIL -> EMAIL)."""
    return prop_name(line).split(".")[-1]


def utc_stamp():
    return datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def parse_cal(text):
    """-> (VCALENDAR-level property lines, [(component name, lines incl. BEGIN/END)])."""
    head, comps, cur, depth, cname = [], [], None, 0, None
    for ln in unfold(text):
        up = ln.upper()
        if up in ("BEGIN:VCALENDAR", "END:VCALENDAR"):
            continue
        if cur is None:
            if up.startswith("BEGIN:"):
                cur, depth, cname = [ln], 1, up[6:]
            else:
                head.append(ln)
            continue
        cur.append(ln)
        if up.startswith("BEGIN:"):
            depth += 1
        elif up.startswith("END:"):
            depth -= 1
            if depth == 0:
                comps.append((cname, cur))
                cur = None
    return head, comps


def top_props(comp):
    """Property lines directly on a component (not inside its VALARMs etc.)."""
    out, depth = [], 0
    for ln in comp[1:-1]:
        up = ln.upper()
        if up.startswith("BEGIN:"):
            depth += 1
        elif up.startswith("END:"):
            depth -= 1
        elif depth == 0:
            out.append(ln)
    return out


def comp_value(comp, name):
    for ln in top_props(comp):
        if prop_name(ln) == name:
            return split_prop(ln)[1]
    return None


# What Outlook actually edits. Everything else on the server copy (attendees, organizer,
# alarms, HTML description, Apple/Google extensions) is preserved when pushing.
EVENT_MANAGED = {"SUMMARY", "LOCATION", "DESCRIPTION", "DTSTART", "DTEND", "DURATION",
                 "RRULE", "RDATE", "EXDATE", "RECURRENCE-ID"}
CARD_MANAGED = {"FN", "N", "NICKNAME", "ORG", "TITLE", "TEL", "EMAIL", "ADR", "URL", "NOTE", "BDAY"}


def _merge_vevent(srv, ol):
    ol_props = [l for l in top_props(ol) if prop_name(l) in EVENT_MANAGED]
    squash = lambda v: re.sub(r"(\\n|\s)+", "", v or "")
    desc_changed = squash(comp_value(srv, "DESCRIPTION")) != squash(comp_value(ol, "DESCRIPTION"))
    out, depth = [srv[0]], 0
    for ln in srv[1:-1]:
        up = ln.upper()
        if up.startswith("BEGIN:"):
            depth += 1
        elif up.startswith("END:"):
            depth -= 1
        elif depth == 0:
            name = prop_name(ln)
            if name in EVENT_MANAGED or name in ("DTSTAMP", "LAST-MODIFIED"):
                continue
            if name == "X-ALT-DESC" and desc_changed:
                continue
        out.append(ln)
    stamp = utc_stamp()
    return out + ["DTSTAMP:" + stamp, "LAST-MODIFIED:" + stamp] + ol_props + [srv[-1]]


def merge_ics(server_text, outlook_text, uid):
    """Server copy + the fields Outlook manages, taken from Outlook's export."""
    ol_clean = outlook_ics_to_server(outlook_text, uid)
    if not server_text:
        return ol_clean
    s_head, s_comps = parse_cal(server_text)
    _, o_comps = parse_cal(ol_clean)

    def rid(c):
        v = comp_value(c, "RECURRENCE-ID")
        return re.sub(r"[^0-9T]", "", v)[:15] if v else None

    s_events = {rid(c): c for n, c in s_comps if n == "VEVENT"}
    tzids = {comp_value(c, "TZID") for n, c in s_comps if n == "VTIMEZONE"}
    comps = [c for n, c in s_comps if n != "VEVENT"]
    comps += [c for n, c in o_comps if n == "VTIMEZONE" and comp_value(c, "TZID") not in tzids]
    for n, c in o_comps:
        if n == "VEVENT":
            r = rid(c)
            comps.append(_merge_vevent(s_events[r], c) if r in s_events else c)
    lines = ["BEGIN:VCALENDAR"] + [l for l in s_head if prop_name(l) != "METHOD"]
    lines += [l for c in comps for l in c] + ["END:VCALENDAR"]
    return join_lines(lines)


def merge_vcf(server_text, outlook_text, uid):
    ol = outlook_vcf_to_server(outlook_text, uid)
    if not server_text:
        return ol
    ol_props = [l for l in unfold(ol) if bare_name(l) not in ("BEGIN", "END", "VERSION", "PRODID", "UID")]
    managed = {bare_name(l) for l in ol_props} | CARD_MANAGED
    out = [l for l in unfold(server_text) if bare_name(l) not in managed | {"END", "REV"}]
    return join_lines(out + ol_props + ["REV:" + utc_stamp(), "END:VCARD"])


def outlook_ics_to_server(text, uid):
    """Clean Outlook's iCalendar export for CalDAV storage."""
    out, ev = [], None
    for ln in unfold(text):
        name = split_prop(ln)[0].split(";")[0].upper()
        if ev is None and name == "METHOD":
            continue  # RFC 4791: stored objects must not carry METHOD
        if name.startswith(STRIP) and name not in ("X-ENTOURAGE-CFTIMEZONE", "X-ENTOURAGE-TZID"):
            continue
        if ln.upper() == "BEGIN:VEVENT":
            ev = [ln]
            continue
        if ev is not None:
            ev.append(ln)
            if ln.upper() == "END:VEVENT":
                out.extend(_fix_event(ev, uid))
                ev = None
            continue
        out.append(ln)
    return join_lines(out)


def _fix_event(ev, uid):
    allday = any(l.upper().replace(" ", "") == "X-MICROSOFT-CDO-ALLDAYEVENT:TRUE" for l in ev)
    res = []
    for ln in ev:
        name, val = split_prop(ln)
        pname = name.split(";")[0].upper()
        if pname == "UID" and uid:
            ln = "UID:" + uid
        elif allday and pname in ("DTSTART", "DTEND", "RECURRENCE-ID", "EXDATE"):
            dates = ",".join(v.strip()[:8] for v in val.split(","))
            ln = f"{pname};VALUE=DATE:{dates}"
        res.append(ln)
    return res


def outlook_vcf_to_server(text, uid):
    """Normalise Outlook's lower-case vCard export to clean vCard 3.0."""
    out, has_uid = [], False
    for ln in unfold(text):
        name, val = split_prop(ln)
        parts = name.split(";")
        pname = parts[0].upper()
        if pname.startswith(STRIP):
            continue
        params = []
        for p in parts[1:]:
            if p.lower().startswith("charset="):
                continue
            if "=" in p:
                k, v = p.split("=", 1)
                params.append(k.upper() + "=" + v)
            else:
                params.append(p.upper())
        if pname in ("BEGIN", "END"):
            val = val.upper()
        if pname == "UID":
            val, has_uid = uid, True
        out.append(";".join([pname] + params) + ":" + val)
    if not has_uid:
        idx = next((i for i, l in enumerate(out) if l.upper().startswith("VERSION:")), 0)
        out.insert(idx + 1, "UID:" + uid)
    return join_lines(out)


def backup(kind, label, text, sub=None):
    d = os.path.join(BACKUP_DIR, sub or datetime.date.today().isoformat())
    os.makedirs(d, exist_ok=True)
    ext = ".ics" if kind == "events" else ".vcf"
    path = os.path.join(d, f"{now_tag()}-{kind}-{safe_name(label)}{ext}")
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path


# --------------------------------------------------------------------------- server

class Dav:
    def __init__(self, user, password):
        self.auth = "Basic " + base64.b64encode(f"{user}:{password}".encode("utf-8")).decode()
        self.ctx = ssl.create_default_context()

    @staticmethod
    def key(href):
        return unquote(urlparse(href).path)

    @staticmethod
    def url(base, key):
        return urljoin(base, quote(key, safe="/@:+~!$&'()*,;="))

    def req(self, method, url, body=None, headers=None):
        h = {"Authorization": self.auth, "User-Agent": f"outlook-dav-sync/{__version__}"}
        h.update(headers or {})
        data = body.encode("utf-8") if isinstance(body, str) else body
        r = urllib.request.Request(url, data=data, method=method, headers=h)
        try:
            with urllib.request.urlopen(r, context=self.ctx, timeout=90) as resp:
                return resp.status, resp.headers, resp.read()
        except urllib.error.HTTPError as e:
            raise DavHTTPError(e.code, f"{method} {url} -> HTTP {e.code}") from None
        except urllib.error.URLError as e:
            raise SyncError(f"{method} {url} -> {e.reason}") from None

    def propfind(self, url, depth, props):
        body = ('<?xml version="1.0" encoding="utf-8"?><d:propfind xmlns:d="DAV:" '
                'xmlns:c="urn:ietf:params:xml:ns:caldav" xmlns:a="urn:ietf:params:xml:ns:carddav">'
                f'<d:prop>{props}</d:prop></d:propfind>')
        _, _, data = self.req("PROPFIND", url, body,
                              {"Depth": str(depth), "Content-Type": "application/xml; charset=utf-8"})
        out = []
        for resp in ET.fromstring(data).findall("d:response", NS):
            href = resp.findtext("d:href", "", NS)
            props_ok = [ps.find("d:prop", NS) for ps in resp.findall("d:propstat", NS)
                        if " 200" in (ps.findtext("d:status", "", NS) or "")]
            out.append((href, [p for p in props_ok if p is not None]))
        return out

    def list_items(self, coll_url):
        """{key: etag} for every non-collection member."""
        coll_key = self.key(coll_url).rstrip("/")
        items = {}
        for href, props in self.propfind(coll_url, 1, "<d:getetag/><d:resourcetype/>"):
            key = self.key(href)
            if key.rstrip("/") == coll_key:
                continue
            etag, is_coll = None, False
            for p in props:
                etag = p.findtext("d:getetag", None, NS) or etag
                rt = p.find("d:resourcetype", NS)
                if rt is not None and rt.find("d:collection", NS) is not None:
                    is_coll = True
            if etag and not is_coll:
                items[key] = etag
        return items

    def fetch(self, coll_url, keys, kind):
        """{key: (etag, text)} via multiget REPORT, falling back to GET."""
        out = {}
        if kind == "events":
            root, data_tag = "c:calendar-multiget", "c:calendar-data"
        else:
            root, data_tag = "a:addressbook-multiget", "a:address-data"
        for i in range(0, len(keys), 50):
            chunk = keys[i:i + 50]
            hrefs = "".join(f"<d:href>{xml_escape(quote(k, safe='/@:+~'))}</d:href>" for k in chunk)
            body = ('<?xml version="1.0" encoding="utf-8"?>'
                    f'<{root} xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav" '
                    'xmlns:a="urn:ietf:params:xml:ns:carddav">'
                    f'<d:prop><d:getetag/><{data_tag}/></d:prop>{hrefs}</{root}>')
            try:
                _, _, data = self.req("REPORT", coll_url, body,
                                      {"Depth": "1", "Content-Type": "application/xml; charset=utf-8"})
                for resp in ET.fromstring(data).findall("d:response", NS):
                    key = self.key(resp.findtext("d:href", "", NS))
                    for ps in resp.findall("d:propstat", NS):
                        p = ps.find("d:prop", NS)
                        if p is None:
                            continue
                        txt = p.findtext(data_tag, None, NS)
                        if txt:
                            out[key] = (p.findtext("d:getetag", None, NS), txt)
            except (DavHTTPError, ET.ParseError) as e:
                log.debug("multiget failed (%s), falling back to GET", e)
        for k in keys:
            if k not in out:
                try:
                    _, h, data = self.req("GET", self.url(coll_url, k))
                    out[k] = (h.get("ETag"), data.decode("utf-8", "replace"))
                except DavHTTPError as e:
                    if e.code != 404:
                        raise
        return out

    def etag_of(self, url):
        for _, props in self.propfind(url, 0, "<d:getetag/>"):
            for p in props:
                e = p.findtext("d:getetag", None, NS)
                if e:
                    return e
        return None

    def put(self, url, text, ctype, etag=None, create=False):
        h = {"Content-Type": ctype}
        if create:
            h["If-None-Match"] = "*"
        elif etag:
            h["If-Match"] = etag
        _, hd, _ = self.req("PUT", url, text, h)
        return hd.get("ETag") or self.etag_of(url)

    def delete(self, url, etag=None):
        try:
            self.req("DELETE", url, None, {"If-Match": etag} if etag else {})
        except DavHTTPError as e:
            if e.code != 404:
                raise


# --------------------------------------------------------------------------- outlook

class Outlook:
    def __init__(self):
        self.js = os.path.join(HERE, "outlook.js")

    def call(self, *args):
        r = subprocess.run(["osascript", "-l", "JavaScript", self.js, *map(str, args)],
                           capture_output=True, text=True, timeout=900)
        if r.returncode != 0:
            raise SyncError(f"Outlook bridge failed ({args[0]}): {r.stderr.strip()}")
        res = json.loads(r.stdout)
        if not res.get("ok"):
            raise SyncError(f"Outlook: {res.get('error')}")
        return res["result"]


# --------------------------------------------------------------------------- sync engine

class Kind:
    """Per-type plumbing (calendar events vs. contacts)."""

    def __init__(self, name, ol, folder_id, url):
        self.name, self.ol, self.fid, self.url = name, ol, folder_id, url
        ev = name == "events"
        self.ext = ".ics" if ev else ".vcf"
        self.ctype = "text/calendar; charset=utf-8" if ev else "text/vcard; charset=utf-8"
        self.cmd = ("listEvents", "exportEvents", "importEvent", "deleteEvent") if ev else \
                   ("listContacts", "exportContacts", "importContact", "deleteContact")

    def ol_list(self):
        return {str(k): v for k, v in self.ol.call(self.cmd[0], self.fid).items()}

    def ol_export(self, ids):
        out = {}
        ids = list(ids)
        for i in range(0, len(ids), 40):
            out.update({str(k): v for k, v in self.ol.call(self.cmd[1], self.fid, *ids[i:i + 40]).items()})
        return out

    def ol_import(self, text):
        fd, path = tempfile.mkstemp(suffix=self.ext, prefix="davsync-")
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write(text)
            r = self.ol.call(self.cmd[2], self.fid, path)
        finally:
            os.unlink(path)
        if not r.get("master"):
            raise SyncError("Outlook imported nothing")
        return str(r["master"]), r["mod"]

    def ol_delete(self, oid):
        self.ol.call(self.cmd[3], self.fid, oid)

    def to_server(self, text, uid, server_text=None):
        merge = merge_ics if self.name == "events" else merge_vcf
        return merge(server_text, text, uid)


class Syncer:
    def __init__(self, cfg, dav, ol, dry=False, allow_mass_delete=False):
        self.cfg, self.dav, self.ol, self.dry, self.allow_mass = cfg, dav, ol, dry, allow_mass_delete
        self.state = json.load(open(STATE_FILE)) if os.path.exists(STATE_FILE) else {}
        self.stats = {}

    def save(self):
        if self.dry:
            return
        tmp = STATE_FILE + ".tmp"
        with open(tmp, "w") as f:
            json.dump(self.state, f, indent=1)
        os.replace(tmp, STATE_FILE)

    def run(self):
        kinds = []
        if self.cfg["sync_calendar"] and self.cfg["calendar_url"]:
            fid = self.ol.call("ensure", "calendar", self.cfg["outlook_calendar"]) if not self.dry else \
                self.state.get("_folders", {}).get("events")
            kinds.append(Kind("events", self.ol, fid, self.cfg["calendar_url"]))
        if self.cfg["sync_contacts"] and self.cfg["addressbook_url"]:
            fid = self.ol.call("ensure", "addressbook", self.cfg["outlook_addressbook"]) if not self.dry else \
                self.state.get("_folders", {}).get("contacts")
            kinds.append(Kind("contacts", self.ol, fid, self.cfg["addressbook_url"]))
        for k in kinds:
            folders = self.state.setdefault("_folders", {})
            if k.fid is None:
                log.info("[%s] dry run: Outlook folder not created yet, treating it as empty", k.name)
            elif folders.get(k.name) not in (None, k.fid):
                # Folder was recreated in Outlook: old ids are meaningless. Start this kind over.
                log.warning("[%s] Outlook folder changed (%s -> %s); resetting its sync state",
                            k.name, folders.get(k.name), k.fid)
                self.state[k.name] = {}
            if k.fid is not None:
                folders[k.name] = k.fid
            self.sync_kind(k)
            self.save()
        log.info("done: %s", json.dumps(self.stats))

    def sync_kind(self, k):
        st = self.state.setdefault(k.name, {})
        server = self.dav.list_items(k.url)
        ol = k.ol_list() if k.fid is not None else {}
        st_by_ol = {rec["ol"]: key for key, rec in st.items() if rec.get("ol")}
        coll_key = Dav.key(k.url).rstrip("/") + "/"

        if not st and server and not self.dry:
            sub = f"initial-{now_tag()}"
            for key, (_, text) in self.dav.fetch(k.url, list(server), k.name).items():
                backup(k.name, os.path.basename(key), text, sub)
            log.info("[%s] first run: backed up %d server items to %s", k.name, len(server),
                     os.path.join(BACKUP_DIR, sub))

        # ---- Outlook-side changes: confirm by content, not just modification date
        maybe = {rec["ol"]: key for key, rec in st.items()
                 if rec.get("ol") in ol and ol[rec["ol"]] > rec.get("olmod", 0)}
        ol_text = k.ol_export(maybe) if maybe else {}
        o_changed = set()
        for oid, key in maybe.items():
            if fingerprint(ol_text[oid]) != st[key].get("fp"):
                o_changed.add(key)
            elif not self.dry:
                st[key]["olmod"] = ol[oid]  # Outlook touched the timestamp only

        plan = []
        for key, rec in st.items():
            oid = rec.get("ol")
            s_has, o_has = key in server, bool(oid) and oid in ol
            s_chg = s_has and server[key] != rec.get("etag")
            o_chg = key in o_changed
            if not oid:  # earlier import into Outlook failed; never delete on server because of it
                plan.append(("forget", key) if not s_has else ("pull", key) if s_chg else None)
            elif not s_has and not o_has:
                plan.append(("forget", key))
            elif not s_has:
                plan.append(("push_recreate", key) if o_chg else ("del_outlook", key))
            elif not o_has:
                plan.append(("pull", key) if s_chg else ("del_server", key))
            elif s_chg and o_chg:
                plan.append(("conflict", key))
            elif s_chg:
                plan.append(("pull", key))
            elif o_chg:
                plan.append(("push", key))
        plan = [p for p in plan if p]
        plan += [("pull_new", key) for key in server if key not in st]
        plan += [("push_new", oid) for oid in ol if oid not in st_by_ol]

        counts = {}
        for a, _ in plan:
            counts[a] = counts.get(a, 0) + 1
        self.stats[k.name] = counts
        if not plan:
            log.info("[%s] up to date (%d server, %d Outlook)", k.name, len(server), len(ol))
            return
        log.info("[%s] plan: %s", k.name, counts)

        dels = counts.get("del_outlook", 0) + counts.get("del_server", 0)
        limit = max(5, int(self.cfg["max_delete_fraction"] * max(len(st), 1)))
        if dels > limit and not self.allow_mass:
            raise SyncError(f"[{k.name}] refusing to delete {dels} items in one run (limit {limit}). "
                            "If this is intended, run once with --allow-mass-delete.")
        if self.dry:
            for a, key in plan:
                log.info("  would %s %s", a, key)
            return

        need = [key for a, key in plan if a in ("pull", "pull_new", "conflict", "push")]
        srv = self.dav.fetch(k.url, need, k.name) if need else {}
        need_ol = [st[key]["ol"] for a, key in plan
                   if a in ("push", "push_recreate", "conflict", "del_outlook") and st[key]["ol"] not in ol_text]
        need_ol += [oid for a, oid in plan if a == "push_new"]
        if need_ol:
            ol_text.update(k.ol_export(need_ol))

        for action, key in plan:
            try:
                self.apply(k, st, action, key, server, srv, ol, ol_text, coll_key)
            except DavHTTPError as e:
                if e.code == 412:
                    log.warning("[%s] %s %s: changed on server meanwhile, will retry next run", k.name, action, key)
                    continue
                log.error("[%s] %s %s failed: %s", k.name, action, key, e)
            except SyncError as e:
                log.error("[%s] %s %s failed: %s", k.name, action, key, e)
            self.save()

    def apply(self, k, st, action, key, server, srv, ol, ol_text, coll_key):
        url = Dav.url(k.url, key) if action != "push_new" else None
        rec = st.get(key)

        if action == "forget":
            st.pop(key, None)

        elif action in ("pull", "pull_new", "conflict"):
            if key not in srv:
                log.warning("[%s] %s vanished from server", k.name, key)
                return
            etag, text = srv[key]
            if action == "conflict":
                path = backup(k.name, "conflict-outlook-" + os.path.basename(key), ol_text.get(rec["ol"], ""))
                log.warning("[%s] conflict on %s: server version kept, Outlook version saved to %s",
                            k.name, key, path)
            if rec and rec.get("ol") and rec["ol"] in ol:
                k.ol_delete(rec["ol"])
            new = {"etag": etag or server.get(key), "uid": get_uid(text), "ol": None, "olmod": 0}
            st[key] = new
            try:
                oid, mod = k.ol_import(text)
            except SyncError as e:
                backup(k.name, "import-failed-" + os.path.basename(key), text)
                raise SyncError(f"import into Outlook failed: {e}")
            new.update(ol=oid, olmod=mod, fp=fingerprint(k.ol_export([oid])[oid]))
            log.info("[%s] %s %s -> Outlook #%s", k.name, action, key, oid)

        elif action in ("push", "push_recreate"):
            base = srv.get(key, (None, None))[1] if action == "push" else None
            text = k.to_server(ol_text[rec["ol"]], rec.get("uid"), base)
            etag = self.dav.put(url, text, k.ctype,
                                            etag=None if action == "push_recreate" else rec.get("etag"),
                                            create=action == "push_recreate")
            rec.update(etag=etag, olmod=ol[rec["ol"]], fp=fingerprint(ol_text[rec["ol"]]))
            log.info("[%s] %s Outlook #%s -> %s", k.name, action, rec["ol"], key)

        elif action == "push_new":
            oid = key
            raw = ol_text[oid]
            uid = get_uid(raw) if k.name == "events" else None
            if not uid or (coll_key + safe_name(uid) + k.ext) in server or uid.startswith("BE000000-"):
                uid = str(uuid.uuid4()).upper()
            new_key = coll_key + safe_name(uid) + k.ext
            etag = self.dav.put(Dav.url(k.url, new_key), k.to_server(raw, uid), k.ctype, create=True)
            st[new_key] = {"etag": etag, "uid": uid, "ol": oid, "olmod": ol[oid], "fp": fingerprint(raw)}
            log.info("[%s] new in Outlook #%s -> %s", k.name, oid, new_key)

        elif action == "del_outlook":
            backup(k.name, "deleted-on-server-" + os.path.basename(key), ol_text.get(rec["ol"], ""))
            k.ol_delete(rec["ol"])
            st.pop(key)
            log.info("[%s] deleted on server -> removed Outlook #%s", k.name, rec["ol"])

        elif action == "del_server":
            got = self.dav.fetch(k.url, [key], k.name)
            if key in got:
                backup(k.name, "deleted-in-outlook-" + os.path.basename(key), got[key][1])
            self.dav.delete(url, rec.get("etag"))
            st.pop(key)
            log.info("[%s] deleted in Outlook -> removed %s", k.name, key)


# --------------------------------------------------------------------------- CLI

def load_config():
    os.makedirs(APP_DIR, exist_ok=True)
    if not os.path.exists(CONFIG_FILE):
        raise SyncError("Not configured yet. Run: davsync.py setup")
    cfg = dict(DEFAULT_CONFIG)
    cfg.update(json.load(open(CONFIG_FILE)))
    if not cfg["username"]:
        raise SyncError(f"Set 'username' in {CONFIG_FILE}")
    return cfg


def get_password(user):
    r = subprocess.run(["security", "find-generic-password", "-s", KEYCHAIN_SERVICE, "-a", user, "-w"],
                       capture_output=True, text=True)
    if r.returncode != 0:
        raise SyncError(f"No password in Keychain. Add it with:\n  security add-generic-password "
                        f"-s {KEYCHAIN_SERVICE} -a {user} -U -w")
    return r.stdout.rstrip("\n")


def collections(dav, start_url):
    """-> ([(url, name, can_hold_events)], [(url, name)]) found from any URL on the server."""
    def first_href(url, prop):
        for _, props in dav.propfind(url, 0, prop):
            for p in props:
                for el in p.iter("{DAV:}href"):
                    return urljoin(url, el.text.strip())
        return None

    principal = first_href(start_url, "<d:current-user-principal/>") or start_url
    cals, books = [], []
    for kind, prop in (("calendar", "<c:calendar-home-set/>"), ("addressbook", "<a:addressbook-home-set/>")):
        home = first_href(principal, prop)
        if not home:
            continue
        for href, props in dav.propfind(home, 1, "<d:displayname/><d:resourcetype/>"
                                                 "<c:supported-calendar-component-set/>"):
            rt, name, comps = None, "", None
            for p in props:
                rt = p.find("d:resourcetype", NS) if rt is None else rt
                name = p.findtext("d:displayname", "", NS) or name
                cs = p.find("c:supported-calendar-component-set", NS)
                if cs is not None:
                    comps = {c.get("name") for c in cs}
            if rt is None or rt.find(("c:" if kind == "calendar" else "a:") + kind, NS) is None:
                continue
            url = urljoin(home, href)
            if kind == "calendar":
                cals.append((url, name, comps is None or "VEVENT" in comps))
            else:
                books.append((url, name))
    return cals, books


def discover(dav, start_url):
    cals, books = collections(dav, start_url)
    print("calendars:")
    for url, name, ev in cals:
        print(f"  {url}   ({name}){'' if ev else '  [no events: tasks only]'}")
    print("address books:")
    for url, name in books:
        print(f"  {url}   ({name})")


def _ask(prompt, default=""):
    v = input(f"{prompt}{f' [{default}]' if default else ''}: ").strip()
    return v or default


def _choose(label, items):
    if not items:
        print(f"No {label} found on the server; {label} sync will be off.")
        return ""
    for i, it in enumerate(items, 1):
        print(f"  {i}) {it[1] or it[0]}   {it[0]}")
    print("  0) do not sync")
    while True:
        v = _ask(f"Which {label}", "1")
        if v.isdigit() and 0 <= int(v) <= len(items):
            return items[int(v) - 1][0] if int(v) else ""


def setup():
    os.makedirs(APP_DIR, exist_ok=True)
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_FILE):
        cfg.update(json.load(open(CONFIG_FILE)))
    print("Server: any CalDAV/CardDAV URL works, e.g. https://mail.example.com:2080/ for cPanel.")
    server = _ask("Server URL", cfg["calendar_url"] or cfg["addressbook_url"])
    if not server.startswith("https://"):
        raise SyncError("Use an https:// URL.")
    cfg["username"] = _ask("Username (usually your e-mail address)", cfg["username"])
    have = subprocess.run(["security", "find-generic-password", "-s", KEYCHAIN_SERVICE, "-a", cfg["username"]],
                          capture_output=True).returncode == 0
    if not have or _ask("A password is already in Keychain. Replace it? (y/N)", "n").lower().startswith("y"):
        print("Enter the account password (it goes straight into the macOS Keychain):")
        r = subprocess.run(["security", "add-generic-password", "-s", KEYCHAIN_SERVICE,
                            "-a", cfg["username"], "-U", "-w"])
        if r.returncode != 0:
            raise SyncError("Could not store the password in Keychain.")
    dav = Dav(cfg["username"], get_password(cfg["username"]))
    try:
        cals, books = collections(dav, server)
    except DavHTTPError as e:
        raise SyncError(f"Server refused the login or URL ({e}). Check username/password and URL.")
    cfg["calendar_url"] = _choose("calendar", [c for c in cals if c[2]])
    cfg["addressbook_url"] = _choose("address book", books)
    cfg["sync_calendar"], cfg["sync_contacts"] = bool(cfg["calendar_url"]), bool(cfg["addressbook_url"])
    if cfg["calendar_url"]:
        cfg["outlook_calendar"] = _ask("Calendar name in Outlook (On My Computer)", cfg["outlook_calendar"])
    if cfg["addressbook_url"]:
        cfg["outlook_addressbook"] = _ask("Address book name in Outlook (On My Computer)", cfg["outlook_addressbook"])
    if cfg["calendar_url"] and cfg["outlook_calendar"] == cfg["outlook_addressbook"]:
        raise SyncError("Outlook needs different names for the calendar and the address book.")
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)
    print(f"Saved {CONFIG_FILE}")


def setup_logging(verbose, quiet=False):
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    fh = logging.handlers.RotatingFileHandler(LOG_FILE, maxBytes=2_000_000, backupCount=3)
    fh.setFormatter(fmt)
    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    log.addHandler(fh)
    if not quiet:  # launchd passes --quiet: the log file is enough there
        log.addHandler(sh)
    log.setLevel(logging.DEBUG if verbose else logging.INFO)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=["run", "setup", "discover", "status"])
    ap.add_argument("--version", action="version", version=__version__)
    ap.add_argument("--dry-run", action="store_true", help="show what would change, change nothing")
    ap.add_argument("--allow-mass-delete", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    ap.add_argument("-q", "--quiet", action="store_true", help="log to file only (used by launchd)")
    a = ap.parse_args()
    setup_logging(a.verbose, a.quiet)
    try:
        if a.command == "setup":
            setup()
            return 0
        cfg = load_config()
        if a.command == "status":
            st = json.load(open(STATE_FILE)) if os.path.exists(STATE_FILE) else {}
            print(json.dumps({k: len(v) for k, v in st.items() if not k.startswith("_")}, indent=1))
            return 0
        dav = Dav(cfg["username"], get_password(cfg["username"]))
        if a.command == "discover":
            discover(dav, cfg["calendar_url"] or cfg["addressbook_url"])
            return 0
        lock = open(os.path.join(APP_DIR, "lock"), "w")
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            log.info("another sync is running; skipping")
            return 0
        ol = Outlook()
        try:
            Syncer(cfg, dav, ol, dry=a.dry_run, allow_mass_delete=a.allow_mass_delete).run()
        except SyncError as e:
            if "OUTLOOK_NOT_RUNNING" in str(e):
                log.info("Outlook is not running; skipping")
                return 0
            raise
        return 0
    except SyncError as e:
        log.error("%s", e)
        return 1


if __name__ == "__main__":
    sys.exit(main())
