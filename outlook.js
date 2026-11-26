// Bridge between davsync.py and Legacy Outlook for Mac (JavaScript for Automation).
// Usage: osascript -l JavaScript outlook.js <command> [args...]
// Always prints one JSON object: {"ok": true, "result": ...} or {"ok": false, "error": "..."}.

function run(argv) {
  const o = Application("Microsoft Outlook");
  const cmd = argv[0];
  const args = argv.slice(1);
  try {
    if (cmd !== "running" && !o.running()) {
      return JSON.stringify({ ok: false, error: "OUTLOOK_NOT_RUNNING" });
    }
    return JSON.stringify({ ok: true, result: H[cmd](o, args) });
  } catch (e) {
    return JSON.stringify({ ok: false, error: String(e) });
  }
}

// Folders of an account (e.g. Microsoft 365) have account.name(); "On My Computer" folders throw.
function isLocal(folder) {
  try { folder.account.name(); return false; } catch (e) { return true; }
}

function ms(d) { return d ? d.getTime() : 0; }

const H = {
  running(o) { return o.running(); },

  // ensure <calendar|addressbook> <name>  -> folder id (created under "On My Computer" if missing)
  ensure(o, [kind, name]) {
    const isCal = kind === "calendar";
    const all = isCal ? o.calendars() : o.addressBooks();
    const hit = all.filter(f => f.name() === name && isLocal(f));
    if (hit.length) return hit[0].id();
    const made = o.make({ new: isCal ? "calendar" : "addressBook", at: o.onMyComputer, withProperties: { name: name } });
    return made.id();
  },

  // removeFolder <calendar|addressbook> <name>  -> number of local folders deleted (used by uninstall)
  removeFolder(o, [kind, name]) {
    const all = kind === "calendar" ? o.calendars() : o.addressBooks();
    const hit = all.filter(f => f.name() === name && isLocal(f));
    hit.forEach(f => o.delete(f));
    return hit.length;
  },

  // listEvents <calId> -> { masterId: latestModificationMs } (exceptions folded into their series)
  listEvents(o, [calId]) {
    const evs = o.calendars.byId(+calId).calendarEvents;
    const ids = evs.id();
    if (!ids.length) return {};
    const occ = evs.isOccurrence();
    const mods = evs.modificationDate();
    const g = {};
    ids.forEach((id, i) => {
      let key = id;
      if (occ[i]) {
        try { const m = evs.byId(id).master(); key = m ? m.id() : null; } catch (e) { key = null; }
        if (key === null) return; // orphan exception: ignore
      }
      const t = ms(mods[i]);
      if (!(key in g) || g[key] < t) g[key] = t;
    });
    return g;
  },

  // exportEvents <calId> <id>... -> { id: icsText }
  exportEvents(o, [calId, ...ids]) {
    const evs = o.calendars.byId(+calId).calendarEvents;
    const r = {};
    ids.forEach(id => { r[id] = evs.byId(+id).icalendarData(); });
    return r;
  },

  // importEvent <calId> <path.ics> -> { master, mod, n }
  importEvent(o, [calId, path]) {
    const r = o.importIcs(Path(path), { to: o.calendars.byId(+calId) });
    let master = null, mod = 0;
    r.forEach(e => {
      if (master === null && !e.isOccurrence()) master = e.id();
      mod = Math.max(mod, ms(e.modificationDate()));
    });
    if (master === null && r.length) master = r[0].id();
    return { master: master, mod: mod, n: r.length };
  },

  // deleteEvent <calId> <id>
  deleteEvent(o, [calId, id]) {
    o.delete(o.calendars.byId(+calId).calendarEvents.byId(+id));
    return true;
  },

  // listContacts <abId> -> { id: modificationMs }
  listContacts(o, [abId]) {
    const cs = o.addressBooks.byId(+abId).contacts;
    const ids = cs.id();
    if (!ids.length) return {};
    const mods = cs.modificationDate();
    const g = {};
    ids.forEach((id, i) => { g[id] = ms(mods[i]); });
    return g;
  },

  // exportContacts <abId> <id>... -> { id: vcfText }
  exportContacts(o, [abId, ...ids]) {
    const cs = o.addressBooks.byId(+abId).contacts;
    const r = {};
    ids.forEach(id => { r[id] = cs.byId(+id).vcardData(); });
    return r;
  },

  // importContact <abId> <path.vcf> -> { master, mod, n }
  importContact(o, [abId, path]) {
    const r = o.importVcf(Path(path), { to: o.addressBooks.byId(+abId) });
    if (!r.length) return { master: null, mod: 0, n: 0 };
    return { master: r[0].id(), mod: ms(r[0].modificationDate()), n: r.length };
  },

  // deleteContact <abId> <id>
  deleteContact(o, [abId, id]) {
    o.delete(o.addressBooks.byId(+abId).contacts.byId(+id));
    return true;
  },
};
