import { useState, useEffect, useMemo, useRef } from "react";
import { Plus, X, Search, Clock, Wrench, MapPin, User, ChevronRight, AlertTriangle, CheckCircle2, PackagePlus, Loader2, Lock, ShieldCheck, KeyRound, UserPlus, Trash2, ArrowLeft, Menu, LayoutGrid, List as ListIcon, LogOut, Pencil, BarChart3, ChevronDown, Check, Download, Cog, ClipboardCopy, MessageSquare, FileText, Eye, EyeOff, Printer } from "lucide-react";
import * as XLSX from "xlsx";
import {
  fetchLocks as apiFetchLocks,
  acquireLock as apiAcquireLock,
  releaseLock as apiReleaseLock,
} from "./lib/apiStorage";
// jsPDF pulls in html2canvas + dompurify (unused here — we only render
// tables/text, never HTML), so it's loaded on demand rather than in the
// main bundle.

// ---------- Status / condition definitions (from the shop's own legend) ----------
const STATUSES = [
  { code: "M/C", label: "Fixed on Machine" },
  { code: "D.RACK", label: "Avbl. in Dept Rack" },
  { code: "MRTN", label: "Avbl. in Maintenance Area" },
  { code: "STRS", label: "Avbl. in Stores" },
  { code: "URWP", label: "Under Repair — with Party" },
  { code: "RDY.WP", label: "Ready — with Party" },
  { code: "SCRAP", label: "Scrapped / Written Off" },
];

const CONDITIONS = {
  NEW: { label: "New", fg: "#2F5233", bg: "#DCEEDD", ring: "#7BAE7F" },
  OKAY: { label: "Okay", fg: "#1E4E8C", bg: "#DCE6F6", ring: "#6E93C4" },
  REPAIRED: { label: "Repaired", fg: "#28454C", bg: "#DCEBEE", ring: "#6FA3AC" },
  "IN REPAIR": { label: "In Repair", fg: "#7A4A12", bg: "#F6E3C7", ring: "#D9A34B" },
  "NOT OKAY": { label: "Not Okay", fg: "#7A2620", bg: "#F6D9D6", ring: "#C1443C" },
  SCRAP: { label: "Scrap", fg: "#57534E", bg: "#E7E5E0", ring: "#A8A29E" },
};

const STATUS_MAP = Object.fromEntries(STATUSES.map((s) => [s.code, s.label]));

// Real-world classification the shop actually thinks in terms of — each
// drum lands in exactly one bucket, checked top to bottom, so the counts
// always add up to the total. This is distinct from the raw status/condition
// tags: e.g. "Okay" alone doesn't say whether a drum is new, repaired, or
// currently running, but this does.
const CLASSIFICATIONS = [
  { key: "SCRAPPED", label: "Scrapped", match: (d) => d.status === "SCRAP", fg: "#57534E", bg: "#E7E5E0" },
  { key: "NOT_OKAY", label: "Not Okay", match: (d) => d.condition === "NOT OKAY", fg: "#7A2620", bg: "#F6D9D6" },
  { key: "RUNNING", label: "Running — On Machine", match: (d) => d.status === "M/C", fg: "#1E4E8C", bg: "#DCE6F6" },
  { key: "IN_REPAIR", label: "In Repair — With Party", match: (d) => d.status === "URWP", fg: "#7A4A12", bg: "#F6E3C7" },
  { key: "READY_PARTY", label: "Ready — With Party", match: (d) => d.status === "RDY.WP", fg: "#28454C", bg: "#DCEBEE" },
  { key: "MAINTENANCE", label: "In Maintenance Area", match: (d) => d.status === "MRTN", fg: "#5A5045", bg: "#EDE7DC" },
  { key: "NEW_STORES", label: "New — In Stores", match: (d) => d.status === "STRS" && d.condition === "NEW", fg: "#2F5233", bg: "#DCEEDD" },
  { key: "REPAIRED_RACK", label: "Repaired — In Dept Rack", match: (d) => d.status === "D.RACK" && d.condition === "REPAIRED", fg: "#28454C", bg: "#DCEBEE" },
  { key: "OKAY_RACK", label: "Okay — In Dept Rack", match: (d) => d.status === "D.RACK" && d.condition === "OKAY", fg: "#1E4E8C", bg: "#DCE6F6" },
  { key: "OTHER", label: "Other", match: () => true, fg: "#6B7580", bg: "#F2F3F4" }, // catch-all so totals always reconcile
];

function classifyDrum(d) {
  return CLASSIFICATIONS.find((c) => c.match(d)).key;
}

// Buckets that count as "ready to use right now" for par-level tracking —
// new/available stock, not drums currently running or off with a party.
const USABLE_BUCKETS = ["NEW_STORES", "REPAIRED_RACK", "OKAY_RACK"];

// Drum styles, as used on the shop floor. Add more any time from the Admin panel.
const DRUM_STYLES = ["16-13", "19-13", "20-58-15", "20-75-15", "22-58-15", "24-15", "25-65-23", "28-15", "28-25"];

function uid(prefix = "D") {
  return `${prefix}-${Math.random().toString(36).slice(2, 7).toUpperCase()}`;
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

function fmtDate(d) {
  if (!d) return "—";
  const dt = new Date(d + "T00:00:00");
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function diffDays(dateA, dateB) {
  const a = Date.parse(dateA + "T00:00:00");
  const b = Date.parse(dateB + "T00:00:00");
  if (isNaN(a) || isNaN(b)) return null;
  return Math.round((b - a) / 86400000);
}

// Below this width the app behaves like a phone (bottom sheets, single
// column). Above it, panels center as normal dialogs and content gets
// breathing room instead of stretching edge to edge.
const DESKTOP_BREAKPOINT = 860;

function useIsDesktop() {
  const [isDesktop, setIsDesktop] = useState(
    typeof window !== "undefined" ? window.innerWidth >= DESKTOP_BREAKPOINT : false
  );
  useEffect(() => {
    function handleResize() {
      setIsDesktop(window.innerWidth >= DESKTOP_BREAKPOINT);
    }
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);
  return isDesktop;
}

// Rebuilds a drum's derived fields (status, condition, party, lastFixedDate,
// totalLifeDays, repairCount) by replaying its remaining history events in
// chronological order. Used whenever a history entry is edited or deleted,
// so correcting a mistake automatically fixes any totals that depended on it.
function recomputeDrum(fallback, historyForDrumAsc) {
  let status = fallback.status;
  let condition = fallback.condition;
  let party = fallback.party;
  let lastFixedDate = null;
  let totalLifeDays = 0;
  let repairCount = 0;
  let prevStatus = null;

  historyForDrumAsc.forEach((ev) => {
    const goingOnMachine = ev.status === "M/C";
    const wasOnMachine = prevStatus === "M/C";
    if (goingOnMachine) {
      lastFixedDate = ev.fixedDate || ev.date;
    } else if (wasOnMachine && lastFixedDate) {
      const cycleDays = diffDays(lastFixedDate, ev.date);
      if (cycleDays !== null && cycleDays >= 0) totalLifeDays += cycleDays;
      lastFixedDate = null;
    }
    const goingToParty = ev.status === "URWP" && prevStatus !== "URWP";
    if (goingToParty) repairCount += 1;
    status = ev.status;
    condition = ev.condition;
    party = ev.party;
    prevStatus = ev.status;
  });

  return { status, condition, party, lastFixedDate, totalLifeDays, repairCount };
}

// ---------- seed data, used only if storage is empty ----------
const SEED_DRUMS = [];

const SEED_HISTORY = SEED_DRUMS.map((d) => ({
  drumId: d.id,
  date: d.createdAt,
  status: d.status,
  condition: d.condition,
  party: d.party,
  notes: "Initial record",
  by: "System",
  ts: Date.parse(d.createdAt),
}));

// ---------- Employee roster & auth ----------
// Pre-load your team here (or add them later from the in-app Admin panel).
// Passwords are set by each employee on their first login — nothing is pre-set.
const SEED_ROSTER = [
  { id: "17107", name: "Giresh" },
];

// Fixed admin account — not part of the roster, not editable via the Admin
// panel. This is the one account that can create and manage other employees.
const ADMIN_ID = "ADMIN";
const ADMIN_PASSWORD = "admin@123";

// If someone starts editing a drum and closes the app/tab without saving,
// their lock shouldn't block everyone forever — treat a lock as expired
// after this long with no heartbeat.
const LOCK_TIMEOUT_MS = 2 * 60 * 1000;

async function hashPassword(pw) {
  const enc = new TextEncoder().encode(pw);
  const buf = await crypto.subtle.digest("SHA-256", enc);
  return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

export default function DrumTracker() {
  const isDesktop = useIsDesktop();
  const [drums, setDrums] = useState(null); // null = loading
  const [history, setHistory] = useState([]);
  const [roster, setRoster] = useState(null); // [{id, name}]
  const [credentials, setCredentials] = useState({}); // { [employeeId]: passwordHash }
  const [drumStyles, setDrumStyles] = useState(null); // ["20-58-15", ...]
  const [parLevels, setParLevels] = useState({}); // { [style]: minUsableCount }
  const [currentEmployee, setCurrentEmployee] = useState(null);
  const [query, setQuery] = useState("");
  const [typeFilter, setTypeFilter] = useState("ALL");
  const [classFilter, setClassFilter] = useState("ALL"); // classification bucket filter for Board/List
  const [drillDown, setDrillDown] = useState(null); // { style, key } — dashboard tile drill-down
  const [selectedDrum, setSelectedDrum] = useState(null); // drum id
  const [showAdd, setShowAdd] = useState(false);
  const [showLog, setShowLog] = useState(null); // drum id to log an update for
  const [editingEntry, setEditingEntry] = useState(null); // history event being edited
  const [locks, setLocks] = useState({}); // { [drumId]: { by, ts } } — who's currently editing what
  const [notes, setNotes] = useState([]); // [{ id, drumId, text, by, ts }] — free-form comments, separate from status history
  const [showSummary, setShowSummary] = useState(true);
  const [showAdmin, setShowAdmin] = useState(false);
  const [showMenu, setShowMenu] = useState(false);
  const [showChangePassword, setShowChangePassword] = useState(false);
  const [view, setView] = useState("board"); // "board" | "list" | "dashboard" | "machines"
  const [saving, setSaving] = useState(false);
  const [toast, setToast] = useState(null);

  const userName = currentEmployee?.name || "";

  async function fetchFresh(key, fallback) {
    try {
      const r = await window.storage.get(key, true);
      return r ? JSON.parse(r.value) : fallback;
    } catch (e) {
      return fallback;
    }
  }

  // ---------- load from shared persistent storage ----------
  useEffect(() => {
    (async () => {
      try {
        const d = await window.storage.get("drums", true).catch(() => null);
        const h = await window.storage.get("history", true).catch(() => null);
        const r = await window.storage.get("roster", true).catch(() => null);
        const c = await window.storage.get("credentials", true).catch(() => null);
        const ds = await window.storage.get("drumStyles", true).catch(() => null);
        const lockMap = await apiFetchLocks().catch(() => ({}));
        const pl = await window.storage.get("parLevels", true).catch(() => null);
        const nt = await window.storage.get("notes", true).catch(() => null);
        setDrums(d ? JSON.parse(d.value) : SEED_DRUMS);
        setHistory(h ? JSON.parse(h.value) : SEED_HISTORY);
        setRoster(r ? JSON.parse(r.value) : SEED_ROSTER);
        // The admin password lives in the same hashed-credentials store as
        // everyone else's, so it can be changed the same way — seed it from
        // the default on first run only, never overwrite an existing hash.
        const rawCredentials = c ? JSON.parse(c.value) : {};
        const credentialsNeededSeed = !rawCredentials[ADMIN_ID];
        const loadedCredentials = credentialsNeededSeed
          ? { ...rawCredentials, [ADMIN_ID]: await hashPassword(ADMIN_PASSWORD) }
          : rawCredentials;
        setCredentials(loadedCredentials);
        setDrumStyles(ds ? JSON.parse(ds.value) : DRUM_STYLES);
        setLocks(lockMap);
        setParLevels(pl ? JSON.parse(pl.value) : {});
        setNotes(nt ? JSON.parse(nt.value) : []);
        if (!d) await window.storage.set("drums", JSON.stringify(SEED_DRUMS), true);
        if (!h) await window.storage.set("history", JSON.stringify(SEED_HISTORY), true);
        if (!r) await window.storage.set("roster", JSON.stringify(SEED_ROSTER), true);
        if (credentialsNeededSeed) await window.storage.set("credentials", JSON.stringify(loadedCredentials), true);
        if (!ds) await window.storage.set("drumStyles", JSON.stringify(DRUM_STYLES), true);
        if (!pl) await window.storage.set("parLevels", JSON.stringify({}), true);
        if (!nt) await window.storage.set("notes", JSON.stringify([]), true);
      } catch (e) {
        setDrums(SEED_DRUMS);
        setHistory(SEED_HISTORY);
        setRoster(SEED_ROSTER);
        setCredentials({});
        setDrumStyles(DRUM_STYLES);
        setLocks({});
        setParLevels({});
        setNotes([]);
      }
    })();
  }, []);

  // ---------- live refresh: pick up other people's changes automatically ----------
  // Everyone shares one register, so without this, two people editing around the
  // same time would each be looking at a stale local copy — and whoever saves
  // last would silently wipe out the other's change. Polling keeps everyone's
  // view current, and (combined with the fresh-read-before-write pattern in
  // every save function below) narrows that risk down to a same-second collision.
  useEffect(() => {
    const interval = setInterval(async () => {
      if (drums === null) return; // don't poll until initial load finished
      const [freshDrums, freshHistory, freshRoster, freshStyles, freshLocks, freshPar, freshNotes] = await Promise.all([
        fetchFresh("drums", null),
        fetchFresh("history", null),
        fetchFresh("roster", null),
        fetchFresh("drumStyles", null),
        apiFetchLocks().catch(() => null),
        fetchFresh("parLevels", null),
        fetchFresh("notes", null),
      ]);
      if (freshDrums) setDrums(freshDrums);
      if (freshHistory) setHistory(freshHistory);
      if (freshRoster) setRoster(freshRoster);
      if (freshStyles) setDrumStyles(freshStyles);
      if (freshLocks) setLocks(freshLocks);
      if (freshPar) setParLevels(freshPar);
      if (freshNotes) setNotes(freshNotes);
    }, 10000);
    return () => clearInterval(interval);
  }, [drums === null]);

  // ---------- keep the editing lock alive while a modal is open ----------
  const activeLockDrumId = showLog || (editingEntry ? editingEntry.drumId : null);
  useEffect(() => {
    if (!activeLockDrumId) return;
    const heartbeat = setInterval(() => {
      acquireLock(activeLockDrumId); // refreshes our own lock's timestamp
    }, 30000);
    const releaseOnUnload = () => { releaseLock(activeLockDrumId); };
    window.addEventListener("beforeunload", releaseOnUnload);
    return () => {
      clearInterval(heartbeat);
      window.removeEventListener("beforeunload", releaseOnUnload);
    };
  }, [activeLockDrumId]);

  async function persist(nextDrums, nextHistory) {
    setSaving(true);
    try {
      if (nextDrums) await window.storage.set("drums", JSON.stringify(nextDrums), true);
      if (nextHistory) await window.storage.set("history", JSON.stringify(nextHistory), true);
    } catch (e) {
      showToast("Couldn't save — check connection.", true);
    } finally {
      setSaving(false);
    }
  }

  async function persistRoster(nextRoster) {
    setRoster(nextRoster);
    try {
      await window.storage.set("roster", JSON.stringify(nextRoster), true);
    } catch (e) {
      showToast("Couldn't save roster.", true);
    }
  }

  async function persistCredentials(nextCreds) {
    setCredentials(nextCreds);
    try {
      await window.storage.set("credentials", JSON.stringify(nextCreds), true);
    } catch (e) {
      showToast("Couldn't save.", true);
    }
  }

  // Self-service password change — works the same for the admin account and
  // every employee, since both live in the same hashed-credentials store.
  async function changeOwnPassword(currentPw, newPw) {
    const freshCredentials = await fetchFresh("credentials", credentials);
    const currentHash = await hashPassword(currentPw);
    if (currentHash !== freshCredentials[currentEmployee.id]) {
      return { ok: false, error: "Current password is incorrect." };
    }
    const newHash = await hashPassword(newPw);
    await persistCredentials({ ...freshCredentials, [currentEmployee.id]: newHash });
    return { ok: true };
  }

  async function persistDrumStyles(nextStyles) {
    setDrumStyles(nextStyles);
    try {
      await window.storage.set("drumStyles", JSON.stringify(nextStyles), true);
    } catch (e) {
      showToast("Couldn't save drum styles.", true);
    }
  }

  async function persistParLevels(nextLevels) {
    setParLevels(nextLevels);
    try {
      await window.storage.set("parLevels", JSON.stringify(nextLevels), true);
    } catch (e) {
      showToast("Couldn't save par levels.", true);
    }
  }

  async function addNote(drumId, text) {
    const freshNotes = await fetchFresh("notes", notes);
    const note = { id: `n-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`, drumId, text, by: userName || "Unnamed", ts: Date.now() };
    const nextNotes = [note, ...freshNotes];
    setNotes(nextNotes);
    try {
      await window.storage.set("notes", JSON.stringify(nextNotes), true);
    } catch (e) {
      showToast("Couldn't save note.", true);
    }
  }

  async function deleteNote(noteId) {
    const freshNotes = await fetchFresh("notes", notes);
    const nextNotes = freshNotes.filter((n) => n.id !== noteId);
    setNotes(nextNotes);
    try {
      await window.storage.set("notes", JSON.stringify(nextNotes), true);
    } catch (e) {
      showToast("Couldn't delete note.", true);
    }
  }

  function isLockActive(lock) {
    return lock && Date.now() - lock.ts < LOCK_TIMEOUT_MS;
  }

  // Tries to claim the editing lock for a drum. Returns the holder's name if
  // someone else already has an active lock on it, or null if the lock was
  // successfully claimed (or refreshed, if it's already ours). The acquire
  // itself is one atomic operation on the server (see server/index.js), so
  // two people clicking "edit" at the same instant can't both win it.
  async function acquireLock(drumId) {
    try {
      const result = await apiAcquireLock(drumId, userName);
      if (result.lock) {
        setLocks((prev) => ({ ...prev, [drumId]: result.lock }));
      }
      if (!result.ok) {
        return result.lock?.by ?? null;
      }
      return null;
    } catch (e) {
      // Couldn't reach the server — fail open rather than blocking the user
      // from editing; worst case the lock indicator lags briefly.
      return null;
    }
  }

  async function releaseLock(drumId) {
    setLocks((prev) => {
      const next = { ...prev };
      delete next[drumId];
      return next;
    });
    try {
      await apiReleaseLock(drumId, userName);
    } catch (e) {
      // non-fatal
    }
  }

  function showToast(msg, isError = false) {
    setToast({ msg, isError });
    setTimeout(() => setToast(null), 2600);
  }

  const filtered = useMemo(() => {
    if (!drums) return [];
    return drums.filter((d) => {
      const matchesQuery =
        !query ||
        d.id.toLowerCase().includes(query.toLowerCase()) ||
        d.size.toLowerCase().includes(query.toLowerCase()) ||
        (d.party || "").toLowerCase().includes(query.toLowerCase());
      const matchesType = typeFilter === "ALL" || d.size === typeFilter;
      const matchesClass = classFilter === "ALL" || classifyDrum(d) === classFilter;
      return matchesQuery && matchesType && matchesClass;
    });
  }, [drums, query, typeFilter, classFilter]);

  const board = useMemo(() => {
    const cols = Object.fromEntries(STATUSES.map((s) => [s.code, []]));
    filtered.forEach((d) => cols[d.status]?.push(d));
    return cols;
  }, [filtered]);

  const [sortKey, setSortKey] = useState("id");
  const [sortDir, setSortDir] = useState("asc");

  function toggleSort(key) {
    if (sortKey === key) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortKey(key);
      setSortDir("asc");
    }
  }

  const sortedList = useMemo(() => {
    const arr = [...filtered];
    arr.sort((a, b) => {
      let av, bv;
      if (sortKey === "id") { av = a.id; bv = b.id; }
      else if (sortKey === "style") { av = a.size; bv = b.size; }
      else if (sortKey === "status") { av = STATUS_MAP[a.status] || ""; bv = STATUS_MAP[b.status] || ""; }
      else if (sortKey === "life") { av = a.totalLifeDays || 0; bv = b.totalLifeDays || 0; }
      if (typeof av === "number" && typeof bv === "number") {
        return sortDir === "asc" ? av - bv : bv - av;
      }
      const cmp = String(av).localeCompare(String(bv), undefined, { numeric: true });
      return sortDir === "asc" ? cmp : -cmp;
    });
    return arr;
  }, [filtered, sortKey, sortDir]);

  function buildPartyCounts(styleDrums) {
    const counts = {};
    styleDrums.forEach((d) => {
      if (d.party && d.party.trim()) {
        counts[d.party] = (counts[d.party] || 0) + 1;
      }
    });
    return Object.entries(counts)
      .map(([party, count]) => ({ party, count }))
      .sort((a, b) => b.count - a.count);
  }

  function buildClassificationCounts(styleDrums) {
    const counts = Object.fromEntries(CLASSIFICATIONS.map((c) => [c.key, 0]));
    styleDrums.forEach((d) => {
      counts[classifyDrum(d)] += 1;
    });
    return counts;
  }

  function avgLife(styleDrums) {
    const withLife = styleDrums.filter((d) => (d.totalLifeDays || 0) > 0);
    if (withLife.length === 0) return 0;
    return Math.round(withLife.reduce((sum, d) => sum + (d.totalLifeDays || 0), 0) / withLife.length);
  }

  function totalRepairsOf(styleDrums) {
    return styleDrums.reduce((sum, d) => sum + (d.repairCount || 0), 0);
  }

  // How long each drum has sat in its current status — shown as a quiet
  // "Xd" tag on cards/rows, no threshold or alert, just information. Also
  // used below to help pick which available drum has been waiting longest.
  const daysInStatusMap = useMemo(() => {
    if (!drums || !history) return {};
    const map = {};
    drums.forEach((d) => {
      const drumHistory = history.filter((h) => h.drumId === d.id);
      if (drumHistory.length === 0) return;
      const latest = drumHistory.reduce((a, b) => (b.ts > a.ts ? b : a));
      const days = diffDays(latest.date, today());
      if (days !== null) map[d.id] = days;
    });
    return map;
  }, [drums, history]);

  const dashboardStats = useMemo(() => {
    if (!drums) return [];
    const stylesPresent = Array.from(new Set([...drumStyles, ...drums.map((d) => d.size)]));
    const rows = stylesPresent.map((style) => {
      const styleDrums = drums.filter((d) => d.size === style);
      const conditionCounts = Object.fromEntries(Object.keys(CONDITIONS).map((k) => [k, 0]));
      const statusCounts = Object.fromEntries(STATUSES.map((s) => [s.code, 0]));
      styleDrums.forEach((d) => {
        if (conditionCounts[d.condition] !== undefined) conditionCounts[d.condition] += 1;
        if (statusCounts[d.status] !== undefined) statusCounts[d.status] += 1;
      });
      return {
        style,
        total: styleDrums.length,
        conditionCounts,
        statusCounts,
        partyCounts: buildPartyCounts(styleDrums),
        avgLifeDays: avgLife(styleDrums),
        totalRepairs: totalRepairsOf(styleDrums),
        classificationCounts: buildClassificationCounts(styleDrums),
      };
    });
    rows.sort((a, b) => b.total - a.total || a.style.localeCompare(b.style, undefined, { numeric: true }));
    return rows.filter((r) => r.total > 0 || drumStyles.includes(r.style));
  }, [drums, drumStyles, daysInStatusMap]);

  const overallStats = useMemo(() => {
    if (!drums) return null;
    const conditionCounts = Object.fromEntries(Object.keys(CONDITIONS).map((k) => [k, 0]));
    const statusCounts = Object.fromEntries(STATUSES.map((s) => [s.code, 0]));
    drums.forEach((d) => {
      if (conditionCounts[d.condition] !== undefined) conditionCounts[d.condition] += 1;
      if (statusCounts[d.status] !== undefined) statusCounts[d.status] += 1;
    });
    return {
      total: drums.length,
      conditionCounts,
      statusCounts,
      partyCounts: buildPartyCounts(drums),
      avgLifeDays: avgLife(drums),
      totalRepairs: totalRepairsOf(drums),
      classificationCounts: buildClassificationCounts(drums),
    };
  }, [drums]);

  const machineView = useMemo(() => {
    if (!drums) return [];
    const onMachine = drums.filter((d) => d.status === "M/C" && d.machineNo && d.machineNo.trim());
    const groups = {};
    onMachine.forEach((d) => {
      const key = d.machineNo.trim();
      if (!groups[key]) groups[key] = [];
      groups[key].push(d);
    });
    // Sort strictly by the numeric value of the machine number (so "#5" comes
    // before "12" regardless of formatting), falling back to plain text
    // comparison only if a machine number has no digits at all.
    function machineNumValue(s) {
      const match = s.match(/\d+/);
      return match ? parseInt(match[0], 10) : null;
    }
    return Object.entries(groups)
      .map(([machineNo, machineDrums]) => ({ machineNo, drums: machineDrums }))
      .sort((a, b) => {
        const av = machineNumValue(a.machineNo);
        const bv = machineNumValue(b.machineNo);
        if (av !== null && bv !== null && av !== bv) return av - bv;
        if (av !== null && bv === null) return -1;
        if (av === null && bv !== null) return 1;
        return a.machineNo.localeCompare(b.machineNo, undefined, { numeric: true });
      });
  }, [drums]);

  const weeklyDigest = useMemo(() => {
    if (!drums || !history) return null;
    const isThisWeek = (dateStr) => {
      const d = diffDays(dateStr, today());
      return d !== null && d >= 0 && d <= 7;
    };
    const addedThisWeek = drums.filter((d) => d.createdAt && isThisWeek(d.createdAt)).length;
    const repairsThisWeek = history.filter((h) => h.repairNumber && isThisWeek(h.date)).length;
    const updatesThisWeek = history.filter((h) => h.notes !== "Drum added to register" && isThisWeek(h.date)).length;
    if (addedThisWeek === 0 && repairsThisWeek === 0 && updatesThisWeek === 0) return null;
    return { addedThisWeek, repairsThisWeek, updatesThisWeek };
  }, [drums, history]);

  const drillDownDrums = useMemo(() => {
    if (!drillDown || !drums) return [];
    return drums
      .filter((d) => (drillDown.style === "ALL" || d.size === drillDown.style) && classifyDrum(d) === drillDown.key)
      .sort((a, b) => (daysInStatusMap[b.id] || 0) - (daysInStatusMap[a.id] || 0));
  }, [drillDown, drums, daysInStatusMap]);

  function copySummaryText() {
    if (!overallStats) return;
    const scopeRow = typeFilter === "ALL" ? null : dashboardStats.find((r) => r.style === typeFilter);
    const stats = scopeRow || overallStats;
    const title = typeFilter === "ALL" ? "All styles" : typeFilter;
    const lines = [
      `*Drum Register Summary — ${title}*`,
      `Total: ${stats.total}`,
      "",
      ...CLASSIFICATIONS.filter((c) => c.key !== "OTHER" || (stats.classificationCounts[c.key] || 0) > 0).map(
        (c) => `${c.label}: ${stats.classificationCounts[c.key] || 0}`
      ),
    ];
    const text = lines.join("\n");
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(
        () => showToast("Summary copied — paste it into WhatsApp"),
        () => showToast("Couldn't copy — try again", true)
      );
    } else {
      showToast("Clipboard not available on this device", true);
    }
  }

  function exportDrumsExcel() {
    if (!drums) return;

    const registerHeaders = [
      "Drum ID", "Style", "Status Code", "Status", "Condition", "Machine No",
      "With Party", "Days In Status", "Total Life Days", "Times Repaired", "Added On",
    ];
    const registerRows = drums.map((d) => [
      d.id, d.size, d.status, STATUS_MAP[d.status] || "", CONDITIONS[d.condition]?.label || d.condition,
      d.machineNo || "", d.party || "", daysInStatusMap[d.id] ?? "", d.totalLifeDays || 0, d.repairCount || 0, d.createdAt || "",
    ]);
    const registerSheet = XLSX.utils.aoa_to_sheet([registerHeaders, ...registerRows]);
    registerSheet["!cols"] = registerHeaders.map((h) => ({ wch: Math.max(h.length + 2, 14) }));

    const historyHeaders = ["Drum ID", "Date", "Status", "Condition", "Party", "Days in Cycle", "Notes", "Logged By"];
    const historySorted = [...history].sort((a, b) => (a.drumId === b.drumId ? a.ts - b.ts : a.drumId.localeCompare(b.drumId)));
    const historyRows = historySorted.map((h) => [
      h.drumId, h.date, STATUS_MAP[h.status] || h.status, CONDITIONS[h.condition]?.label || h.condition,
      h.party || "", h.cycleDays ?? "", h.notes || "", h.by || "",
    ]);
    const historySheet = XLSX.utils.aoa_to_sheet([historyHeaders, ...historyRows]);
    historySheet["!cols"] = historyHeaders.map((h) => ({ wch: Math.max(h.length + 2, 14) }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, registerSheet, "Drum Register");
    XLSX.utils.book_append_sheet(wb, historySheet, "Life History");
    XLSX.writeFile(wb, `drum-register-${today()}.xlsx`);
    showToast("Excel file downloaded");
  }

  async function exportDrumsPDF() {
    if (!drums || !overallStats) return;

    const [{ default: jsPDF }, { default: autoTable }] = await Promise.all([
      import("jspdf"),
      import("jspdf-autotable"),
    ]);

    const doc = new jsPDF({ orientation: "landscape", unit: "pt" });
    const amber = [200, 16, 46];
    const ink = [26, 32, 41];
    const marginX = 40;

    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.setTextColor(...ink);
    doc.text("MRF Tyres — Goa Plant", marginX, 44);
    doc.setFontSize(11);
    doc.setFont("helvetica", "normal");
    doc.text("Tyre Building Drum Register — Summary Report", marginX, 62);
    doc.setFontSize(9);
    doc.setTextColor(107, 117, 128);
    doc.text(`Generated ${fmtDate(today())}`, marginX, 76);

    autoTable(doc, {
      startY: 92,
      margin: { left: marginX },
      tableWidth: 250,
      head: [["Overall", ""]],
      body: [
        ["Total drums", String(overallStats.total)],
        ["Average life", `${overallStats.avgLifeDays} days`],
        ["Total repairs", String(overallStats.totalRepairs)],
      ],
      theme: "grid",
      headStyles: { fillColor: ink, fontSize: 9 },
      styles: { fontSize: 9, cellPadding: 5 },
      columnStyles: { 1: { halign: "right" } },
    });

    const classRows = CLASSIFICATIONS.filter(
      (c) => c.key !== "OTHER" || (overallStats.classificationCounts[c.key] || 0) > 0
    ).map((c) => [c.label, String(overallStats.classificationCounts[c.key] || 0)]);

    autoTable(doc, {
      startY: 92,
      margin: { left: marginX + 270 },
      tableWidth: 250,
      head: [["Classification", "Count"]],
      body: classRows,
      theme: "grid",
      headStyles: { fillColor: amber, fontSize: 9 },
      styles: { fontSize: 9, cellPadding: 5 },
      columnStyles: { 1: { halign: "right" } },
    });

    const byStyleRows = dashboardStats
      .filter((r) => r.total > 0)
      .map((r) => [r.style, String(r.total), `${r.avgLifeDays}d`, String(r.totalRepairs)]);

    autoTable(doc, {
      startY: doc.lastAutoTable.finalY + 20,
      margin: { left: marginX },
      head: [["Style", "Drums", "Avg. Life", "Repairs"]],
      body: byStyleRows,
      theme: "striped",
      headStyles: { fillColor: ink, fontSize: 9 },
      styles: { fontSize: 9, cellPadding: 5 },
    });

    doc.addPage();
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.setTextColor(...ink);
    doc.text("Full Drum Register", marginX, 40);

    autoTable(doc, {
      startY: 54,
      margin: { left: marginX, right: marginX },
      head: [[
        "Drum ID", "Style", "Status", "Condition", "Machine", "With Party",
        "Days in Status", "Life Days", "Repairs", "Added On",
      ]],
      body: drums.map((d) => [
        d.id,
        d.size,
        STATUS_MAP[d.status] || "",
        CONDITIONS[d.condition]?.label || d.condition,
        d.machineNo || "—",
        d.party || "—",
        daysInStatusMap[d.id] ?? "—",
        d.totalLifeDays || 0,
        d.repairCount || 0,
        d.createdAt ? fmtDate(d.createdAt) : "—",
      ]),
      theme: "striped",
      headStyles: { fillColor: amber, fontSize: 8.5 },
      styles: { fontSize: 8, cellPadding: 4 },
      didDrawPage: () => {
        const pageCount = doc.internal.getNumberOfPages();
        doc.setFontSize(8);
        doc.setTextColor(150, 150, 150);
        doc.text(
          `Page ${doc.internal.getCurrentPageInfo().pageNumber} of ${pageCount}`,
          doc.internal.pageSize.getWidth() - marginX - 60,
          doc.internal.pageSize.getHeight() - 16
        );
      },
    });

    doc.save(`drum-register-${today()}.pdf`);
    showToast("PDF report downloaded");
  }

  async function addDrum(newDrum) {
    const freshDrums = await fetchFresh("drums", drums);
    const freshHistory = await fetchFresh("history", history);
    const record = {
      ...newDrum,
      machineNo: newDrum.status === "M/C" ? newDrum.machineNo : "",
      prNumber: newDrum.status === "URWP" ? newDrum.prNumber : "",
      poNumber: newDrum.status === "URWP" ? newDrum.poNumber : "",
      scrapReason: newDrum.status === "SCRAP" ? newDrum.scrapReason : "",
      createdAt: today(),
      lastFixedDate: newDrum.status === "M/C" ? today() : null,
      totalLifeDays: 0,
      // A repair is counted the moment a drum goes out to a party for
      // repair, not later when it comes back marked Repaired — that's the
      // event the shop actually wants reflected in "times repaired".
      repairCount: newDrum.status === "URWP" ? 1 : 0,
    };
    if (freshDrums.some((d) => d.id === record.id)) {
      showToast(`${record.id} already exists — someone may have just added it.`, true);
      return;
    }
    const nextDrums = [record, ...freshDrums];
    const event = {
      drumId: record.id,
      date: today(),
      status: record.status,
      condition: record.condition,
      party: record.party,
      machineNo: record.status === "M/C" ? record.machineNo : undefined,
      notes: "Drum added to register",
      by: userName || "Unnamed",
      ts: Date.now(),
    };
    const nextHistory = [event, ...freshHistory];
    setDrums(nextDrums);
    setHistory(nextHistory);
    persist(nextDrums, nextHistory);
    setShowAdd(false);
    showToast(`${record.id} added`);
  }

  async function logUpdate(drumId, update) {
    const freshDrums = await fetchFresh("drums", drums);
    const freshHistory = await fetchFresh("history", history);
    const drum = freshDrums.find((d) => d.id === drumId);
    if (!drum) {
      showToast("This drum no longer exists — it may have been deleted by someone else.", true);
      setShowLog(null);
      return;
    }
    const wasOnMachine = drum.status === "M/C";
    const goingOnMachine = update.status === "M/C";
    let lastFixedDate = drum.lastFixedDate;
    let totalLifeDays = drum.totalLifeDays || 0;
    let cycleDays = null;

    if (goingOnMachine) {
      // Starting a new cycle on the machine
      lastFixedDate = update.fixedDate || update.date || today();
    } else if (wasOnMachine && drum.lastFixedDate) {
      // Coming off the machine — close out this cycle's life
      cycleDays = diffDays(drum.lastFixedDate, update.date || today());
      if (cycleDays !== null && cycleDays >= 0) {
        totalLifeDays += cycleDays;
      }
      lastFixedDate = null;
    }

    // A repair is counted the moment a drum goes out to a party for repair
    // (status URWP), not later when it comes back marked Repaired — and
    // only on the transition into URWP, so re-saving while already URWP
    // doesn't count it twice.
    const justSentToParty = update.status === "URWP" && drum.status !== "URWP";
    const repairCount = (drum.repairCount || 0) + (justSentToParty ? 1 : 0);

    const nextDrums = freshDrums.map((d) =>
      d.id === drumId
        ? {
            ...d,
            status: update.status,
            condition: update.condition,
            party: update.party,
            machineNo: update.status === "M/C" ? update.machineNo : "",
            prNumber: update.status === "URWP" ? update.prNumber : "",
            poNumber: update.status === "URWP" ? update.poNumber : "",
            scrapReason: update.status === "SCRAP" ? update.scrapReason : "",
            lastFixedDate,
            totalLifeDays,
            repairCount,
          }
        : d
    );
    const event = {
      drumId,
      date: update.date || today(),
      status: update.status,
      condition: update.condition,
      party: update.party,
      machineNo: update.status === "M/C" ? update.machineNo : undefined,
      prNumber: update.status === "URWP" ? update.prNumber : undefined,
      poNumber: update.status === "URWP" ? update.poNumber : undefined,
      scrapReason: update.status === "SCRAP" ? update.scrapReason : undefined,
      notes: update.notes,
      by: userName || "Unnamed",
      cycleDays,
      fixedDate: goingOnMachine ? lastFixedDate : (wasOnMachine ? drum.lastFixedDate : undefined),
      repairNumber: justSentToParty ? repairCount : undefined,
      ts: Date.now(),
    };
    const nextHistory = [event, ...freshHistory];
    setDrums(nextDrums);
    setHistory(nextHistory);
    persist(nextDrums, nextHistory);
    setShowLog(null);
    showToast(cycleDays !== null ? `${drumId} updated — ${cycleDays} day(s) in service this cycle` : `${drumId} updated`);
  }

  async function deleteDrum(drumId) {
    const freshDrums = await fetchFresh("drums", drums);
    const freshHistory = await fetchFresh("history", history);
    const freshNotes = await fetchFresh("notes", notes);
    const nextDrums = freshDrums.filter((d) => d.id !== drumId);
    const nextHistory = freshHistory.filter((h) => h.drumId !== drumId);
    const nextNotes = freshNotes.filter((n) => n.drumId !== drumId);
    setDrums(nextDrums);
    setHistory(nextHistory);
    setNotes(nextNotes);
    persist(nextDrums, nextHistory);
    try {
      await window.storage.set("notes", JSON.stringify(nextNotes), true);
    } catch (e) {
      // non-fatal
    }
    setSelectedDrum(null);
    showToast(`${drumId} deleted`);
  }

  async function deleteHistoryEntry(drumId, ts) {
    const freshDrums = await fetchFresh("drums", drums);
    const freshHistory = await fetchFresh("history", history);
    const nextHistory = freshHistory.filter((h) => !(h.drumId === drumId && h.ts === ts));
    const ascForDrum = nextHistory.filter((h) => h.drumId === drumId).sort((a, b) => a.ts - b.ts);
    const baseDrum = freshDrums.find((d) => d.id === drumId);
    if (!baseDrum) return;
    const recomputed = recomputeDrum(baseDrum, ascForDrum);
    const nextDrums = freshDrums.map((d) => (d.id === drumId ? { ...d, ...recomputed } : d));
    setDrums(nextDrums);
    setHistory(nextHistory);
    persist(nextDrums, nextHistory);
    showToast("Entry deleted — totals recalculated");
  }

  async function editHistoryEntry(drumId, ts, updates) {
    const freshDrums = await fetchFresh("drums", drums);
    const freshHistory = await fetchFresh("history", history);
    const nextHistory = freshHistory.map((h) =>
      h.drumId === drumId && h.ts === ts ? { ...h, ...updates } : h
    );
    const ascForDrum = nextHistory.filter((h) => h.drumId === drumId).sort((a, b) => a.ts - b.ts);
    const baseDrum = freshDrums.find((d) => d.id === drumId);
    if (!baseDrum) { setEditingEntry(null); return; }
    const recomputed = recomputeDrum(baseDrum, ascForDrum);
    const nextDrums = freshDrums.map((d) => (d.id === drumId ? { ...d, ...recomputed } : d));
    setDrums(nextDrums);
    setHistory(nextHistory);
    persist(nextDrums, nextHistory);
    setEditingEntry(null);
    showToast("Entry updated — totals recalculated");
  }

  if (drums === null || roster === null || drumStyles === null) {
    return (
      <div style={{ ...styles.app, alignItems: "center", justifyContent: "center", display: "flex" }}>
        <Loader2 className="spin" size={28} color="#8A9199" />
        <style>{`.spin{animation:spin 0.9s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      </div>
    );
  }

  if (!currentEmployee) {
    return (
      <AuthGate
        roster={roster}
        credentials={credentials}
        onLogin={(emp) => setCurrentEmployee(emp)}
      />
    );
  }

  if (showAdmin) {
    return (
      <AdminPanel
        roster={roster}
        credentials={credentials}
        drumStyles={drumStyles}
        onClose={() => setShowAdmin(false)}
        onAddEmployee={async (emp) => {
          const hash = await hashPassword(emp.password);
          await persistCredentials({ ...credentials, [emp.id]: hash });
          persistRoster([...roster, { id: emp.id, name: emp.name }]);
        }}
        onRemoveEmployee={(id) => {
          persistRoster(roster.filter((r) => r.id !== id));
          const nc = { ...credentials };
          delete nc[id];
          persistCredentials(nc);
        }}
        onResetPassword={async (id, newPassword) => {
          const hash = await hashPassword(newPassword);
          await persistCredentials({ ...credentials, [id]: hash });
        }}
        onAddStyle={(style) => persistDrumStyles([...drumStyles, style])}
        onRemoveStyle={(style) => persistDrumStyles(drumStyles.filter((s) => s !== style))}
        parLevels={parLevels}
        onSetParLevel={(style, value) => persistParLevels({ ...parLevels, [style]: value })}
      />
    );
  }

  const selected = selectedDrum ? drums.find((d) => d.id === selectedDrum) : null;
  const logTarget = showLog ? drums.find((d) => d.id === showLog) : null;

  return (
    <div style={styles.app}>
      <style>{globalCss}</style>
      <div style={isDesktop ? styles.desktopContainer : undefined}>

      <header style={styles.header}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button style={styles.iconBtn} onClick={() => setShowMenu(true)} title="Menu">
            <Menu size={20} color="#6B7580" />
          </button>
          <div>
            <div style={styles.eyebrow}>MRF TYRES · GOA PLANT</div>
            <h1 style={styles.h1}>Drum Registry</h1>
          </div>
        </div>
        <div style={styles.headerRight}>
          <div style={styles.userBadge}>
            <User size={14} />
            {userName}
          </div>
          <button
            type="button"
            style={styles.iconBtn}
            title="Log out"
            onClick={() => { setCurrentEmployee(null); setShowAdmin(false); setShowMenu(false); }}
          >
            <LogOut size={16} color="#6B7580" />
          </button>
          <button style={styles.primaryBtn} onClick={() => setShowAdd(true)}>
            <Plus size={16} /> Add drum
          </button>
        </div>
      </header>

      {showMenu && (
        <SideMenu
          view={view}
          isAdmin={currentEmployee?.id === ADMIN_ID}
          onSelectBoard={() => { setView("board"); setShowMenu(false); }}
          onSelectList={() => { setView("list"); setShowMenu(false); }}
          onSelectDashboard={() => { setView("dashboard"); setShowMenu(false); }}
          onSelectMachines={() => { setView("machines"); setShowMenu(false); }}
          onSelectAdmin={() => { setShowAdmin(true); setShowMenu(false); }}
          onExportCSV={() => { exportDrumsExcel(); setShowMenu(false); }}
          onExportPDF={() => { exportDrumsPDF(); setShowMenu(false); }}
          onChangePassword={() => { setShowChangePassword(true); setShowMenu(false); }}
          onClose={() => setShowMenu(false)}
        />
      )}

      {showSummary && weeklyDigest && (
        <div style={styles.summaryBanner}>
          <div style={styles.summaryBannerText}>
            This week: <strong>{weeklyDigest.addedThisWeek}</strong> drum{weeklyDigest.addedThisWeek === 1 ? "" : "s"} added ·{" "}
            <strong>{weeklyDigest.repairsThisWeek}</strong> repair{weeklyDigest.repairsThisWeek === 1 ? "" : "s"} logged ·{" "}
            <strong>{weeklyDigest.updatesThisWeek}</strong> update{weeklyDigest.updatesThisWeek === 1 ? "" : "s"} overall
          </div>
          <button type="button" style={styles.summaryBannerClose} onClick={() => setShowSummary(false)}>
            <X size={14} />
          </button>
        </div>
      )}

      {view !== "dashboard" && view !== "machines" && (
        <div style={styles.toolbar}>
          <div style={styles.searchBox}>
            <Search size={15} color="#8A9199" />
            <input
              placeholder="Search by drum ID, size, or party..."
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              style={styles.searchInput}
            />
          </div>
          <div style={{ width: 160 }}>
            <CustomSelect
              value={typeFilter}
              onChange={setTypeFilter}
              options={[{ value: "ALL", label: "All styles" }, ...drumStyles.map((t) => ({ value: t, label: t }))]}
            />
          </div>
          <div style={{ width: 190 }}>
            <CustomSelect
              value={classFilter}
              onChange={setClassFilter}
              options={[{ value: "ALL", label: "All classifications" }, ...CLASSIFICATIONS.map((c) => ({ value: c.key, label: c.label }))]}
            />
          </div>
          <div style={styles.countPill}>{filtered.length} drum{filtered.length !== 1 ? "s" : ""}</div>
        </div>
      )}

      {view === "dashboard" && (
        <div style={styles.toolbar}>
          <div style={{ width: 160 }}>
            <CustomSelect
              value={typeFilter}
              onChange={setTypeFilter}
              options={[{ value: "ALL", label: "All styles" }, ...drumStyles.map((t) => ({ value: t, label: t }))]}
            />
          </div>
          {typeFilter !== "ALL" && (
            <button type="button" style={styles.clearFilterBtn} onClick={() => setTypeFilter("ALL")}>
              <X size={13} /> Clear
            </button>
          )}
          <button type="button" style={styles.clearFilterBtn} onClick={copySummaryText}>
            <ClipboardCopy size={13} /> Copy summary
          </button>
        </div>
      )}

      {view === "board" ? (
        <div style={styles.board}>
          {STATUSES.map((s) => {
            const colDrums = board[s.code];
            const okCount = colDrums.filter((d) => d.condition === "OKAY").length;
            const notOkCount = colDrums.filter((d) => d.condition === "NOT OKAY").length;
            return (
              <div key={s.code} style={styles.column}>
                <div style={styles.columnHead}>
                  <span style={styles.columnCode}>{s.code}</span>
                  <span style={styles.columnLabel}>{s.label}</span>
                  <span style={styles.columnCount}>{colDrums.length}</span>
                  {(okCount > 0 || notOkCount > 0) && (
                    <div style={styles.columnBreakdown}>
                      {okCount > 0 && (
                        <span style={{ ...styles.columnBreakdownChip, color: CONDITIONS.OKAY.fg, background: CONDITIONS.OKAY.bg }}>
                          {okCount} Okay
                        </span>
                      )}
                      {notOkCount > 0 && (
                        <span style={{ ...styles.columnBreakdownChip, color: CONDITIONS["NOT OKAY"].fg, background: CONDITIONS["NOT OKAY"].bg }}>
                          {notOkCount} Not Okay
                        </span>
                      )}
                    </div>
                  )}
                </div>
                <div style={styles.columnBody}>
                  {colDrums.length === 0 && <div style={styles.emptyCol}>No drums here</div>}
                  {colDrums.map((d) => {
                    const lock = locks[d.id];
                    const lockedBy = isLockActive(lock) && lock.by !== userName ? lock.by : null;
                    return (
                      <DrumTag
                        key={d.id}
                        drum={d}
                        onClick={() => setSelectedDrum(d.id)}
                        lockedBy={lockedBy}
                        daysInStatus={daysInStatusMap[d.id]}
                      />
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      ) : view === "list" ? (
        <div style={styles.listView}>
          <div style={styles.sortChipRow}>
            <SortChip label="ID" active={sortKey === "id"} dir={sortDir} onClick={() => toggleSort("id")} />
            <SortChip label="Style" active={sortKey === "style"} dir={sortDir} onClick={() => toggleSort("style")} />
            <SortChip label="Status" active={sortKey === "status"} dir={sortDir} onClick={() => toggleSort("status")} />
            <SortChip label="Life" active={sortKey === "life"} dir={sortDir} onClick={() => toggleSort("life")} />
          </div>
          {sortedList.length === 0 && <div style={styles.emptyCol}>No drums match.</div>}
          {sortedList.map((d) => {
            const cond = CONDITIONS[d.condition] || CONDITIONS.NEW;
            const lock = locks[d.id];
            const lockedBy = isLockActive(lock) && lock.by !== userName ? lock.by : null;
            const daysInStatus = daysInStatusMap[d.id];
            return (
              <button key={d.id} className="dt-hover" style={styles.listCard} onClick={() => setSelectedDrum(d.id)}>
                <div style={styles.listCardTop}>
                  <span style={styles.listCardId}>{d.id}</span>
                  <span style={{ ...styles.tagCondition, color: cond.fg, background: cond.bg }}>{cond.label}</span>
                </div>
                <div style={styles.listCardMeta}>
                  {d.size} · {STATUS_MAP[d.status]}
                  {daysInStatus !== undefined && <span style={styles.listCardDays}> · {daysInStatus}d here</span>}
                </div>
                <div style={styles.listCardFooter}>
                  <span style={styles.listCardLife}>{d.totalLifeDays || 0}d life · {d.repairCount || 0} repair{d.repairCount === 1 ? "" : "s"}</span>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    {lockedBy && (
                      <div style={styles.listRowLock} title={`Being edited by ${lockedBy}`}><Lock size={11} /></div>
                    )}
                    <ChevronRight size={16} color="#9AA2AA" />
                  </div>
                </div>
              </button>
            );
          })}
        </div>
      ) : view === "dashboard" ? (
        <Dashboard
          overall={overallStats}
          rows={dashboardStats.filter((r) => typeFilter === "ALL" || r.style === typeFilter)}
          showOverall={typeFilter === "ALL"}
          parLevels={parLevels}
          onTileClick={(style, key) => setDrillDown({ style, key })}
        />
      ) : (
        <MachineView machines={machineView} onSelectDrum={(id) => setSelectedDrum(id)} />
      )}

      {drillDown && (
        <DrillDownModal
          drillDown={drillDown}
          drums={drillDownDrums}
          daysInStatusMap={daysInStatusMap}
          onClose={() => setDrillDown(null)}
          onSelectDrum={(id) => { setDrillDown(null); setSelectedDrum(id); }}
        />
      )}

      {saving && <div style={styles.savingIndicator}>Saving…</div>}
      {toast && (
        <div style={{ ...styles.toast, ...(toast.isError ? styles.toastError : {}) }}>
          {toast.isError ? <AlertTriangle size={15} /> : <CheckCircle2 size={15} />}
          {toast.msg}
        </div>
      )}

      {selected && (
        <DrumDetail
          drum={selected}
          history={history.filter((h) => h.drumId === selected.id).sort((a, b) => b.ts - a.ts)}
          activeLock={isLockActive(locks[selected.id]) ? locks[selected.id] : null}
          currentUser={userName}
          onClose={() => setSelectedDrum(null)}
          onLog={async () => {
            const heldBy = await acquireLock(selected.id);
            if (heldBy) {
              showToast(`${heldBy} is currently editing this drum — try again shortly.`, true);
              return;
            }
            setShowLog(selected.id);
            setSelectedDrum(null);
          }}
          onDelete={async () => {
            const heldBy = await acquireLock(selected.id);
            if (heldBy) {
              showToast(`${heldBy} is currently editing this drum — try again shortly.`, true);
              return;
            }
            deleteDrum(selected.id);
          }}
          onEditEntry={async (h) => {
            const heldBy = await acquireLock(h.drumId);
            if (heldBy) {
              showToast(`${heldBy} is currently editing this drum — try again shortly.`, true);
              return;
            }
            setEditingEntry(h);
          }}
          onDeleteEntry={async (h) => {
            const heldBy = await acquireLock(h.drumId);
            if (heldBy) {
              showToast(`${heldBy} is currently editing this drum — try again shortly.`, true);
              return;
            }
            await deleteHistoryEntry(h.drumId, h.ts);
            releaseLock(h.drumId);
          }}
          notes={notes.filter((n) => n.drumId === selected.id).sort((a, b) => b.ts - a.ts)}
          onAddNote={(text) => addNote(selected.id, text)}
          onDeleteNote={(noteId) => deleteNote(noteId)}
        />
      )}

      {showAdd && <AddDrumModal onClose={() => setShowAdd(false)} onSave={addDrum} existingIds={drums.map((d) => d.id)} drumStyles={drumStyles} allDrums={drums} />}

      {showChangePassword && (
        <ChangePasswordModal
          onClose={() => setShowChangePassword(false)}
          onSave={changeOwnPassword}
        />
      )}

      {logTarget && (
        <LogUpdateModal
          drum={logTarget}
          allDrums={drums}
          onClose={() => { setShowLog(null); releaseLock(logTarget.id); }}
          onSave={async (u) => { await logUpdate(logTarget.id, u); releaseLock(logTarget.id); }}
        />
      )}

      {editingEntry && (
        <EditHistoryModal
          entry={editingEntry}
          onClose={() => { const id = editingEntry.drumId; setEditingEntry(null); releaseLock(id); }}
          onSave={async (updates) => { await editHistoryEntry(editingEntry.drumId, editingEntry.ts, updates); releaseLock(editingEntry.drumId); }}
        />
      )}
      </div>
      <Footer />
    </div>
  );
}

// Steps: "id" -> enter employee ID
//        "setpw" -> ID recognized but has no password yet, set one now
//        "password" -> ID has a password, enter it
//        "forgot" -> shows instructions + a way into the admin panel to reset
function SideMenu({ view, isAdmin, onSelectBoard, onSelectList, onSelectDashboard, onSelectMachines, onSelectAdmin, onExportCSV, onExportPDF, onChangePassword, onClose }) {
  return (
    <div style={styles.menuOverlay} onClick={onClose}>
      <div style={styles.menuDrawer} onClick={(e) => e.stopPropagation()}>
        <div style={styles.panelHead}>
          <div style={styles.eyebrow}>MENU</div>
          <button style={styles.iconBtn} onClick={onClose}><X size={18} /></button>
        </div>
        <nav style={styles.menuNav}>
          <button
            style={{ ...styles.menuItem, ...(view === "board" ? styles.menuItemActive : {}) }}
            onClick={onSelectBoard}
          >
            <LayoutGrid size={17} /> Board view
          </button>
          <button
            style={{ ...styles.menuItem, ...(view === "list" ? styles.menuItemActive : {}) }}
            onClick={onSelectList}
          >
            <ListIcon size={17} /> All drums
          </button>
          <button
            style={{ ...styles.menuItem, ...(view === "dashboard" ? styles.menuItemActive : {}) }}
            onClick={onSelectDashboard}
          >
            <BarChart3 size={17} /> Dashboard
          </button>
          <button
            style={{ ...styles.menuItem, ...(view === "machines" ? styles.menuItemActive : {}) }}
            onClick={onSelectMachines}
          >
            <Cog size={17} /> Machine view
          </button>
          {isAdmin && (
            <button style={styles.menuItem} onClick={onSelectAdmin}>
              <ShieldCheck size={17} /> Admin panel
            </button>
          )}
          <div style={styles.menuDivider} />
          <button style={styles.menuItem} onClick={onExportCSV}>
            <Download size={17} /> Export Excel
          </button>
          <button style={styles.menuItem} onClick={onExportPDF}>
            <FileText size={17} /> Export PDF Report
          </button>
          <div style={styles.menuDivider} />
          <button style={styles.menuItem} onClick={onChangePassword}>
            <KeyRound size={17} /> Change Password
          </button>
        </nav>
      </div>
    </div>
  );
}

// Decorative tire cross-section — pure SVG, no external asset, no trademarked
// artwork. Concentric rings + tread ticks evoke the product without
// reproducing any actual MRF logo or mark.
// Large, faint brand mark for the login backdrop. A radial mask feathers
// the image's hard rectangular edge into nothing, so it reads as a soft
// glow bled into the background rather than a sticker pasted on top.
function LogoWatermark({ size = 640, opacity = 0.16, style }) {
  return (
    <img
      src="/brand/mrf-logo.png"
      alt=""
      aria-hidden="true"
      style={{
        position: "absolute",
        width: size,
        height: size,
        opacity,
        pointerEvents: "none",
        userSelect: "none",
        WebkitMaskImage: "radial-gradient(circle at center, #000 38%, transparent 68%)",
        maskImage: "radial-gradient(circle at center, #000 38%, transparent 68%)",
        ...style,
      }}
    />
  );
}

function AuthGate({ roster, credentials, onLogin }) {
  const [empId, setEmpId] = useState("");
  const [pw, setPw] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [showForgot, setShowForgot] = useState(false);

  async function submit() {
    setError("");
    setShowForgot(false);
    if (!empId.trim() || !pw) {
      setError("Enter your employee ID and password.");
      return;
    }
    const isAdminId = empId.trim().toUpperCase() === ADMIN_ID;
    const matched = isAdminId
      ? { id: ADMIN_ID, name: "Administrator" }
      : roster.find((r) => r.id.toLowerCase() === empId.trim().toLowerCase());
    if (!matched) {
      setError("That employee ID isn't on the roster. Ask your admin to add you.");
      return;
    }
    setBusy(true);
    const hash = await hashPassword(pw);
    setBusy(false);
    // Admin's password lives in the same hashed store as everyone else's —
    // the plaintext default is only a fallback for the rare case where the
    // seed-on-load step hasn't run yet.
    const matchesStoredHash = hash === credentials[matched.id];
    const matchesAdminDefault = isAdminId && !credentials[matched.id] && pw === ADMIN_PASSWORD;
    if (matchesStoredHash || matchesAdminDefault) {
      onLogin(matched);
    } else {
      setError("Wrong employee ID or password.");
    }
  }

  return (
    <div style={styles.authScreen}>
      <style>{globalCss}</style>
      <LogoWatermark
        size={1100}
        opacity={0.14}
        style={{ top: "50%", left: "100%", transform: "translate(-46%, -50%)" }}
      />

      <div style={styles.authPanel}>
        <div style={styles.authBrandRow}>
          <img src="/brand/mrf-logo.png" alt="MRF Tyres" style={styles.authBadge} />
          <div>
            <div style={styles.authEyebrow}>MRF Tyres · Goa Plant</div>
            <div style={styles.authTitle}>Drum Registry</div>
          </div>
        </div>

        <p style={styles.authTagline}>
          Sign in to track tyre building drums through the shop floor — status, repairs, and life history in one place.
        </p>

        <div style={styles.authCard}>
          <div style={styles.field}>
            <label style={styles.fieldLabel}>Employee ID</label>
            <input
              autoFocus
              placeholder="e.g. 17107"
              value={empId}
              onChange={(e) => setEmpId(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
              style={styles.input}
            />
          </div>
          <div style={styles.field}>
            <label style={styles.fieldLabel}>Password</label>
            <PasswordInput
              placeholder="Password"
              value={pw}
              onChange={(e) => setPw(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
            />
          </div>

          {error && <div style={{ ...styles.formError, marginBottom: 14 }}><AlertTriangle size={14} /> {error}</div>}

          <button
            type="button"
            disabled={busy}
            onClick={submit}
            style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginBottom: 10 }}
          >
            {busy ? "Checking…" : "Sign In"}
          </button>

          <button type="button" onClick={() => setShowForgot((v) => !v)} style={styles.linkBtn}>
            Forgot password?
          </button>
          {showForgot && (
            <p style={styles.authForgotNote}>
              Ask your admin to sign in and set a new password for you from the Admin panel.
            </p>
          )}
        </div>
      </div>

      <Footer light />
    </div>
  );
}

function AdminPanel({ roster, credentials, drumStyles, onClose, onAddEmployee, onRemoveEmployee, onResetPassword, onAddStyle, onRemoveStyle, parLevels, onSetParLevel }) {
  const [newId, setNewId] = useState("");
  const [newName, setNewName] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [newStyle, setNewStyle] = useState("");
  const [error, setError] = useState("");
  const [styleError, setStyleError] = useState("");
  const [parDrafts, setParDrafts] = useState({});
  const [resetTargetId, setResetTargetId] = useState(null);
  const [resetDraft, setResetDraft] = useState("");
  const [resetError, setResetError] = useState("");

  function submitReset(id) {
    setResetError("");
    if (resetDraft.trim().length < 4) {
      setResetError("Password should be at least 4 characters.");
      return;
    }
    onResetPassword(id, resetDraft.trim());
    setResetTargetId(null);
    setResetDraft("");
  }

  function addEmployee() {
    setError("");
    if (!newId.trim() || !newName.trim() || !newPassword.trim()) {
      setError("Enter employee ID, name, and a password.");
      return;
    }
    if (newId.trim().toUpperCase() === ADMIN_ID) {
      setError(`"${ADMIN_ID}" is reserved for the admin account.`);
      return;
    }
    if (roster.some((r) => r.id.toLowerCase() === newId.trim().toLowerCase())) {
      setError("That employee ID already exists."); return;
    }
    if (newPassword.trim().length < 4) {
      setError("Password should be at least 4 characters."); return;
    }
    onAddEmployee({ id: newId.trim(), name: newName.trim(), password: newPassword.trim() });
    setNewId(""); setNewName(""); setNewPassword("");
  }

  function addStyle() {
    setStyleError("");
    if (!newStyle.trim()) { setStyleError("Enter a drum style code."); return; }
    if (drumStyles.some((s) => s.toLowerCase() === newStyle.trim().toLowerCase())) {
      setStyleError("That style already exists."); return;
    }
    onAddStyle(newStyle.trim());
    setNewStyle("");
  }

  return (
    <div style={styles.app}>
      <style>{globalCss}</style>
      <div style={styles.panelHead}>
        <div>
          <button type="button" onClick={onClose} style={styles.backLink}><ArrowLeft size={13} /> Back to register</button>
          <h1 style={{ ...styles.h1, marginTop: 8 }}>Admin panel</h1>
        </div>
      </div>

      <div style={{ ...styles.adminCard, marginTop: 16 }}>
        <div style={styles.historyHead}><UserPlus size={14} /> Add an employee</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
          <Field label="Employee ID"><input style={styles.input} value={newId} onChange={(e) => setNewId(e.target.value)} /></Field>
          <Field label="Name"><input style={styles.input} value={newName} onChange={(e) => setNewName(e.target.value)} /></Field>
          <Field label="Password"><PasswordInput value={newPassword} onChange={(e) => setNewPassword(e.target.value)} /></Field>
          <button type="button" onClick={addEmployee} style={{ ...styles.primaryBtn, height: 40 }}><Plus size={15} /> Add</button>
        </div>
        {error && <div style={{ ...styles.formError, marginTop: 10 }}><AlertTriangle size={14} /> {error}</div>}
      </div>

      <div style={{ ...styles.adminCard, marginTop: 14 }}>
        <div style={styles.historyHead}><User size={14} /> Roster ({roster.length})</div>
        {roster.length === 0 && <div style={styles.emptyCol}>No employees added yet.</div>}
        {roster.map((r) => (
          <div key={r.id} style={{ borderTop: `1px solid ${COLORS.line}` }}>
            <div style={{ ...styles.rosterRow, borderTop: "none" }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 14 }}>{r.name}</div>
                <div style={{ fontSize: 12, color: COLORS.sub, fontFamily: "'IBM Plex Mono', monospace" }}>
                  {r.id} · password set
                </div>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button
                  type="button"
                  style={styles.smallBtn}
                  onClick={() => {
                    setResetTargetId(resetTargetId === r.id ? null : r.id);
                    setResetDraft("");
                    setResetError("");
                  }}
                >
                  <KeyRound size={13} /> {resetTargetId === r.id ? "Cancel" : "Set new password"}
                </button>
                <button type="button" style={{ ...styles.smallBtn, color: "#B0362E" }} onClick={() => onRemoveEmployee(r.id)}>
                  <Trash2 size={13} /> Remove
                </button>
              </div>
            </div>
            {resetTargetId === r.id && (
              <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap", padding: "0 0 12px" }}>
                <Field label={`New password for ${r.name}`}>
                  <PasswordInput value={resetDraft} onChange={(e) => setResetDraft(e.target.value)} autoFocus />
                </Field>
                <button type="button" onClick={() => submitReset(r.id)} style={{ ...styles.primaryBtn, height: 40 }}>
                  Save
                </button>
                {resetError && (
                  <div style={{ ...styles.formError, width: "100%" }}>
                    <AlertTriangle size={14} /> {resetError}
                  </div>
                )}
              </div>
            )}
          </div>
        ))}
      </div>

      <div style={{ ...styles.adminCard, marginTop: 14 }}>
        <div style={styles.historyHead}><PackagePlus size={14} /> Drum styles ({drumStyles.length})</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end", marginBottom: 10 }}>
          <Field label="New drum style (e.g. 20-58-15)"><input style={styles.input} value={newStyle} onChange={(e) => setNewStyle(e.target.value)} /></Field>
          <button type="button" onClick={addStyle} style={{ ...styles.primaryBtn, height: 40 }}><Plus size={15} /> Add</button>
        </div>
        {styleError && <div style={{ ...styles.formError, marginBottom: 10 }}><AlertTriangle size={14} /> {styleError}</div>}
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
          {drumStyles.map((s) => (
            <div key={s} style={styles.styleChip}>
              <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{s}</span>
              <button type="button" onClick={() => onRemoveStyle(s)} style={styles.styleChipRemove}><X size={12} /></button>
            </div>
          ))}
        </div>
      </div>

      <div style={{ ...styles.adminCard, marginTop: 14 }}>
        <div style={styles.historyHead}><BarChart3 size={14} /> Par levels</div>
        <p style={{ fontSize: 13, color: COLORS.sub, marginBottom: 10 }}>
          Set a minimum "ready to use" count per style (New in Stores + Okay/Repaired in Dept Rack).
          The dashboard flags a style when it drops below this. Leave blank for no target.
        </p>
        {drumStyles.length === 0 && <div style={styles.emptyCol}>Add drum styles first.</div>}
        {drumStyles.map((s) => {
          const draft = parDrafts[s] !== undefined ? parDrafts[s] : (parLevels[s] ?? "");
          return (
            <div key={s} style={styles.rosterRow}>
              <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 13.5 }}>{s}</span>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input
                  style={{ ...styles.input, width: 70, padding: "6px 8px" }}
                  inputMode="numeric"
                  pattern="[0-9]*"
                  value={draft}
                  onChange={(e) => setParDrafts({ ...parDrafts, [s]: e.target.value.replace(/\D/g, "") })}
                  placeholder="—"
                />
                <button
                  type="button"
                  style={styles.smallBtn}
                  onClick={() => onSetParLevel(s, draft === "" ? null : parseInt(draft, 10))}
                >
                  Save
                </button>
              </div>
            </div>
          );
        })}
      </div>
      <Footer />
    </div>
  );
}

function DrillDownModal({ drillDown, drums, daysInStatusMap, onClose, onSelectDrum }) {
  const meta = CLASSIFICATIONS.find((c) => c.key === drillDown.key);
  const scopeLabel = drillDown.style === "ALL" ? "All styles" : drillDown.style;
  return (
    <div className="dt-overlay" style={styles.overlay} onClick={onClose}>
      <div className="dt-sheet" style={{ ...styles.panel, position: "relative" }} onClick={(e) => e.stopPropagation()}>
        <div style={styles.panelScroll}>
          <div className="dt-handle" style={styles.sheetHandle} />
          <div style={styles.panelHead}>
            <div>
              <div style={styles.eyebrow}>{scopeLabel}</div>
              <h2 style={{ ...styles.panelTitle, fontSize: 18 }}>{meta?.label}</h2>
            </div>
            <button style={styles.iconBtn} onClick={onClose}><X size={18} /></button>
          </div>
          <p style={styles.historyHint}>{drums.length} drum{drums.length === 1 ? "" : "s"} — longest-waiting first.</p>
          <div style={styles.listView}>
            {drums.length === 0 && <div style={styles.emptyCol}>No drums in this bucket.</div>}
            {drums.map((d, i) => {
              const cond = CONDITIONS[d.condition] || CONDITIONS.NEW;
              return (
                <button key={d.id} className="dt-hover" style={styles.listCard} onClick={() => onSelectDrum(d.id)}>
                  <div style={styles.listCardTop}>
                    <span style={styles.listCardId}>
                      {d.id} {i === 0 && drums.length > 1 && <span style={styles.longestBadge}>longest waiting</span>}
                    </span>
                    <span style={{ ...styles.tagCondition, color: cond.fg, background: cond.bg }}>{cond.label}</span>
                  </div>
                  <div style={styles.listCardMeta}>
                    {d.size} · {STATUS_MAP[d.status]}
                    {daysInStatusMap[d.id] !== undefined && ` · ${daysInStatusMap[d.id]}d here`}
                  </div>
                </button>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}

function MachineView({ machines, onSelectDrum }) {
  return (
    <div style={styles.dashboard}>
      <div style={styles.machineViewIntro}>
        {machines.length} machine{machines.length === 1 ? "" : "s"} currently have a drum fixed on them.
      </div>
      {machines.length === 0 && (
        <div style={styles.emptyCol}>No drums are currently marked "Fixed on Machine" with a machine number.</div>
      )}
      {machines.map((m) => (
        <div key={m.machineNo} style={styles.machineCard}>
          <div style={styles.machineCardHead}>
            <Cog size={15} color={COLORS.amberDark} />
            <span style={styles.machineCardTitle}>Machine {m.machineNo}</span>
            {m.drums.length > 1 && (
              <span style={styles.machineConflictTag}><AlertTriangle size={11} /> {m.drums.length} drums assigned</span>
            )}
          </div>
          {m.drums.map((d) => {
            const cond = CONDITIONS[d.condition] || CONDITIONS.NEW;
            return (
              <button key={d.id} type="button" className="dt-hover" style={styles.machineDrumRow} onClick={() => onSelectDrum(d.id)}>
                <div>
                  <div style={styles.machineDrumId}>{d.id}</div>
                  <div style={styles.machineDrumMeta}>
                    {d.size} · On since {fmtDate(d.lastFixedDate)}
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ ...styles.tagCondition, color: cond.fg, background: cond.bg }}>{cond.label}</span>
                  <ChevronRight size={16} color="#9AA2AA" />
                </div>
              </button>
            );
          })}
        </div>
      ))}
    </div>
  );
}

function Dashboard({ overall, rows, showOverall = true, parLevels, onTileClick }) {
  if (!overall) return null;
  return (
    <div style={styles.dashboard}>
      {showOverall && (
        <StyleStatCard
          title="All styles"
          total={overall.total}
          conditionCounts={overall.conditionCounts}
          statusCounts={overall.statusCounts}
          partyCounts={overall.partyCounts}
          avgLifeDays={overall.avgLifeDays}
          totalRepairs={overall.totalRepairs}
          classificationCounts={overall.classificationCounts}
          onTileClick={(key) => onTileClick("ALL", key)}
          emphasized
        />
      )}
      {rows.length === 0 && <div style={styles.emptyCol}>No data for this style.</div>}
      {rows.map((r) => (
        <StyleStatCard
          key={r.style}
          title={r.style}
          total={r.total}
          conditionCounts={r.conditionCounts}
          statusCounts={r.statusCounts}
          partyCounts={r.partyCounts}
          avgLifeDays={r.avgLifeDays}
          totalRepairs={r.totalRepairs}
          classificationCounts={r.classificationCounts}
          parLevel={parLevels[r.style]}
          onTileClick={(key) => onTileClick(r.style, key)}
        />
      ))}
    </div>
  );
}

function StyleStatCard({ title, total, conditionCounts, statusCounts, partyCounts, avgLifeDays, totalRepairs, classificationCounts, parLevel, onTileClick, emphasized }) {
  const usableCount = USABLE_BUCKETS.reduce((sum, k) => sum + (classificationCounts[k] || 0), 0);
  const belowPar = parLevel !== undefined && parLevel !== null && parLevel > 0 && usableCount < parLevel;
  return (
    <div style={{ ...styles.dashCard, ...(emphasized ? styles.dashCardEmphasized : {}) }}>
      <div style={styles.dashCardHead}>
        <div style={styles.dashCardTitle}>{title}</div>
        <div style={styles.dashTotalTile}>
          <div style={styles.dashTotalNumber}>{total}</div>
          <div style={styles.dashTotalLabel}>Total</div>
        </div>
      </div>

      <div style={styles.dashMiniStatsRow}>
        <div style={styles.dashMiniStat}>
          <span style={styles.dashMiniStatNumber}>{avgLifeDays}</span> avg. life days
        </div>
        <div style={styles.dashMiniStat}>
          <span style={styles.dashMiniStatNumber}>{totalRepairs}</span> total repairs
        </div>
      </div>

      {belowPar && (
        <div style={styles.parWarning}>
          <AlertTriangle size={13} /> Only {usableCount} ready to use — below your target of {parLevel}.
        </div>
      )}

      <div style={styles.dashSectionLabel}>Classification <span style={styles.dashSectionHint}>(tap a number to see the drums)</span></div>
      <div style={styles.dashTileGrid}>
        {CLASSIFICATIONS.filter((c) => c.key !== "OTHER" || (classificationCounts[c.key] || 0) > 0).map((c) => (
          <button
            key={c.key}
            type="button"
            className="dt-hover"
            style={{ ...styles.dashTile, ...styles.dashTileClickable, color: c.fg, background: c.bg }}
            onClick={() => onTileClick(c.key)}
            disabled={(classificationCounts[c.key] || 0) === 0}
          >
            <div style={styles.dashTileNumber}>{classificationCounts[c.key] || 0}</div>
            <div style={styles.dashTileLabel}>{c.label}</div>
          </button>
        ))}
      </div>

      <div style={styles.dashSectionLabel}>Condition</div>
      <div style={styles.dashTileGrid}>
        {Object.entries(CONDITIONS).map(([k, v]) => (
          <div key={k} style={{ ...styles.dashTile, color: v.fg, background: v.bg }}>
            <div style={styles.dashTileNumber}>{conditionCounts[k] || 0}</div>
            <div style={styles.dashTileLabel}>{v.label}</div>
          </div>
        ))}
      </div>

      <div style={styles.dashSectionLabel}>Status</div>
      <div style={styles.dashTileGrid}>
        {STATUSES.map((s) => (
          <div key={s.code} style={styles.dashTileNeutral}>
            <div style={styles.dashTileNumber}>{statusCounts[s.code] || 0}</div>
            <div style={styles.dashTileLabel}>{s.code}</div>
          </div>
        ))}
      </div>

      {partyCounts && partyCounts.length > 0 && (
        <>
          <div style={styles.dashSectionLabel}>With party</div>
          <div style={styles.dashPartyList}>
            {partyCounts.map((p) => (
              <div key={p.party} style={styles.dashPartyRow}>
                <span style={styles.dashPartyName}>{p.party}</span>
                <span style={styles.dashPartyCount}>{p.count}</span>
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

function SortHeader({ label, active, dir, onClick, style }) {
  return (
    <button type="button" onClick={onClick} style={{ ...styles.sortHeaderBtn, ...style }}>
      {label}
      {active && (dir === "asc" ? <ChevronRight size={11} style={{ transform: "rotate(-90deg)" }} /> : <ChevronRight size={11} style={{ transform: "rotate(90deg)" }} />)}
    </button>
  );
}

function SortChip({ label, active, dir, onClick }) {
  return (
    <button type="button" onClick={onClick} style={{ ...styles.sortChip, ...(active ? styles.sortChipActive : {}) }}>
      {label}
      {active && (dir === "asc" ? <ChevronRight size={12} style={{ transform: "rotate(-90deg)" }} /> : <ChevronRight size={12} style={{ transform: "rotate(90deg)" }} />)}
    </button>
  );
}

function DrumTag({ drum, onClick, lockedBy, daysInStatus }) {
  const c = CONDITIONS[drum.condition] || CONDITIONS.NEW;
  return (
    <button onClick={onClick} style={{ ...styles.tag, borderColor: c.ring }}>
      <div style={styles.tagRivet} />
      <div style={styles.tagId}>{drum.id}</div>
      <div style={styles.tagMeta}>Style {drum.size}</div>
      <div style={{ ...styles.tagCondition, color: c.fg, background: c.bg }}>{c.label}</div>
      {daysInStatus !== undefined && <div style={styles.tagDays}>{daysInStatus}d here</div>}
      {drum.party && <div style={styles.tagParty}>{drum.party}</div>}
      {lockedBy && (
        <div style={styles.tagLock}><Lock size={9} /> {lockedBy}</div>
      )}
    </button>
  );
}

// Formal per-drum traceability record — identification, life summary, full
// chronological history, and signature lines for physical sign-off. Doubles
// as the soft copy (the PDF itself) and the hard copy (printed from it) for
// audit purposes (e.g. IATF 16949 tooling traceability).
async function exportDrumHistoryCard(drum, historyForDrum) {
  const [{ default: jsPDF }, { default: autoTable }] = await Promise.all([
    import("jspdf"),
    import("jspdf-autotable"),
  ]);

  const doc = new jsPDF({ orientation: "portrait", unit: "pt" });
  const red = [200, 16, 46];
  const ink = [26, 32, 41];
  const marginX = 40;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.setTextColor(...ink);
  doc.text("MRF Tyres — Goa Plant", marginX, 40);
  doc.setFontSize(11);
  doc.setFont("helvetica", "normal");
  doc.text("Tyre Building Drum — History Card", marginX, 58);
  doc.setFontSize(9);
  doc.setTextColor(107, 117, 128);
  doc.text(`Generated ${fmtDate(today())}`, marginX, 72);

  const cond = CONDITIONS[drum.condition]?.label || drum.condition || "—";

  autoTable(doc, {
    startY: 86,
    margin: { left: marginX, right: marginX },
    theme: "grid",
    styles: { fontSize: 9.5, cellPadding: 6 },
    headStyles: { fillColor: red, textColor: 255 },
    head: [["Field", "Value", "Field", "Value"]],
    body: [
      ["Drum ID", drum.id, "Style / Spec", drum.size],
      ["Status", STATUS_MAP[drum.status] || drum.status, "Condition", cond],
      ["Registered On", drum.createdAt ? fmtDate(drum.createdAt) : "—", "Total Life in Service", `${drum.totalLifeDays || 0} day(s)`],
      ["Times Repaired", String(drum.repairCount || 0), "Machine No.", drum.machineNo || "—"],
      ["With Party", drum.party || "—", "PR / PO No.", [drum.prNumber, drum.poNumber].filter(Boolean).join(" / ") || "—"],
      ["Scrap Reason", drum.scrapReason || "—", "", ""],
    ],
    columnStyles: {
      0: { fontStyle: "bold", textColor: ink, cellWidth: 105 },
      2: { fontStyle: "bold", textColor: ink, cellWidth: 105 },
    },
  });

  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(...ink);
  doc.text("Life History", marginX, doc.lastAutoTable.finalY + 24);

  const sorted = [...historyForDrum].sort((a, b) => a.ts - b.ts);

  autoTable(doc, {
    startY: doc.lastAutoTable.finalY + 34,
    margin: { left: marginX, right: marginX },
    theme: "striped",
    styles: { fontSize: 8.5, cellPadding: 5 },
    headStyles: { fillColor: ink },
    head: [["Date", "Status", "Condition", "Life This Cycle (days)", "Party / Machine", "PR / PO", "Notes"]],
    body: sorted.map((h) => [
      fmtDate(h.date),
      STATUS_MAP[h.status] || h.status,
      CONDITIONS[h.condition]?.label || h.condition,
      h.cycleDays !== null && h.cycleDays !== undefined ? String(h.cycleDays) : "—",
      h.party || h.machineNo || "—",
      [h.prNumber, h.poNumber].filter(Boolean).join(" / ") || "—",
      h.notes || h.scrapReason || "—",
    ]),
    didDrawPage: () => {
      const pageCount = doc.internal.getNumberOfPages();
      doc.setFontSize(8);
      doc.setTextColor(150, 150, 150);
      doc.text(
        `Page ${doc.internal.getCurrentPageInfo().pageNumber} of ${pageCount}`,
        doc.internal.pageSize.getWidth() - marginX - 60,
        doc.internal.pageSize.getHeight() - 16
      );
    },
  });

  doc.save(`drum-history-card-${drum.id}.pdf`);
}

function DrumDetail({ drum, history, onClose, onLog, onDelete, onEditEntry, onDeleteEntry, activeLock, currentUser, notes, onAddNote, onDeleteNote }) {
  const [confirmingDelete, setConfirmingDelete] = useState(false);
  const [confirmingEntryTs, setConfirmingEntryTs] = useState(null);
  const [confirmingNoteId, setConfirmingNoteId] = useState(null);
  const [noteDraft, setNoteDraft] = useState("");
  const c = CONDITIONS[drum.condition] || CONDITIONS.NEW;
  const lockedByOther = activeLock && activeLock.by !== currentUser;

  function submitNote() {
    if (!noteDraft.trim()) return;
    onAddNote(noteDraft.trim());
    setNoteDraft("");
  }
  return (
    <div className="dt-overlay" style={styles.overlay} onClick={onClose}>
      <div className="dt-sheet" style={{ ...styles.panel, position: "relative" }} onClick={(e) => e.stopPropagation()}>
        <div style={styles.panelScroll}>
          <div className="dt-handle" style={styles.sheetHandle} />
          <div style={styles.panelHead}>
            <div>
              <div style={styles.eyebrow}>DRUM ID</div>
              <h2 style={styles.panelTitle}>{drum.id}</h2>
            </div>
            <button style={styles.iconBtn} onClick={onClose}><X size={18} /></button>
          </div>

          {lockedByOther && (
            <div style={styles.lockBanner}>
              <Lock size={13} /> Currently being edited by <strong>{activeLock.by}</strong>
            </div>
          )}

          <div style={styles.detailGrid}>
            <DetailStat label="Drum style" value={drum.size} />
            <DetailStat label="Status" value={STATUS_MAP[drum.status]} />
            <DetailStat label="Condition" value={<span style={{ ...styles.pill, color: c.fg, background: c.bg }}>{c.label}</span>} />
            {drum.machineNo && <DetailStat label="Machine" value={drum.machineNo} />}
            {drum.party && <DetailStat label="With party" value={drum.party} />}
            {drum.prNumber && <DetailStat label="PR Number" value={drum.prNumber} />}
            {drum.poNumber && <DetailStat label="PO Number" value={drum.poNumber} />}
            {drum.scrapReason && <DetailStat label="Reason for Scrapping" value={drum.scrapReason} />}
            <DetailStat label="Total life in service" value={`${drum.totalLifeDays || 0} day${(drum.totalLifeDays || 0) === 1 ? "" : "s"}`} />
            {drum.status === "M/C" && drum.lastFixedDate && (
              <DetailStat label="On machine since" value={fmtDate(drum.lastFixedDate)} />
            )}
            <DetailStat label="Times repaired" value={drum.repairCount || 0} />
          </div>

          <button
            style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 4, ...(lockedByOther ? styles.primaryBtnDisabled : {}) }}
            onClick={onLog}
            disabled={lockedByOther}
          >
            <Wrench size={15} /> Log status update
          </button>

          <button
            type="button"
            style={{ ...styles.smallBtn, width: "100%", justifyContent: "center", marginTop: 8, padding: "9px 0" }}
            onClick={() => exportDrumHistoryCard(drum, history)}
          >
            <Printer size={14} /> Print History Card
          </button>

          <div style={{ ...styles.historyHead, marginTop: 20 }}>
            <MessageSquare size={14} /> Notes
          </div>
          <div style={styles.noteComposer}>
            <textarea
              style={styles.noteInput}
              placeholder="Add a note about this drum..."
              value={noteDraft}
              onChange={(e) => setNoteDraft(e.target.value)}
            />
            <button
              type="button"
              style={{ ...styles.noteSendBtn, ...(!noteDraft.trim() ? { opacity: 0.5, cursor: "not-allowed" } : {}) }}
              onClick={submitNote}
              disabled={!noteDraft.trim()}
            >
              Post
            </button>
          </div>
          <div style={styles.notesList}>
            {notes.length === 0 && <div style={styles.emptyCol}>No notes yet.</div>}
            {notes.map((n) => (
              <div key={n.id} style={styles.noteRow}>
                <div style={styles.noteText}>{n.text}</div>
                <div style={styles.noteFooter}>
                  <span style={styles.noteMeta}><User size={11} /> {n.by} · {fmtDate(new Date(n.ts).toISOString().slice(0, 10))}</span>
                  <button type="button" style={styles.noteDeleteBtn} onClick={() => setConfirmingNoteId(n.id)}>
                    <Trash2 size={11} />
                  </button>
                </div>
                {confirmingNoteId === n.id && (
                  <div style={styles.entryDeleteConfirm}>
                    <span>Delete this note?</span>
                    <div style={{ display: "flex", gap: 6 }}>
                      <button type="button" style={styles.smallBtn} onClick={() => setConfirmingNoteId(null)}>Cancel</button>
                      <button
                        type="button"
                        style={{ ...styles.smallBtn, color: "#fff", background: "#B0362E", border: "none" }}
                        onClick={() => { onDeleteNote(n.id); setConfirmingNoteId(null); }}
                      >
                        Confirm
                      </button>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>

          <div style={{ ...styles.historyHead, marginTop: 20 }}>
            <Clock size={14} /> Life history
          </div>
          <p style={styles.historyHint}>Made a mistake logging an entry? Edit or remove it below — totals recalculate automatically.</p>
          <div style={styles.timeline}>
            {history.length === 0 && <div style={styles.emptyCol}>No history yet</div>}
            {history.map((h, i) => (
              <div key={i} style={styles.timelineRow}>
                <div style={styles.timelineDot} />
                <div style={styles.timelineContent}>
                  <div style={styles.timelineTop}>
                    <span style={styles.timelineStatus}>{STATUS_MAP[h.status] || h.status}</span>
                    <span style={styles.timelineDate}>{fmtDate(h.date)}</span>
                  </div>
                  <div style={styles.timelineMeta}>
                    {CONDITIONS[h.condition]?.label || h.condition}
                    {h.machineNo ? ` · Machine #${h.machineNo}` : ""}
                    {h.party ? ` · ${h.party}` : ""}
                    {h.prNumber ? ` · PR ${h.prNumber}` : ""}
                    {h.poNumber ? ` · PO ${h.poNumber}` : ""}
                    {h.scrapReason ? ` · Scrapped: ${h.scrapReason}` : ""}
                  </div>
                  {h.notes && <div style={styles.timelineNotes}>{h.notes}</div>}
                  {h.cycleDays !== null && h.cycleDays !== undefined && h.cycleDays >= 0 && (
                    <div style={styles.timelineLife}>{h.cycleDays} day{h.cycleDays === 1 ? "" : "s"} in service this cycle</div>
                  )}
                  {h.repairNumber && (
                    <div style={styles.timelineRepair}>Repair #{h.repairNumber}</div>
                  )}
                  <div style={styles.timelineFooter}>
                    <div style={styles.timelineBy}><User size={11} /> {h.by}</div>
                    <div style={{ display: "flex", gap: 4 }}>
                      <button type="button" style={styles.entryActionBtn} onClick={() => onEditEntry(h)} title="Edit this entry" disabled={lockedByOther}>
                        <Pencil size={12} />
                      </button>
                      <button type="button" style={styles.entryActionBtn} onClick={() => setConfirmingEntryTs(h.ts)} title="Delete this entry" disabled={lockedByOther}>
                        <Trash2 size={12} />
                      </button>
                    </div>
                  </div>
                  {confirmingEntryTs === h.ts && (
                    <div style={styles.entryDeleteConfirm}>
                      <span>Delete this entry?</span>
                      <div style={{ display: "flex", gap: 6 }}>
                        <button type="button" style={styles.smallBtn} onClick={() => setConfirmingEntryTs(null)}>Cancel</button>
                        <button
                          type="button"
                          style={{ ...styles.smallBtn, color: "#fff", background: "#B0362E", border: "none" }}
                          onClick={() => { onDeleteEntry(h); setConfirmingEntryTs(null); }}
                        >
                          Confirm
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>

        {confirmingDelete && (
          <div style={styles.deleteConfirmPopover}>
            <div style={{ fontSize: 12.5, marginBottom: 8 }}>
              Delete <strong>{drum.id}</strong> and its full history? This can't be undone.
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button type="button" style={{ ...styles.smallBtn, flex: 1, justifyContent: "center" }} onClick={() => setConfirmingDelete(false)}>
                Cancel
              </button>
              <button type="button" style={{ ...styles.smallBtn, flex: 1, justifyContent: "center", color: "#fff", background: "#B0362E", border: "none" }} onClick={onDelete}>
                <Trash2 size={12} /> Confirm
              </button>
            </div>
          </div>
        )}

        <button
          type="button"
          style={{ ...styles.deleteFab, ...(lockedByOther ? styles.deleteFabDisabled : {}) }}
          onClick={() => !lockedByOther && setConfirmingDelete((v) => !v)}
          title={lockedByOther ? `Locked by ${activeLock.by}` : "Delete this drum"}
          disabled={lockedByOther}
        >
          <Trash2 size={14} />
        </button>
      </div>
    </div>
  );
}

function DetailStat({ label, value }) {
  return (
    <div>
      <div style={styles.statLabel}>{label}</div>
      <div style={styles.statValue}>{value}</div>
    </div>
  );
}

function ChangePasswordModal({ onClose, onSave }) {
  const [currentPw, setCurrentPw] = useState("");
  const [newPw, setNewPw] = useState("");
  const [newPw2, setNewPw2] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState(false);

  async function submit() {
    setError("");
    if (!currentPw) {
      setError("Enter your current password.");
      return;
    }
    if (newPw.length < 4) {
      setError("New password should be at least 4 characters.");
      return;
    }
    if (newPw !== newPw2) {
      setError("New passwords don't match.");
      return;
    }
    setBusy(true);
    const result = await onSave(currentPw, newPw);
    setBusy(false);
    if (result.ok) {
      setDone(true);
    } else {
      setError(result.error || "Couldn't change password.");
    }
  }

  return (
    <div className="dt-overlay" style={styles.overlay} onClick={onClose}>
      <div className="dt-sheet" style={styles.modal} onClick={(e) => e.stopPropagation()}>
        <div className="dt-handle" style={styles.sheetHandle} />
        <div style={styles.panelHead}>
          <div style={styles.eyebrow}>ACCOUNT</div>
          <button type="button" style={styles.iconBtn} onClick={onClose}><X size={18} /></button>
        </div>
        <h2 style={{ ...styles.panelTitle, marginBottom: 16 }}>
          <KeyRound size={18} style={{ marginRight: 8, verticalAlign: -3 }} />
          Change Password
        </h2>
        {done ? (
          <>
            <div style={{ ...styles.formError, color: "#2F5233", background: "#DCEEDD" }}>
              <CheckCircle2 size={14} /> Password updated.
            </div>
            <button
              type="button"
              onClick={onClose}
              style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 12 }}
            >
              Done
            </button>
          </>
        ) : (
          <div style={styles.form}>
            <Field label="Current Password">
              <PasswordInput
                placeholder="Current password"
                value={currentPw}
                onChange={(e) => setCurrentPw(e.target.value)}
                autoFocus
              />
            </Field>
            <Field label="New Password">
              <PasswordInput placeholder="New password" value={newPw} onChange={(e) => setNewPw(e.target.value)} />
            </Field>
            <Field label="Confirm New Password">
              <PasswordInput
                placeholder="Confirm new password"
                value={newPw2}
                onChange={(e) => setNewPw2(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") submit();
                }}
              />
            </Field>
            {error && (
              <div style={styles.formError}>
                <AlertTriangle size={14} /> {error}
              </div>
            )}
            <button
              type="button"
              disabled={busy}
              onClick={submit}
              style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 4 }}
            >
              {busy ? "Saving…" : "Save New Password"}
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

function AddDrumModal({ onClose, onSave, existingIds, drumStyles, allDrums }) {
  const [id, setId] = useState(uid());
  const [size, setSize] = useState(drumStyles[0] || "");
  const [status, setStatus] = useState("STRS");
  const [condition, setCondition] = useState("NEW");
  const [machineNo, setMachineNo] = useState("");
  const [party, setParty] = useState("");
  const [prNumber, setPrNumber] = useState("");
  const [poNumber, setPoNumber] = useState("");
  const [scrapReason, setScrapReason] = useState("");
  const [error, setError] = useState("");

  const machineConflict = status === "M/C" && machineNo.trim()
    ? allDrums.find((d) => d.status === "M/C" && (d.machineNo || "").trim() === machineNo.trim())
    : null;

  // Sending a drum out to a party for repair means it's in repair, and
  // scrapping a drum means condition is Scrap — set those automatically instead
  // of relying on someone to remember.
  function handleStatusChange(next) {
    setStatus(next);
    if (next === "URWP") setCondition("IN REPAIR");
    if (next === "SCRAP") setCondition("SCRAP");
  }

  function submit() {
    if (!id.trim() || !size) {
      setError("Drum ID and style are required.");
      return;
    }
    if (existingIds.includes(id.trim())) {
      setError("A drum with this ID already exists.");
      return;
    }
    onSave({
      id: id.trim(),
      size,
      status,
      condition,
      machineNo: machineNo.trim(),
      party: party.trim(),
      prNumber: prNumber.trim(),
      poNumber: poNumber.trim(),
      scrapReason: scrapReason.trim(),
    });
  }

  return (
    <div className="dt-overlay" style={styles.overlay} onClick={onClose}>
      <div className="dt-sheet" style={styles.modal} onClick={(e) => e.stopPropagation()}>
        <div className="dt-handle" style={styles.sheetHandle} />
        <div style={styles.panelHead}>
          <div style={styles.eyebrow}>NEW ENTRY</div>
          <button type="button" style={styles.iconBtn} onClick={onClose}><X size={18} /></button>
        </div>
        <h2 style={{ ...styles.panelTitle, marginBottom: 16 }}><PackagePlus size={18} style={{ marginRight: 8, verticalAlign: -3 }} />Add a drum</h2>
        <div style={styles.form}>
          <Field label="Drum ID">
            <input style={styles.input} value={id} onChange={(e) => setId(e.target.value)} />
          </Field>
          <Field label="Drum style">
            {drumStyles.length === 0 ? (
              <div style={{ fontSize: 12.5, color: COLORS.sub }}>
                No drum styles set up yet — add some from the Admin panel first.
              </div>
            ) : (
              <CustomSelect value={size} onChange={setSize} options={drumStyles.map((s) => ({ value: s, label: s }))} />
            )}
          </Field>
          <Field label="Status">
            <CustomSelect
              value={status}
              onChange={handleStatusChange}
              options={STATUSES.map((s) => ({ value: s.code, label: `${s.code} — ${s.label}` }))}
            />
          </Field>
          <Field label="Condition">
            <CustomSelect
              value={condition}
              onChange={setCondition}
              options={Object.entries(CONDITIONS).map(([k, v]) => ({ value: k, label: v.label }))}
            />
          </Field>
          {status === "M/C" && (
            <Field label="Machine no. (if fixed on machine)">
              <input
                style={styles.input}
                value={machineNo}
                onChange={(e) => setMachineNo(e.target.value.replace(/\D/g, ""))}
                inputMode="numeric"
                pattern="[0-9]*"
                placeholder="e.g. 42"
              />
            </Field>
          )}
          {machineConflict && (
            <div style={styles.formError}>
              <AlertTriangle size={14} /> Machine {machineNo.trim()} already has {machineConflict.id} fixed on it.
            </div>
          )}
          <Field label="Party (if with an outside party)">
            <input style={styles.input} value={party} onChange={(e) => setParty(e.target.value)} />
          </Field>
          {status === "URWP" && (
            <>
              <Field label="PR Number">
                <input style={styles.input} value={prNumber} onChange={(e) => setPrNumber(e.target.value)} placeholder="e.g. PR-2024-0451" />
              </Field>
              <Field label="PO Number">
                <input style={styles.input} value={poNumber} onChange={(e) => setPoNumber(e.target.value)} placeholder="e.g. PO-8821" />
              </Field>
            </>
          )}
          {status === "SCRAP" && (
            <>
              <div style={styles.formError}>
                <AlertTriangle size={14} /> Scrapping a drum sets condition to Scrap and takes it out of active circulation.
              </div>
              <Field label="Reason for Scrapping">
                <input style={styles.input} value={scrapReason} onChange={(e) => setScrapReason(e.target.value)} placeholder="e.g. Shell cracked, beyond repair" />
              </Field>
            </>
          )}
          {error && <div style={styles.formError}><AlertTriangle size={14} /> {error}</div>}
          <button type="button" onClick={submit} style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 4 }}>
            Add drum to register
          </button>
        </div>
      </div>
    </div>
  );
}

function LogUpdateModal({ drum, allDrums, onClose, onSave }) {
  const [status, setStatus] = useState(drum.status);
  const [condition, setCondition] = useState(drum.condition);
  const [machineNo, setMachineNo] = useState(drum.machineNo || "");
  const [party, setParty] = useState(drum.party || "");
  const [prNumber, setPrNumber] = useState(drum.prNumber || "");
  const [poNumber, setPoNumber] = useState(drum.poNumber || "");
  const [scrapReason, setScrapReason] = useState(drum.scrapReason || "");
  const [date, setDate] = useState(today());
  const [fixedDate, setFixedDate] = useState(today());
  const [notes, setNotes] = useState("");

  const wasOnMachine = drum.status === "M/C";
  const goingOnMachine = status === "M/C";
  const comingOffMachine = wasOnMachine && !goingOnMachine;
  const goingToParty = status === "URWP";
  const isNewTripToParty = goingToParty && drum.status !== "URWP";
  const goingToScrap = status === "SCRAP";
  const isNewToScrap = goingToScrap && drum.status !== "SCRAP";

  const cycleDays = comingOffMachine && drum.lastFixedDate ? diffDays(drum.lastFixedDate, date) : null;
  const totalAfter = (drum.totalLifeDays || 0) + (cycleDays && cycleDays > 0 ? cycleDays : 0);

  const machineConflict = goingOnMachine && machineNo.trim()
    ? allDrums.find((d) => d.id !== drum.id && d.status === "M/C" && (d.machineNo || "").trim() === machineNo.trim())
    : null;

  // Sending a drum out to a party for repair means it's in repair, and
  // scrapping a drum means condition is Scrap — set those automatically instead
  // of relying on someone to remember.
  function handleStatusChange(next) {
    setStatus(next);
    if (next === "URWP") setCondition("IN REPAIR");
    if (next === "SCRAP") setCondition("SCRAP");
  }

  function submit() {
    onSave({ status, condition, machineNo, party, prNumber, poNumber, scrapReason, date, fixedDate, notes });
  }

  return (
    <div className="dt-overlay" style={styles.overlay} onClick={onClose}>
      <div className="dt-sheet" style={styles.modal} onClick={(e) => e.stopPropagation()}>
        <div className="dt-handle" style={styles.sheetHandle} />
        <div style={styles.panelHead}>
          <div style={styles.eyebrow}>{drum.id}</div>
          <button type="button" style={styles.iconBtn} onClick={onClose}><X size={18} /></button>
        </div>
        <h2 style={{ ...styles.panelTitle, marginBottom: 16 }}><Wrench size={18} style={{ marginRight: 8, verticalAlign: -3 }} />Log a status update</h2>
        <div style={styles.form}>
          <Field label="New status">
            <CustomSelect
              value={status}
              onChange={handleStatusChange}
              options={STATUSES.map((s) => ({ value: s.code, label: `${s.code} — ${s.label}` }))}
            />
          </Field>

          {goingOnMachine && (
            <Field label="Date fixed on machine">
              <input type="date" style={styles.input} value={fixedDate} onChange={(e) => setFixedDate(e.target.value)} />
            </Field>
          )}

          {comingOffMachine && (
            <>
              <Field label="Date drum was fixed on machine">
                <input style={{ ...styles.input, background: "#EEEFF0", color: COLORS.sub }} value={fmtDate(drum.lastFixedDate)} disabled />
              </Field>
              <Field label="Date of removal">
                <input type="date" style={styles.input} value={date} onChange={(e) => setDate(e.target.value)} />
              </Field>
              <div style={styles.lifeBox}>
                <div style={styles.lifeBoxRow}>
                  <span>Drum life this cycle</span>
                  <strong>{cycleDays !== null && cycleDays >= 0 ? `${cycleDays} day${cycleDays === 1 ? "" : "s"}` : "—"}</strong>
                </div>
                <div style={styles.lifeBoxRow}>
                  <span>Total life to date</span>
                  <strong>{totalAfter} day{totalAfter === 1 ? "" : "s"}</strong>
                </div>
              </div>
            </>
          )}

          {!goingOnMachine && !comingOffMachine && (
            <Field label="Date">
              <input type="date" style={styles.input} value={date} onChange={(e) => setDate(e.target.value)} />
            </Field>
          )}

          <Field label="New condition">
            <CustomSelect
              value={condition}
              onChange={setCondition}
              options={Object.entries(CONDITIONS).map(([k, v]) => ({ value: k, label: v.label }))}
            />
          </Field>
          <Field label="Machine no. (if applicable)">
            <input
              style={styles.input}
              value={machineNo}
              onChange={(e) => setMachineNo(e.target.value.replace(/\D/g, ""))}
              disabled={!goingOnMachine}
              inputMode="numeric"
              pattern="[0-9]*"
              placeholder="e.g. 42"
            />
          </Field>
          {machineConflict && (
            <div style={styles.formError}>
              <AlertTriangle size={14} /> Machine {machineNo.trim()} already has {machineConflict.id} fixed on it.
            </div>
          )}
          <Field label="Party (if applicable)">
            <input style={styles.input} value={party} onChange={(e) => setParty(e.target.value)} />
          </Field>
          {goingToParty && (
            <>
              {isNewTripToParty && (
                <div style={styles.formError}>
                  <Wrench size={14} /> Sending this drum to a party counts as one repair on its total, and sets condition to In Repair.
                </div>
              )}
              <Field label="PR Number">
                <input style={styles.input} value={prNumber} onChange={(e) => setPrNumber(e.target.value)} placeholder="e.g. PR-2024-0451" />
              </Field>
              <Field label="PO Number">
                <input style={styles.input} value={poNumber} onChange={(e) => setPoNumber(e.target.value)} placeholder="e.g. PO-8821" />
              </Field>
            </>
          )}
          {goingToScrap && (
            <>
              {isNewToScrap && (
                <div style={styles.formError}>
                  <AlertTriangle size={14} /> Scrapping this drum sets condition to Scrap and takes it out of active circulation.
                </div>
              )}
              <Field label="Reason for Scrapping">
                <input style={styles.input} value={scrapReason} onChange={(e) => setScrapReason(e.target.value)} placeholder="e.g. Shell cracked, beyond repair" />
              </Field>
            </>
          )}
          <Field label="Notes">
            <textarea style={{ ...styles.input, minHeight: 64, resize: "vertical" }} value={notes} onChange={(e) => setNotes(e.target.value)} />
          </Field>
          <button type="button" onClick={submit} style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 4 }}>
            Save update
          </button>
        </div>
      </div>
    </div>
  );
}

function EditHistoryModal({ entry, onClose, onSave }) {
  const [date, setDate] = useState(entry.date);
  const [status, setStatus] = useState(entry.status);
  const [condition, setCondition] = useState(entry.condition);
  const [party, setParty] = useState(entry.party || "");
  const [notes, setNotes] = useState(entry.notes || "");

  function submit() {
    onSave({ date, status, condition, party, notes });
  }

  return (
    <div className="dt-overlay" style={styles.overlay} onClick={onClose}>
      <div className="dt-sheet" style={styles.modal} onClick={(e) => e.stopPropagation()}>
        <div className="dt-handle" style={styles.sheetHandle} />
        <div style={styles.panelHead}>
          <div style={styles.eyebrow}>EDIT ENTRY</div>
          <button type="button" style={styles.iconBtn} onClick={onClose}><X size={18} /></button>
        </div>
        <h2 style={{ ...styles.panelTitle, marginBottom: 16 }}><Pencil size={18} style={{ marginRight: 8, verticalAlign: -3 }} />Correct this entry</h2>
        <div style={styles.form}>
          <div style={styles.formError}>
            <AlertTriangle size={14} /> Changing status or condition here recalculates this drum's life days and repair count.
          </div>
          <Field label="Date">
            <input type="date" style={styles.input} value={date} onChange={(e) => setDate(e.target.value)} />
          </Field>
          <Field label="Status">
            <CustomSelect
              value={status}
              onChange={setStatus}
              options={STATUSES.map((s) => ({ value: s.code, label: `${s.code} — ${s.label}` }))}
            />
          </Field>
          <Field label="Condition">
            <CustomSelect
              value={condition}
              onChange={setCondition}
              options={Object.entries(CONDITIONS).map(([k, v]) => ({ value: k, label: v.label }))}
            />
          </Field>
          <Field label="Party (if applicable)">
            <input style={styles.input} value={party} onChange={(e) => setParty(e.target.value)} />
          </Field>
          <Field label="Notes">
            <textarea style={{ ...styles.input, minHeight: 64, resize: "vertical" }} value={notes} onChange={(e) => setNotes(e.target.value)} />
          </Field>
          <button type="button" onClick={submit} style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 4 }}>
            Save correction
          </button>
        </div>
      </div>
    </div>
  );
}

// Shared credit line, shown on every screen. `light` switches to a
// white-on-dark treatment for screens with a dark/colored background (the
// login page) instead of the default dark-on-light used everywhere else.
function Footer({ light }) {
  return (
    <footer style={light ? styles.footerLight : styles.footer}>
      © {new Date().getFullYear()} · Designed and developed by Don Cherian
    </footer>
  );
}

function Field({ label, children }) {
  return (
    <label style={styles.field}>
      <span style={styles.fieldLabel}>{label}</span>
      {children}
    </label>
  );
}

function PasswordInput({ value, onChange, placeholder, onKeyDown, autoFocus, style }) {
  const [visible, setVisible] = useState(false);
  return (
    <div style={{ position: "relative" }}>
      <input
        type={visible ? "text" : "password"}
        placeholder={placeholder}
        value={value}
        onChange={onChange}
        onKeyDown={onKeyDown}
        autoFocus={autoFocus}
        style={{ ...styles.input, paddingRight: 40, width: "100%", ...style }}
      />
      <button
        type="button"
        tabIndex={-1}
        onClick={() => setVisible((v) => !v)}
        title={visible ? "Hide password" : "Show password"}
        style={{
          position: "absolute",
          right: 4,
          top: "50%",
          transform: "translateY(-50%)",
          background: "none",
          border: "none",
          cursor: "pointer",
          color: COLORS.sub,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          padding: 6,
        }}
      >
        {visible ? <EyeOff size={16} /> : <Eye size={16} />}
      </button>
    </div>
  );
}

// Native <select> renders as the OS's own picker on mobile (plain list,
// system colors) — it can't be styled to match the app. This is a fully
// custom dropdown that looks the same on every device.
function CustomSelect({ value, onChange, options, disabled }) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef(null);
  const selected = options.find((o) => o.value === value);

  useEffect(() => {
    if (!open) return;
    function handleOutside(e) {
      if (rootRef.current && !rootRef.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", handleOutside);
    document.addEventListener("touchstart", handleOutside);
    return () => {
      document.removeEventListener("mousedown", handleOutside);
      document.removeEventListener("touchstart", handleOutside);
    };
  }, [open]);

  return (
    <div ref={rootRef} style={{ position: "relative" }}>
      <button
        type="button"
        style={{ ...styles.input, ...styles.customSelectBtn, ...(disabled ? styles.customSelectBtnDisabled : {}) }}
        onClick={() => !disabled && setOpen((v) => !v)}
        disabled={disabled}
      >
        <span>{selected ? selected.label : "Select..."}</span>
        <ChevronDown size={16} style={{ transform: open ? "rotate(180deg)" : "none", transition: "transform 0.15s", flexShrink: 0 }} />
      </button>
      {open && (
        <div style={styles.customSelectMenu}>
          {options.map((o) => (
            <button
              key={o.value}
              type="button"
              style={{ ...styles.customSelectOption, ...(o.value === value ? styles.customSelectOptionActive : {}) }}
              onClick={() => { onChange(o.value); setOpen(false); }}
            >
              <span>{o.label}</span>
              {o.value === value && <Check size={15} />}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- style tokens ----------
// Palette: light, professional workspace — near-solid white/grey surfaces,
// crisp 1px borders, restrained shadows. MRF's brand red is the single
// accent color for anything actionable (the `amber`/`amberDark` names are
// legacy from an earlier copper palette — kept as-is to avoid touching
// every call site, but both now hold red values). Deliberately low on
// transparency/blur so it reads as a serious operations tool, not a
// marketing page.
const COLORS = {
  ground: "#F1F3F6",
  panel: "#FFFFFF",
  ink: "#1A2029",
  sub: "#5B6472",
  line: "#E1E4E9",
  amber: "#C8102E",
  amberDark: "#8E0B20",
};

// Kept as small, named tokens (not "glass") so surfaces stay easy to scan:
// solid white cards, a hint of blur only on true overlays (modal scrims,
// dropdown menus) where it helps depth without hurting legibility.
const SURFACE = {
  background: COLORS.panel,
  border: `1px solid ${COLORS.line}`,
};

const SURFACE_RAISED = {
  background: COLORS.panel,
  border: "1px solid #E7E9ED",
  backdropFilter: "blur(6px)",
  WebkitBackdropFilter: "blur(6px)",
};

const SURFACE_MUTED = {
  background: "#F7F8FA",
  border: `1px solid ${COLORS.line}`,
};

const SHADOW_ELEV = "0 20px 50px rgba(16,24,40,0.12)";
const SHADOW_POP = "0 4px 14px rgba(16,24,40,0.07)";
const SHADOW_BTN = "0 2px 6px rgba(200,16,46,0.28)";

const globalCss = `
  * { box-sizing: border-box; }
  html { background: ${COLORS.ground}; }
  body, input, select, textarea, button { font-family: 'DM Sans', ui-sans-serif, system-ui, sans-serif; }
  ::placeholder { color: #98A1AD; }
  ::selection { background: rgba(200,16,46,0.22); color: #1A2029; }
  input:focus, select:focus, textarea:focus, button:focus-visible {
    outline: 2px solid ${COLORS.amber}; outline-offset: 1px;
  }
  ::-webkit-scrollbar { width: 8px; height: 8px; }
  ::-webkit-scrollbar-track { background: #F1F3F6; }
  ::-webkit-scrollbar-thumb { background: #C7CCD3; border-radius: 4px; }
  ::-webkit-scrollbar-thumb:hover { background: #AFB6BF; }

  .dt-overlay { backdrop-filter: blur(3px); -webkit-backdrop-filter: blur(3px); }

  @media (min-width: ${DESKTOP_BREAKPOINT}px) {
    .dt-overlay { align-items: center !important; padding: 24px; }
    .dt-sheet { border-radius: 16px !important; max-height: 85vh !important; }
    .dt-handle { display: none !important; }
    .dt-hover { transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease; }
    .dt-hover:hover {
      transform: translateY(-2px);
      border-color: ${COLORS.amber} !important;
      box-shadow: 0 10px 24px rgba(16,24,40,0.12);
    }
  }
`;

const styles = {
  app: {
    minHeight: "100vh",
    background: COLORS.ground,
    color: COLORS.ink,
    fontFamily: "'DM Sans', ui-sans-serif, system-ui, sans-serif",
    padding: "20px 22px 40px",
  },
  desktopContainer: {
    maxWidth: 1200,
    margin: "0 auto",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-end",
    flexWrap: "wrap",
    gap: 12,
    marginBottom: 18,
  },
  eyebrow: {
    fontFamily: "'IBM Plex Mono', ui-monospace, monospace",
    fontSize: 11,
    letterSpacing: "0.12em",
    color: COLORS.sub,
    fontWeight: 600,
  },
  h1: {
    fontFamily: "'Space Grotesk', sans-serif",
    fontSize: 28,
    fontWeight: 700,
    margin: "4px 0 0",
    letterSpacing: "-0.01em",
    color: COLORS.ink,
  },
  headerRight: { display: "flex", alignItems: "center", gap: 10 },
  userBadge: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    fontSize: 13,
    color: COLORS.sub,
    ...SURFACE,
    borderRadius: 8,
    padding: "6px 10px",
    cursor: "pointer",
    fontFamily: "inherit",
  },
  primaryBtn: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    background: COLORS.amber,
    color: "#FFFFFF",
    border: "none",
    borderRadius: 8,
    padding: "9px 14px",
    fontSize: 14,
    fontWeight: 700,
    cursor: "pointer",
    boxShadow: SHADOW_BTN,
  },
  primaryBtnDisabled: {
    background: "#D8DBE0", color: "#8B93A0", cursor: "not-allowed", boxShadow: "none",
  },
  toolbar: { display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" },
  searchBox: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    ...SURFACE,
    borderRadius: 8,
    padding: "8px 12px",
    flex: "1 1 240px",
  },
  searchInput: { border: "none", outline: "none", fontSize: 14, width: "100%", background: "transparent", color: COLORS.ink },
  select: {
    ...SURFACE,
    borderRadius: 8,
    padding: "8px 12px",
    fontSize: 14,
    color: COLORS.ink,
  },
  countPill: {
    display: "flex",
    alignItems: "center",
    fontSize: 13,
    color: COLORS.sub,
    ...SURFACE,
    borderRadius: 8,
    padding: "8px 12px",
  },
  clearFilterBtn: {
    display: "flex", alignItems: "center", gap: 4, fontSize: 13, color: COLORS.sub,
    ...SURFACE, borderRadius: 8, padding: "8px 12px", cursor: "pointer",
  },
  summaryBanner: {
    display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10,
    background: "#EAF1FB", border: "1px solid #C7DAF2",
    borderRadius: 10, padding: "10px 14px", marginBottom: 14,
  },
  summaryBannerText: { fontSize: 12.5, color: "#1E4E8C", lineHeight: 1.4 },
  summaryBannerClose: {
    background: "none", border: "none", color: "#1E4E8C", cursor: "pointer", padding: 4, flexShrink: 0,
    display: "flex", alignItems: "center", justifyContent: "center",
  },
  board: { display: "flex", gap: 14, overflowX: "auto", paddingBottom: 8 },
  column: {
    background: "#E9EBEF",
    border: `1px solid ${COLORS.line}`,
    borderRadius: 12,
    minWidth: 220,
    flex: "0 0 220px",
    display: "flex",
    flexDirection: "column",
    height: "calc(100vh - 170px)",
  },
  columnHead: {
    padding: "12px 12px 10px",
    borderBottom: `1px solid ${COLORS.line}`,
  },
  columnCode: {
    display: "block",
    fontFamily: "'IBM Plex Mono', ui-monospace, monospace",
    fontSize: 12,
    fontWeight: 700,
    color: COLORS.amberDark,
    letterSpacing: "0.05em",
  },
  columnLabel: { display: "block", fontSize: 12.5, color: COLORS.sub, marginTop: 2, lineHeight: 1.3 },
  columnCount: {
    display: "inline-block",
    marginTop: 6,
    fontSize: 11,
    fontWeight: 700,
    color: COLORS.ink,
    background: "#FFFFFF",
    border: `1px solid ${COLORS.line}`,
    borderRadius: 20,
    padding: "1px 8px",
  },
  columnBreakdown: { display: "flex", flexWrap: "wrap", gap: 4, marginTop: 6 },
  columnBreakdownChip: {
    fontSize: 9.5, fontWeight: 700, borderRadius: 5, padding: "2px 6px",
  },
  columnBody: { flex: 1, minHeight: 0, padding: 8, display: "flex", flexDirection: "column", gap: 6, overflowY: "auto" },
  emptyCol: { fontSize: 12, color: "#9AA2AA", textAlign: "center", padding: "14px 4px", fontStyle: "italic" },
  tag: {
    position: "relative",
    textAlign: "left",
    ...SURFACE,
    border: "1.5px solid",
    borderRadius: 8,
    padding: "7px 8px 7px 18px",
    cursor: "pointer",
    boxShadow: SHADOW_POP,
    fontFamily: "inherit",
    color: COLORS.ink,
  },
  tagRivet: {
    position: "absolute",
    left: 7,
    top: "50%",
    transform: "translateY(-50%)",
    width: 5,
    height: 5,
    borderRadius: "50%",
    background: COLORS.amber,
  },
  tagId: { fontFamily: "'IBM Plex Mono', ui-monospace, monospace", fontWeight: 700, fontSize: 13, color: COLORS.ink },
  tagMeta: { fontSize: 10.5, color: COLORS.sub, marginTop: 1 },
  tagCondition: { display: "inline-block", fontSize: 9.5, fontWeight: 700, borderRadius: 4, padding: "1px 6px", marginTop: 4 },
  tagParty: { fontSize: 10, color: COLORS.sub, marginTop: 3, fontStyle: "italic" },
  tagDays: { fontSize: 9.5, color: "#9AA2AA", marginTop: 3 },
  tagLock: {
    display: "flex", alignItems: "center", gap: 3, fontSize: 9.5, fontWeight: 700,
    color: COLORS.amberDark, marginTop: 4,
  },
  overlay: {
    position: "fixed", inset: 0, background: "rgba(20,24,28,0.45)",
    display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 50,
  },
  panel: {
    ...SURFACE_RAISED, borderRadius: "16px 16px 0 0", width: "100%", maxWidth: 520,
    maxHeight: "90vh", boxShadow: "0 -12px 40px rgba(16,24,40,0.18)", color: COLORS.ink,
  },
  panelScroll: {
    maxHeight: "90vh", overflowY: "auto", padding: "6px 22px 50px",
  },
  modal: {
    ...SURFACE_RAISED, borderRadius: "16px 16px 0 0", width: "100%", maxWidth: 480,
    maxHeight: "90vh", overflowY: "auto", padding: "6px 22px 30px",
    boxShadow: "0 -12px 40px rgba(16,24,40,0.18)", color: COLORS.ink,
  },
  sheetHandle: {
    width: 40, height: 4, borderRadius: 4, background: COLORS.line, margin: "10px auto 14px",
  },
  panelHead: {
    display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6,
    position: "sticky", top: 0, background: "rgba(255,255,255,0.96)",
    paddingTop: 2, paddingBottom: 10,
    borderBottom: `1px solid ${COLORS.line}`, zIndex: 2,
  },
  panelTitle: { fontSize: 22, fontWeight: 700, margin: "2px 0 0", fontFamily: "'IBM Plex Mono', ui-monospace, monospace", color: COLORS.ink },
  iconBtn: {
    background: "#F1F3F6", border: "none", cursor: "pointer", color: COLORS.sub, padding: 8,
    borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center",
    width: 34, height: 34, flexShrink: 0,
  },
  detailGrid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: "16px 12px", margin: "18px 0 20px" },
  lockBanner: {
    display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: "#7A4A12",
    background: "#F6E3C7", border: "1px solid #D9A34B", borderRadius: 8, padding: "8px 10px", marginTop: 10,
  },
  statLabel: { fontSize: 10.5, color: COLORS.sub, textTransform: "uppercase", letterSpacing: "0.06em", fontWeight: 600 },
  statValue: { fontSize: 15, fontWeight: 700, marginTop: 3, color: COLORS.ink },
  pill: { display: "inline-block", fontSize: 12, fontWeight: 700, borderRadius: 5, padding: "2px 8px" },
  historyHead: { display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, fontWeight: 700, color: COLORS.sub, margin: "22px 0 12px", textTransform: "uppercase", letterSpacing: "0.05em" },
  timeline: { display: "flex", flexDirection: "column", gap: 2 },
  timelineRow: { display: "flex", gap: 10 },
  timelineDot: { width: 8, height: 8, borderRadius: "50%", background: COLORS.amber, marginTop: 6, flexShrink: 0 },
  timelineContent: { paddingBottom: 18, borderLeft: `1px solid ${COLORS.line}`, marginLeft: -14, paddingLeft: 18, flex: 1 },
  timelineTop: { display: "flex", justifyContent: "space-between", gap: 8 },
  timelineStatus: { fontSize: 13.5, fontWeight: 700, color: COLORS.ink },
  timelineDate: { fontSize: 12, color: COLORS.sub, whiteSpace: "nowrap" },
  timelineMeta: { fontSize: 12.5, color: COLORS.sub, marginTop: 2 },
  timelineNotes: { fontSize: 13, marginTop: 4, color: "#3A4147" },
  timelineLife: { fontSize: 12, marginTop: 4, color: COLORS.amberDark, fontWeight: 700 },
  timelineRepair: { fontSize: 12, marginTop: 4, color: "#28454C", fontWeight: 700 },
  timelineBy: { display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: "#9AA2AA", marginTop: 5 },
  historyHint: { fontSize: 12, color: COLORS.sub, margin: "-4px 0 12px", lineHeight: 1.4 },
  timelineFooter: { display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 5 },
  entryActionBtn: {
    display: "flex", alignItems: "center", justifyContent: "center", width: 24, height: 24,
    background: "#F1F3F6", border: `1px solid ${COLORS.line}`, borderRadius: 6, color: COLORS.sub, cursor: "pointer",
  },
  entryDeleteConfirm: {
    display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8,
    background: "#F6D9D6", border: "1px solid #E3B3AE", borderRadius: 8, padding: "8px 10px", marginTop: 6, fontSize: 12.5,
  },
  noteComposer: { display: "flex", gap: 8, marginTop: 4, alignItems: "flex-end" },
  noteInput: {
    flex: 1, border: `1px solid ${COLORS.line}`, borderRadius: 10, padding: "10px 12px", fontSize: 13.5,
    background: "#FAFAFA", color: COLORS.ink, minHeight: 40, maxHeight: 100, resize: "vertical", fontFamily: "inherit",
  },
  noteSendBtn: {
    background: COLORS.amber, color: "#fff", border: "none", borderRadius: 10, padding: "10px 16px",
    fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", flexShrink: 0,
  },
  notesList: { display: "flex", flexDirection: "column", gap: 8, marginTop: 12 },
  noteRow: {
    ...SURFACE_MUTED, borderRadius: 10, padding: "10px 12px",
  },
  noteText: { fontSize: 13.5, color: COLORS.ink, lineHeight: 1.4, whiteSpace: "pre-wrap" },
  noteFooter: { display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 6 },
  noteMeta: { display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: "#9AA2AA" },
  noteDeleteBtn: {
    display: "flex", alignItems: "center", justifyContent: "center", width: 22, height: 22,
    background: "none", border: "none", color: "#B0362E", cursor: "pointer",
  },
  form: { display: "flex", flexDirection: "column", gap: 16 },
  field: { display: "flex", flexDirection: "column", gap: 5 },
  fieldLabel: { fontSize: 11.5, fontWeight: 700, color: COLORS.sub, textTransform: "uppercase", letterSpacing: "0.03em" },
  input: {
    border: `1px solid ${COLORS.line}`, borderRadius: 10, padding: "12px 13px", fontSize: 15,
    background: "#FAFAFA", color: COLORS.ink,
  },
  customSelectBtn: {
    display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%",
    cursor: "pointer", textAlign: "left", fontFamily: "inherit", fontWeight: 700,
  },
  customSelectBtnDisabled: { opacity: 0.5, cursor: "not-allowed" },
  customSelectMenu: {
    position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, zIndex: 20,
    ...SURFACE_RAISED, borderRadius: 10,
    boxShadow: "0 12px 32px rgba(16,24,40,0.18)", maxHeight: 260, overflowY: "auto", padding: 4,
  },
  customSelectOption: {
    display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%",
    background: "none", border: "none", borderRadius: 7, padding: "10px 10px", fontSize: 14.5,
    color: COLORS.ink, cursor: "pointer", textAlign: "left", fontFamily: "inherit", fontWeight: 700,
  },
  customSelectOptionActive: { background: "#FBE8EA", color: COLORS.amberDark, fontWeight: 700 },
  formError: { display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: "#B0362E", background: "#F6D9D6", borderRadius: 8, padding: "8px 10px" },
  lifeBox: {
    ...SURFACE_MUTED, borderRadius: 10, padding: "12px 14px",
    display: "flex", flexDirection: "column", gap: 6, marginTop: 2,
  },
  lifeBoxRow: { display: "flex", justifyContent: "space-between", fontSize: 13, color: COLORS.ink },
  authScreen: {
    position: "relative",
    minHeight: "100vh",
    overflow: "hidden",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 24,
    background: "linear-gradient(135deg, #ED1B2F 0%, #C8102E 32%, #6B0E1C 68%, #1C0609 100%)",
  },
  authPanel: {
    position: "relative",
    zIndex: 1,
    width: "100%",
    maxWidth: 400,
    display: "flex",
    flexDirection: "column",
    gap: 28,
  },
  authBrandRow: {
    display: "flex",
    alignItems: "center",
    gap: 14,
  },
  authBadge: {
    width: 52,
    height: 52,
    flexShrink: 0,
    borderRadius: "50%",
    objectFit: "cover",
    boxShadow: "0 10px 26px rgba(0,0,0,0.35), 0 0 0 3px rgba(255,255,255,0.15)",
  },
  authEyebrow: {
    fontFamily: "'IBM Plex Mono', ui-monospace, monospace",
    fontSize: 11,
    letterSpacing: "0.14em",
    textTransform: "uppercase",
    color: "rgba(255,255,255,0.8)",
  },
  authTitle: {
    fontFamily: "'Space Grotesk', sans-serif",
    fontSize: 24,
    fontWeight: 700,
    color: "#fff",
    marginTop: 3,
    textShadow: "0 2px 12px rgba(0,0,0,0.25)",
  },
  authTagline: {
    fontSize: 13.5,
    lineHeight: 1.6,
    color: "rgba(255,255,255,0.85)",
    maxWidth: 360,
    marginTop: -12,
  },
  authCard: {
    background: "#fff",
    borderRadius: 16,
    padding: "30px 28px 28px",
    boxShadow: "0 40px 80px rgba(0,0,0,0.4)",
    borderTop: "4px solid #C8102E",
  },
  authForgotNote: {
    fontSize: 12.5,
    color: COLORS.sub,
    lineHeight: 1.5,
    marginTop: 8,
    textAlign: "center",
  },
  savingIndicator: {
    position: "fixed", bottom: 16, left: 16, fontSize: 12, color: COLORS.sub,
    ...SURFACE_RAISED, borderRadius: 20, padding: "5px 12px",
  },
  toast: {
    position: "fixed", bottom: 16, right: 16, display: "flex", alignItems: "center", gap: 8,
    background: "#20262C", color: "#fff", borderRadius: 10, padding: "10px 16px", fontSize: 13.5,
    boxShadow: "0 8px 24px rgba(0,0,0,0.25)",
  },
  toastError: { background: "#7A2620" },
  backLink: {
    display: "inline-flex", alignItems: "center", gap: 5, background: "none", border: "none",
    color: COLORS.sub, fontSize: 12.5, cursor: "pointer", padding: 0, fontWeight: 600,
  },
  linkBtn: {
    display: "block", width: "100%", textAlign: "center", background: "none", border: "none",
    color: COLORS.amberDark, fontSize: 13, fontWeight: 600, cursor: "pointer", padding: "4px 0",
  },
  adminCard: {
    ...SURFACE, borderRadius: 12, padding: 16, color: COLORS.ink,
  },
  rosterRow: {
    display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10,
    padding: "10px 0", borderTop: `1px solid ${COLORS.line}`,
  },
  smallBtn: {
    display: "flex", alignItems: "center", gap: 4, fontSize: 12, fontWeight: 600, color: COLORS.sub,
    ...SURFACE_MUTED, borderRadius: 7, padding: "6px 9px", cursor: "pointer",
  },
  styleChip: {
    display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, fontWeight: 600, color: COLORS.ink,
    ...SURFACE_MUTED, borderRadius: 20, padding: "5px 6px 5px 12px",
  },
  styleChipRemove: {
    display: "flex", alignItems: "center", justifyContent: "center", background: "none", border: "none",
    color: COLORS.sub, cursor: "pointer", padding: 2, borderRadius: "50%",
  },
  dangerLink: {
    display: "flex", alignItems: "center", justifyContent: "center", gap: 6, width: "100%",
    background: "none", border: "none", color: "#B0362E", fontSize: 13, fontWeight: 600, cursor: "pointer", padding: "6px 0",
  },
  deleteConfirmBox: {
    marginTop: 10, background: "#F6D9D6", border: "1px solid #E3B3AE", borderRadius: 8, padding: 12,
  },
  deleteFab: {
    position: "absolute", bottom: 14, right: 14, width: 32, height: 32, borderRadius: "50%",
    background: "#B0362E", color: "#fff", border: "none", cursor: "pointer",
    display: "flex", alignItems: "center", justifyContent: "center",
    boxShadow: "0 4px 10px rgba(176,54,46,0.4)", zIndex: 6,
  },
  deleteFabDisabled: {
    background: "#D8DBE0", color: "#9AA2AA", cursor: "not-allowed", boxShadow: "none",
  },
  deleteConfirmPopover: {
    position: "absolute", bottom: 54, right: 14, width: 210, ...SURFACE_RAISED,
    borderRadius: 10, padding: 11, color: COLORS.ink,
    boxShadow: "0 12px 32px rgba(16,24,40,0.18)", zIndex: 6,
  },
  listView: { display: "flex", flexDirection: "column", gap: 8 },
  sortChipRow: { display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 4 },
  sortChip: {
    display: "flex", alignItems: "center", gap: 3, fontSize: 12, fontWeight: 700, color: COLORS.sub,
    ...SURFACE, borderRadius: 20, padding: "6px 12px",
    cursor: "pointer", fontFamily: "inherit",
  },
  sortChipActive: { background: "#FBE8EA", borderColor: COLORS.amber, color: COLORS.amberDark },
  listCard: {
    display: "flex", flexDirection: "column", gap: 4, width: "100%", textAlign: "left",
    ...SURFACE, borderRadius: 10,
    padding: "12px 14px", cursor: "pointer", fontFamily: "inherit", color: COLORS.ink,
    boxShadow: SHADOW_POP,
  },
  listCardTop: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 },
  listCardId: { fontFamily: "'IBM Plex Mono', ui-monospace, monospace", fontWeight: 700, fontSize: 15, color: COLORS.ink },
  listCardMeta: { fontSize: 13, color: COLORS.sub },
  listCardDays: { color: "#9AA2AA" },
  listCardFooter: { display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 2 },
  listCardLife: { fontSize: 11.5, color: COLORS.sub },
  listRowLock: { color: COLORS.amberDark, display: "flex", alignItems: "center" },
  dashboard: { display: "flex", flexDirection: "column", gap: 12 },
  dashCard: {
    ...SURFACE, borderRadius: 12, padding: "14px 16px", color: COLORS.ink, boxShadow: SHADOW_POP,
  },
  dashCardEmphasized: {
    background: "#FDECED", borderColor: "#F3C2C7",
  },
  dashCardHead: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 12 },
  dashCardTitle: { fontSize: 16, fontWeight: 700, fontFamily: "'IBM Plex Mono', ui-monospace, monospace", color: COLORS.ink },
  dashTotalTile: { textAlign: "right" },
  dashTotalNumber: { fontSize: 30, fontWeight: 800, color: COLORS.ink, lineHeight: 1 },
  dashTotalLabel: { fontSize: 10, fontWeight: 700, color: COLORS.sub, textTransform: "uppercase", letterSpacing: "0.05em" },
  dashSectionLabel: {
    fontSize: 10.5, fontWeight: 700, color: COLORS.sub, textTransform: "uppercase",
    letterSpacing: "0.05em", margin: "10px 0 6px",
  },
  dashSectionHint: { textTransform: "none", fontWeight: 500, letterSpacing: 0, fontSize: 10.5 },
  dashTileGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(76px, 1fr))", gap: 8 },
  dashTile: {
    borderRadius: 10, padding: "10px 6px", textAlign: "center",
  },
  dashTileClickable: {
    border: "none", cursor: "pointer", fontFamily: "inherit", width: "100%",
  },
  dashTileNeutral: {
    borderRadius: 10, padding: "10px 6px", textAlign: "center",
    ...SURFACE_MUTED, color: COLORS.ink,
  },
  dashTileNumber: { fontSize: 22, fontWeight: 800, lineHeight: 1.1 },
  dashTileLabel: { fontSize: 10, fontWeight: 600, marginTop: 3, opacity: 0.85 },
  parWarning: {
    display: "flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 600, color: "#7A2620",
    background: "#F6D9D6", border: "1px solid #E3B3AE", borderRadius: 8, padding: "8px 10px", marginBottom: 12,
  },
  longestBadge: {
    fontSize: 9, fontWeight: 700, color: COLORS.amberDark, background: "#FBE8EA",
    borderRadius: 4, padding: "1px 5px", marginLeft: 6, verticalAlign: "middle",
  },
  dashMiniStatsRow: { display: "flex", gap: 14, marginBottom: 12, flexWrap: "wrap" },
  dashMiniStat: { fontSize: 12.5, color: COLORS.sub },
  dashMiniStatNumber: { fontSize: 16, fontWeight: 800, color: COLORS.ink, marginRight: 4 },
  dashPartyList: { display: "flex", flexDirection: "column", gap: 4 },
  dashPartyRow: {
    display: "flex", justifyContent: "space-between", alignItems: "center",
    ...SURFACE_MUTED, borderRadius: 7, padding: "7px 10px",
  },
  dashPartyName: { fontSize: 12.5, fontWeight: 600, color: COLORS.ink },
  dashPartyCount: { fontSize: 13, fontWeight: 800, color: COLORS.amberDark },
  machineViewIntro: { fontSize: 12.5, color: COLORS.sub, padding: "0 2px" },
  machineCard: {
    ...SURFACE, borderRadius: 12, padding: 14,
    display: "flex", flexDirection: "column", gap: 8, color: COLORS.ink, boxShadow: SHADOW_POP,
  },
  machineCardHead: { display: "flex", alignItems: "center", gap: 8 },
  machineCardTitle: { fontSize: 15, fontWeight: 800, fontFamily: "'IBM Plex Mono', ui-monospace, monospace", color: COLORS.ink },
  machineConflictTag: {
    display: "flex", alignItems: "center", gap: 4, fontSize: 10.5, fontWeight: 700, color: "#7A2620",
    background: "#F6D9D6", borderRadius: 20, padding: "2px 8px", marginLeft: "auto",
  },
  machineDrumRow: {
    display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%",
    ...SURFACE_MUTED, borderRadius: 8, padding: "9px 12px",
    cursor: "pointer", fontFamily: "inherit", textAlign: "left", color: COLORS.ink,
  },
  machineDrumId: { fontFamily: "'IBM Plex Mono', ui-monospace, monospace", fontWeight: 700, fontSize: 13.5 },
  machineDrumMeta: { fontSize: 11.5, color: COLORS.sub, marginTop: 1 },
  listHeaderRow: {
    display: "flex", alignItems: "center", gap: 12, padding: "0 14px 4px",
  },
  sortHeaderBtn: {
    display: "flex", alignItems: "center", gap: 3, background: "none", border: "none", cursor: "pointer",
    fontSize: 11, fontWeight: 700, color: COLORS.sub, textTransform: "uppercase", letterSpacing: "0.04em",
    padding: 0, fontFamily: "inherit",
  },
  menuOverlay: {
    position: "fixed", inset: 0, background: "rgba(20,24,28,0.4)", zIndex: 60,
    display: "flex", alignItems: "stretch", justifyContent: "flex-start",
  },
  menuDrawer: {
    width: 260, maxWidth: "80vw", ...SURFACE_RAISED, height: "100%",
    padding: 20, boxShadow: "6px 0 30px rgba(16,24,40,0.16)", display: "flex", flexDirection: "column",
    color: COLORS.ink, borderRadius: 0,
  },
  menuNav: { display: "flex", flexDirection: "column", gap: 4, marginTop: 14 },
  menuItem: {
    display: "flex", alignItems: "center", gap: 10, width: "100%", textAlign: "left",
    background: "none", border: "none", borderRadius: 8, padding: "11px 10px",
    fontSize: 14.5, fontWeight: 600, color: COLORS.ink, cursor: "pointer", fontFamily: "inherit",
  },
  menuItemActive: { background: "#FBE8EA", color: COLORS.amberDark },
  menuDivider: { height: 1, background: COLORS.line, margin: "14px 0" },
  footer: {
    textAlign: "center",
    padding: "18px 12px 6px",
    fontSize: 11.5,
    fontWeight: 700,
    color: COLORS.ink,
    fontFamily: "'IBM Plex Mono', ui-monospace, monospace",
    letterSpacing: "0.02em",
  },
  footerLight: {
    position: "absolute",
    bottom: 0,
    left: 0,
    right: 0,
    zIndex: 1,
    textAlign: "center",
    padding: "18px 12px 6px",
    fontSize: 11.5,
    fontWeight: 700,
    color: "rgba(255,255,255,0.75)",
    fontFamily: "'IBM Plex Mono', ui-monospace, monospace",
    letterSpacing: "0.02em",
    textShadow: "0 1px 6px rgba(0,0,0,0.3)",
  },
};
