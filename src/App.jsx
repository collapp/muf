import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  Search,
  X,
  ChevronRight,
  ChevronLeft,
  RotateCcw,
  FileSpreadsheet,
  MapPin,
  Hash,
  Calendar,
  Wallet,
  Plus,
  Check,
  AlertCircle,
  Phone,
  Briefcase,
  Store,
  UserCog,
  ClipboardList,
  LayoutDashboard,
  List,
  PieChart,
  ListChecks,
  AlertTriangle,
  Building2,
  Trophy,
  Users,
  CheckCircle2,
  TrendingUp,
  Download,
  Save,
  Lightbulb,
} from "lucide-react";

// ---------- field mapping ----------
const norm = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");

const FIELD_CANDIDATES = {
  cabang: ["NAMACABANG", "CABANG"],
  noKontrak: ["CONTRACTNO", "NOKONTRAK"],
  konsumen: ["CUSTNAME", "KONSUMEN", "NAMAKONSUMEN"],
  alamat: ["ALAMATNASABAH", "ALAMATKTP", "ALAMATTAGIH", "ALAMAT"],
  noPolisi: ["NOPOL", "NOPOLISI"],
  namaBpkb: ["NAMABPKB"],
  merk: ["BRAND", "MERK"],
  tipe: ["MODEL", "TIPE", "TYPE"],
  tahun: ["TAHUNKENDARAAN", "TAHUN"],
  noMesin: ["NOMESIN"],
  noRangka: ["NORANGKA"],
  tglWO: ["WODATENEW", "TGLWO", "WODATE"],
  tenor: ["TENOR"],
  balPrin: ["BALPRIN", "BALPRINT", "SISAHUTANG"],
  noHp: ["NOHP"],
  pekerjaan: ["PEKERJAAN"],
  dealer: ["DEALER"],
  alamatKtp: ["ALAMATKTP"],
  alamatTagih: ["ALAMATTAGIH"],
  kecamatan: ["KECAMATAN"],
  collector: ["COLLECTOR"],
  recoveryHead: ["RECOVERYHEAD"],
  statusRumah: ["STATUSRUMAHNASABAH"],
  kronologis: ["KRONOLOGIS"],
  tanggalJT: ["TANGGALJT"],
  deliquency: ["DELIQUENCY"],
  statusDebitur: ["STATUSDEBITUR"],
  matriks: ["NEWMATRIKS", "OLDMATRIKS"],
};

const FIELD_LABELS = {
  cabang: "Cabang",
  noKontrak: "No Kontrak",
  konsumen: "Konsumen",
  alamat: "Alamat",
  noPolisi: "No Polisi",
  namaBpkb: "Nama BPKB",
  merk: "Merk",
  tipe: "Tipe",
  tahun: "Tahun",
  noMesin: "No Mesin",
  noRangka: "No Rangka",
  tglWO: "Tgl WO",
  tenor: "Tenor",
  balPrin: "Sisa Hutang",
  noHp: "No HP",
  pekerjaan: "Pekerjaan",
  dealer: "Dealer",
  alamatKtp: "Alamat KTP",
  alamatTagih: "Alamat Tagih",
  kecamatan: "Kecamatan",
  collector: "Collector",
  recoveryHead: "Recovery Head",
  statusRumah: "Status Rumah",
  kronologis: "Kronologis",
  tanggalJT: "Tanggal Jatuh Tempo",
  deliquency: "Deliquency",
  statusDebitur: "Status Debitur",
  matriks: "Matriks Risiko",
};

const STATUS_OPTIONS = [
  { key: "belum_dihubungi", label: "Belum Dihubungi", color: "#8A8F98" },
  { key: "sudah_dikunjungi", label: "Sudah Dikunjungi", color: "#2A6FB0" },
  { key: "janji_bayar", label: "Janji Bayar", color: "#C98A2C" },
  { key: "proses_penarikan", label: "Proses Penarikan", color: "#B23A2E" },
  { key: "lunas", label: "Lunas / Selesai", color: "#2F7A4F" },
  { key: "tidak_ditemukan", label: "Tidak Ditemukan", color: "#6B4FA0" },
];
const statusInfo = (key) =>
  STATUS_OPTIONS.find((s) => s.key === key) || STATUS_OPTIONS[0];

const MATRIKS_COLORS = {
  LOW: "#2F7A4F",
  "MEDIUM LOW": "#4C8FA8",
  MEDIUM: "#C98A2C",
  "MEDIUM HIGH": "#D9702E",
  HIGH: "#B23A2E",
};
function matriksColor(val) {
  const key = String(val || "")
    .toUpperCase()
    .trim();
  return MATRIKS_COLORS[key] || "#8A8F98";
}

function formatRp(n) {
  const v = Number(n);
  if (!n || isNaN(v)) return "-";
  return "Rp " + v.toLocaleString("id-ID", { maximumFractionDigits: 0 });
}
function formatRpCompact(n) {
  const v = Number(n);
  if (!n || isNaN(v)) return "-";
  const abs = Math.abs(v);
  if (abs >= 1e9)
    return (
      "Rp " +
      (v / 1e9).toLocaleString("id-ID", { maximumFractionDigits: 1 }) +
      " M"
    );
  if (abs >= 1e6)
    return (
      "Rp " +
      (v / 1e6).toLocaleString("id-ID", { maximumFractionDigits: 1 }) +
      " Jt"
    );
  if (abs >= 1e3)
    return (
      "Rp " +
      (v / 1e3).toLocaleString("id-ID", { maximumFractionDigits: 0 }) +
      " rb"
    );
  return "Rp " + v.toLocaleString("id-ID");
}
function formatDate(v) {
  if (!v) return "-";
  const d = v instanceof Date ? v : new Date(v);
  if (isNaN(d.getTime())) return String(v);
  return d.toLocaleDateString("id-ID", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

// ---------- parsing ----------
function scoreSheet(rows) {
  if (!rows || !rows.length) return 0;
  const headers = Object.keys(rows[0]).map(norm);
  let score = 0;
  Object.values(FIELD_CANDIDATES).forEach((cands) => {
    if (cands.some((c) => headers.includes(c))) score++;
  });
  return score;
}

function mapRow(row, headerMap, idx) {
  const rec = { _id: `r${idx}` };
  Object.keys(FIELD_CANDIDATES).forEach((field) => {
    const key = headerMap[field];
    rec[field] = key !== undefined ? row[key] : null;
  });
  if (rec.noKontrak) rec._id = String(rec.noKontrak);
  return rec;
}

function buildHeaderMap(sampleRow) {
  const actualKeys = Object.keys(sampleRow);
  const normToActual = {};
  actualKeys.forEach((k) => (normToActual[norm(k)] = k));
  const map = {};
  Object.entries(FIELD_CANDIDATES).forEach(([field, cands]) => {
    for (const c of cands) {
      if (normToActual[c] !== undefined) {
        map[field] = normToActual[c];
        break;
      }
    }
  });
  return map;
}

async function parseWorkbook(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  let best = { score: -1, rows: null, name: null };
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    let rows;
    try {
      rows = XLSX.utils.sheet_to_json(ws, { defval: null });
    } catch (e) {
      continue;
    }
    const s = scoreSheet(rows);
    if (s > best.score) best = { score: s, rows, name };
  }
  if (!best.rows || !best.rows.length) {
    throw new Error("Tidak menemukan data yang cocok di file ini.");
  }
  const headerMap = buildHeaderMap(best.rows[0]);
  const records = best.rows
    .filter((r) => Object.values(r).some((v) => v !== null && v !== ""))
    .map((r, i) => mapRow(r, headerMap, i));
  return { records, sheetName: best.name };
}

// ---------- storage helpers ----------
// Data is kept in this browser's localStorage only — it stays on this
// device/browser and is not synced anywhere else.
async function loadJSON(key, fallback) {
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (e) {
    return fallback;
  }
}
async function saveJSON(key, value) {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
    window.localStorage.setItem(
      "wo-last-saved",
      JSON.stringify(new Date().toISOString())
    );
  } catch (e) {
    console.error("storage error", e);
  }
}

const PAGE_SIZE = 25;

export default function CollectionApp() {
  const [records, setRecords] = useState([]);
  const [activity, setActivity] = useState({}); // contractNo -> {status, notes:[{id,ts,text}]}
  const [ready, setReady] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadMsg, setUploadMsg] = useState(null);
  const [search, setSearch] = useState("");
  const [cabangFilter, setCabangFilter] = useState("");
  const [recoFilter, setRecoFilter] = useState("");
  const [matriksFilter, setMatriksFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [page, setPage] = useState(1);
  const [view, setView] = useState("list"); // 'list' | 'dashboard'
  const [selected, setSelected] = useState(null);
  const [noteDraft, setNoteDraft] = useState("");
  const [warnaDraft, setWarnaDraft] = useState("");
  const [confirmReset, setConfirmReset] = useState(false);
  const [lastSaved, setLastSaved] = useState(null);
  const fileInput = useRef(null);

  useEffect(() => {
    (async () => {
      const r = await loadJSON("wo-records", []);
      const a = await loadJSON("wo-activity", {});
      const ls = await loadJSON("wo-last-saved", null);
      setRecords(r);
      setActivity(a);
      if (ls) setLastSaved(new Date(ls));
      setReady(true);
    })();
  }, []);

  // keep the "tersimpan otomatis" indicator in sync after every save
  useEffect(() => {
    try {
      const raw = window.localStorage.getItem("wo-last-saved");
      if (raw) setLastSaved(new Date(JSON.parse(raw)));
    } catch (e) {}
  }, [records, activity]);

  const cabangList = useMemo(() => {
    const set = new Set(records.map((r) => r.cabang).filter(Boolean));
    return Array.from(set).sort();
  }, [records]);

  const recoList = useMemo(() => {
    const set = new Set(records.map((r) => r.recoveryHead).filter(Boolean));
    return Array.from(set).sort();
  }, [records]);

  const getStatus = (id) => activity[id]?.status || "belum_dihubungi";

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return records.filter((r) => {
      if (cabangFilter && r.cabang !== cabangFilter) return false;
      if (recoFilter && r.recoveryHead !== recoFilter) return false;
      if (
        matriksFilter &&
        String(r.matriks || "").toUpperCase().trim() !== matriksFilter
      )
        return false;
      if (statusFilter && getStatus(r._id) !== statusFilter) return false;
      if (!q) return true;
      return [
        r.noKontrak,
        r.konsumen,
        r.noPolisi,
        r.noMesin,
        r.noRangka,
        r.namaBpkb,
        r.noHp,
      ]
        .filter(Boolean)
        .some((v) => String(v).toLowerCase().includes(q));
    });
  }, [
    records,
    search,
    cabangFilter,
    recoFilter,
    matriksFilter,
    statusFilter,
    activity,
  ]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const pageRecords = filtered.slice(
    (page - 1) * PAGE_SIZE,
    page * PAGE_SIZE
  );

  const totalOutstanding = useMemo(
    () => records.reduce((sum, r) => sum + (Number(r.balPrin) || 0), 0),
    [records]
  );

  useEffect(
    () => setPage(1),
    [search, cabangFilter, recoFilter, matriksFilter, statusFilter]
  );

  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    setUploadMsg(null);
    try {
      const { records: parsed, sheetName } = await parseWorkbook(file);
      setRecords(parsed);
      await saveJSON("wo-records", parsed);
      setUploadMsg({
        ok: true,
        text: `${parsed.length} kontrak dimuat dari sheet "${sheetName}".`,
      });
    } catch (err) {
      setUploadMsg({ ok: false, text: err.message || "Gagal membaca file." });
    } finally {
      setUploading(false);
      if (fileInput.current) fileInput.current.value = "";
    }
  }

  async function setContractStatus(id, statusKey) {
    const next = {
      ...activity,
      [id]: { ...(activity[id] || { notes: [] }), status: statusKey },
    };
    setActivity(next);
    await saveJSON("wo-activity", next);
  }

  async function saveWarna(id) {
    const next = {
      ...activity,
      [id]: {
        ...(activity[id] || { notes: [] }),
        warnaKendaraan: warnaDraft.trim(),
      },
    };
    setActivity(next);
    await saveJSON("wo-activity", next);
  }

  async function addNote(id) {
    const text = noteDraft.trim();
    if (!text) return;
    const entry = { id: Date.now(), ts: new Date().toISOString(), text };
    const prevEntry = activity[id] || {
      status: "belum_dihubungi",
      notes: [],
    };
    const next = {
      ...activity,
      [id]: { ...prevEntry, notes: [entry, ...(prevEntry.notes || [])] },
    };
    setActivity(next);
    setNoteDraft("");
    await saveJSON("wo-activity", next);
  }

  async function resetAll() {
    setRecords([]);
    setActivity({});
    await saveJSON("wo-records", []);
    await saveJSON("wo-activity", {});
    setConfirmReset(false);
    setSelected(null);
  }

  function exportBackup() {
    const rows = records.map((r) => {
      const act = activity[r._id] || {};
      const notesText = (act.notes || [])
        .slice()
        .reverse()
        .map(
          (n) =>
            `[${new Date(n.ts).toLocaleString("id-ID")}] ${n.text}`
        )
        .join(" | ");
      return {
        "No Kontrak": r.noKontrak,
        Konsumen: r.konsumen,
        Cabang: r.cabang,
        "No Polisi": r.noPolisi,
        Merk: r.merk,
        Tipe: r.tipe,
        Tahun: r.tahun,
        "No Mesin": r.noMesin,
        "No Rangka": r.noRangka,
        "Sisa Hutang": r.balPrin,
        "Matriks Risiko": r.matriks,
        "Recovery Head": r.recoveryHead,
        "Status Collection": statusInfo(getStatus(r._id)).label,
        "Warna Kendaraan (manual)": act.warnaKendaraan || "",
        "Catatan Kunjungan": notesText,
      };
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Backup Collection");
    const stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Backup_Collection_${stamp}.xlsx`);
  }

  if (!ready) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#F4F5F7] text-[#12233D]">
        Memuat…
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#F4F5F7] text-[#12233D] font-sans pb-10">
      {/* header */}
      <div className="bg-[#12233D] text-white px-4 pt-6 pb-5 sticky top-0 z-20 shadow-md">
        <div className="flex items-center justify-between">
          <div>
            <p className="text-[11px] uppercase tracking-[0.18em] text-[#C98A2C] font-semibold">
              Collection Tracker
            </p>
            <h1 className="text-lg font-bold leading-tight mt-0.5">
              Portofolio Write Off
            </h1>
          </div>
          <div className="flex items-center gap-1.5">
            {records.length > 0 && (
              <button
                onClick={exportBackup}
                title="Backup ke Excel"
                className="flex items-center gap-1 bg-white/10 hover:bg-white/20 text-white text-sm px-2.5 py-2 rounded-lg transition-colors"
              >
                <Download size={16} />
              </button>
            )}
            <button
              onClick={() => fileInput.current?.click()}
              className="flex items-center gap-1.5 bg-[#C98A2C] hover:bg-[#B67923] text-[#12233D] font-semibold text-sm px-3 py-2 rounded-lg transition-colors"
            >
              <Upload size={16} />
              {records.length ? "Ganti" : "Upload"}
            </button>
          </div>
          <input
            ref={fileInput}
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFile}
            className="hidden"
          />
        </div>

        {records.length > 0 && (
          <div className="grid grid-cols-2 gap-2 mt-4">
            <div className="bg-white/10 rounded-lg px-3 py-2">
              <p className="text-[10px] uppercase tracking-wide text-white/60">
                Total Kontrak
              </p>
              <p className="text-xl font-mono font-bold tabular-nums">
                {records.length.toLocaleString("id-ID")}
              </p>
            </div>
            <div className="bg-white/10 rounded-lg px-3 py-2">
              <p className="text-[10px] uppercase tracking-wide text-white/60">
                Total Sisa Hutang
              </p>
              <p className="text-base font-mono font-bold tabular-nums truncate">
                {formatRp(totalOutstanding)}
              </p>
            </div>
          </div>
        )}

        {records.length > 0 && (
          <p className="flex items-center gap-1 text-[10px] text-white/40 mt-2.5">
            <Save size={11} />
            {lastSaved
              ? `Tersimpan otomatis di HP ini · ${lastSaved.toLocaleTimeString(
                  "id-ID",
                  { hour: "2-digit", minute: "2-digit" }
                )}`
              : "Tersimpan otomatis di HP ini"}
          </p>
        )}
      </div>

      {uploading && (
        <div className="px-4 pt-3 text-sm text-[#12233D]/70 flex items-center gap-2">
          <FileSpreadsheet size={16} className="animate-pulse" /> Membaca
          file…
        </div>
      )}
      {uploadMsg && !uploading && (
        <div
          className={`mx-4 mt-3 rounded-lg px-3 py-2 text-sm flex items-start gap-2 ${
            uploadMsg.ok
              ? "bg-[#2F7A4F]/10 text-[#2F7A4F]"
              : "bg-[#B23A2E]/10 text-[#B23A2E]"
          }`}
        >
          {uploadMsg.ok ? <Check size={16} /> : <AlertCircle size={16} />}
          <span>{uploadMsg.text}</span>
        </div>
      )}

      {records.length > 0 && (
        <div className="px-4 mt-4 flex gap-2">
          <button
            onClick={() => setView("list")}
            className={`flex-1 flex items-center justify-center gap-1.5 text-sm font-semibold py-2 rounded-lg transition-colors ${
              view === "list"
                ? "bg-[#12233D] text-white"
                : "bg-white text-[#12233D]/60 border border-[#12233D]/10"
            }`}
          >
            <List size={15} /> Daftar
          </button>
          <button
            onClick={() => setView("dashboard")}
            className={`flex-1 flex items-center justify-center gap-1.5 text-sm font-semibold py-2 rounded-lg transition-colors ${
              view === "dashboard"
                ? "bg-[#12233D] text-white"
                : "bg-white text-[#12233D]/60 border border-[#12233D]/10"
            }`}
          >
            <LayoutDashboard size={15} /> Dashboard
          </button>
        </div>
      )}

      {records.length === 0 ? (
        <div className="flex flex-col items-center justify-center text-center px-8 mt-20 gap-3">
          <div className="w-14 h-14 rounded-full bg-[#12233D]/5 flex items-center justify-center">
            <FileSpreadsheet size={26} className="text-[#12233D]/40" />
          </div>
          <p className="font-semibold text-[#12233D]">Belum ada data</p>
          <p className="text-sm text-[#12233D]/60 max-w-xs">
            Upload file Excel bulanan (kolom seperti CONTRACTNO, CUSTNAME,
            NOPOL, dst) untuk mulai melacak collection.
          </p>
          <button
            onClick={() => fileInput.current?.click()}
            className="mt-2 flex items-center gap-1.5 bg-[#12233D] text-white text-sm font-semibold px-4 py-2.5 rounded-lg"
          >
            <Upload size={16} /> Pilih File Excel
          </button>
        </div>
      ) : view === "dashboard" ? (
        <Dashboard
          records={records}
          activity={activity}
          onFilter={(type, value) => {
            if (type === "status") setStatusFilter(value);
            if (type === "cabang") setCabangFilter(value);
            if (type === "reco") setRecoFilter(value);
            if (type === "matriks") setMatriksFilter(value);
            setView("list");
          }}
        />
      ) : (
        <>
          {/* search + filters */}
          <div className="px-4 mt-4 space-y-2">
            <div className="relative">
              <Search
                size={16}
                className="absolute left-3 top-1/2 -translate-y-1/2 text-[#12233D]/40"
              />
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Cari no kontrak, nama, no polisi, rangka…"
                className="w-full bg-white border border-[#12233D]/10 rounded-lg pl-9 pr-8 py-2.5 text-sm outline-none focus:border-[#C98A2C] focus:ring-2 focus:ring-[#C98A2C]/20"
              />
              {search && (
                <button
                  onClick={() => setSearch("")}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-[#12233D]/40"
                >
                  <X size={15} />
                </button>
              )}
            </div>
            <div className="flex gap-2 overflow-x-auto pb-1 -mx-4 px-4">
              <select
                value={cabangFilter}
                onChange={(e) => setCabangFilter(e.target.value)}
                className="text-xs bg-white border border-[#12233D]/10 rounded-lg px-2.5 py-1.5 shrink-0"
              >
                <option value="">Semua Cabang</option>
                {cabangList.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
              <select
                value={recoFilter}
                onChange={(e) => setRecoFilter(e.target.value)}
                className="text-xs bg-white border border-[#12233D]/10 rounded-lg px-2.5 py-1.5 shrink-0"
              >
                <option value="">Semua Reco</option>
                {recoList.map((r) => (
                  <option key={r} value={r}>
                    {r}
                  </option>
                ))}
              </select>
              <select
                value={statusFilter}
                onChange={(e) => setStatusFilter(e.target.value)}
                className="text-xs bg-white border border-[#12233D]/10 rounded-lg px-2.5 py-1.5 shrink-0"
              >
                <option value="">Semua Status</option>
                {STATUS_OPTIONS.map((s) => (
                  <option key={s.key} value={s.key}>
                    {s.label}
                  </option>
                ))}
              </select>
              <span className="text-xs text-[#12233D]/50 self-center shrink-0 ml-auto">
                {filtered.length.toLocaleString("id-ID")} hasil
              </span>
            </div>

            {(matriksFilter ||
              (cabangFilter && view === "list") ||
              recoFilter ||
              statusFilter) && (
              <div className="flex items-center gap-1.5 flex-wrap">
                {matriksFilter && (
                  <span className="flex items-center gap-1 text-[11px] font-semibold bg-[#12233D]/5 text-[#12233D] px-2 py-1 rounded-full">
                    Matriks: {matriksFilter}
                    <button onClick={() => setMatriksFilter("")}>
                      <X size={11} />
                    </button>
                  </span>
                )}
                <button
                  onClick={() => {
                    setCabangFilter("");
                    setRecoFilter("");
                    setMatriksFilter("");
                    setStatusFilter("");
                  }}
                  className="text-[11px] text-[#12233D]/40 underline"
                >
                  Reset semua filter
                </button>
              </div>
            )}
          </div>

          {/* list */}
          <div className="px-4 mt-3 space-y-2">
            {pageRecords.map((r) => {
              const st = statusInfo(getStatus(r._id));
              return (
                <button
                  key={r._id}
                  onClick={() => {
                    setSelected(r);
                    setWarnaDraft(activity[r._id]?.warnaKendaraan || "");
                  }}
                  className="w-full text-left bg-white rounded-xl p-3 flex items-center gap-3 shadow-[0_1px_2px_rgba(18,35,61,0.06),0_4px_12px_-4px_rgba(18,35,61,0.10)] border border-[#12233D]/[0.04] active:scale-[0.99] transition-transform"
                  style={{ borderLeft: `4px solid ${st.color}` }}
                >
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center justify-between gap-2">
                      <p className="font-semibold text-sm truncate">
                        {r.konsumen || "Tanpa Nama"}
                      </p>
                      <span
                        className="text-[10px] font-semibold px-1.5 py-0.5 rounded shrink-0"
                        style={{
                          color: st.color,
                          backgroundColor: st.color + "1A",
                        }}
                      >
                        {st.label}
                      </span>
                    </div>
                    <p className="text-[11px] font-mono text-[#12233D]/50 mt-0.5">
                      {r.noKontrak} · {r.cabang}
                    </p>
                    <div className="flex items-center justify-between mt-1.5">
                      <span className="text-[11px] font-mono text-[#12233D]/60">
                        {r.merk} {r.tipe} · {r.noPolisi || "-"}
                      </span>
                      <span className="text-xs font-mono font-bold tabular-nums">
                        {formatRp(r.balPrin)}
                      </span>
                    </div>
                  </div>
                  <ChevronRight
                    size={16}
                    className="text-[#12233D]/30 shrink-0"
                  />
                </button>
              );
            })}
            {pageRecords.length === 0 && (
              <p className="text-center text-sm text-[#12233D]/50 py-10">
                Tidak ada data yang cocok.
              </p>
            )}
          </div>

          {/* pagination */}
          {totalPages > 1 && (
            <div className="flex items-center justify-center gap-3 mt-4">
              <button
                disabled={page === 1}
                onClick={() => setPage((p) => Math.max(1, p - 1))}
                className="p-2 rounded-lg bg-white border border-[#12233D]/10 disabled:opacity-30"
              >
                <ChevronLeft size={16} />
              </button>
              <span className="text-xs font-mono text-[#12233D]/60">
                {page} / {totalPages}
              </span>
              <button
                disabled={page === totalPages}
                onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
                className="p-2 rounded-lg bg-white border border-[#12233D]/10 disabled:opacity-30"
              >
                <ChevronRight size={16} />
              </button>
            </div>
          )}

          <div className="px-4 mt-8">
            {!confirmReset ? (
              <button
                onClick={() => setConfirmReset(true)}
                className="text-xs text-[#12233D]/40 flex items-center gap-1 mx-auto"
              >
                <RotateCcw size={12} /> Hapus semua data
              </button>
            ) : (
              <div className="text-center text-xs text-[#12233D]/70 space-y-2">
                <p>Hapus semua data & catatan collection?</p>
                <div className="flex justify-center gap-2">
                  <button
                    onClick={resetAll}
                    className="px-3 py-1.5 bg-[#B23A2E] text-white rounded-lg font-semibold"
                  >
                    Ya, hapus
                  </button>
                  <button
                    onClick={() => setConfirmReset(false)}
                    className="px-3 py-1.5 bg-[#12233D]/10 rounded-lg font-semibold"
                  >
                    Batal
                  </button>
                </div>
              </div>
            )}
          </div>
        </>
      )}

      {/* detail drawer */}
      {selected && (
        <div className="fixed inset-0 z-30 flex items-end sm:items-center sm:justify-center">
          <div
            className="absolute inset-0 bg-black/40"
            onClick={() => setSelected(null)}
          />
          <div className="relative bg-[#F4F5F7] w-full sm:max-w-lg sm:rounded-2xl rounded-t-2xl max-h-[90vh] overflow-y-auto">
            <div className="sticky top-0 bg-[#12233D] text-white px-4 py-3 flex items-center justify-between">
              <div className="min-w-0">
                <p className="font-semibold text-sm truncate">
                  {selected.konsumen}
                </p>
                <p className="text-[11px] font-mono text-white/60">
                  {selected.noKontrak}
                </p>
              </div>
              <button onClick={() => setSelected(null)} className="p-1.5">
                <X size={18} />
              </button>
            </div>

            <div className="p-4 space-y-4">
              {/* status */}
              <div>
                <p className="text-[11px] uppercase tracking-wide text-[#12233D]/50 font-semibold mb-1.5">
                  Status Collection
                </p>
                <div className="flex flex-wrap gap-1.5">
                  {STATUS_OPTIONS.map((s) => {
                    const active = getStatus(selected._id) === s.key;
                    return (
                      <button
                        key={s.key}
                        onClick={() => setContractStatus(selected._id, s.key)}
                        className="text-xs font-semibold px-2.5 py-1.5 rounded-lg border transition-colors"
                        style={
                          active
                            ? {
                                backgroundColor: s.color,
                                color: "white",
                                borderColor: s.color,
                              }
                            : {
                                backgroundColor: "white",
                                color: s.color,
                                borderColor: s.color + "40",
                              }
                        }
                      >
                        {s.label}
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* key facts */}
              <div className="bg-white rounded-xl p-3 grid grid-cols-2 gap-3">
                <Fact icon={MapPin} label="Cabang" value={selected.cabang} />
                <Fact
                  icon={Wallet}
                  label="Sisa Hutang"
                  value={formatRp(selected.balPrin)}
                  mono
                  strong
                />
                <Fact
                  icon={Hash}
                  label="No Polisi"
                  value={selected.noPolisi}
                  mono
                />
                <Fact
                  icon={Calendar}
                  label="Tgl WO"
                  value={formatDate(selected.tglWO)}
                />
                <Fact
                  label="Kendaraan"
                  value={`${selected.merk || ""} ${selected.tipe || ""} (${
                    selected.tahun || "-"
                  })`}
                />
                {selected.matriks && (
                  <div>
                    <p className="text-[10px] uppercase tracking-wide text-[#12233D]/40">
                      Matriks Risiko
                    </p>
                    <span
                      className="inline-block text-xs font-semibold px-2 py-0.5 rounded mt-0.5"
                      style={{
                        color: matriksColor(selected.matriks),
                        backgroundColor: matriksColor(selected.matriks) + "1A",
                      }}
                    >
                      {selected.matriks}
                    </span>
                  </div>
                )}
                <div>
                  <p className="text-[10px] uppercase tracking-wide text-[#12233D]/40">
                    Warna Kendaraan (isi manual)
                  </p>
                  <div className="flex gap-1.5 mt-0.5">
                    <input
                      value={warnaDraft}
                      onChange={(e) => setWarnaDraft(e.target.value)}
                      placeholder="mis. Hitam"
                      className="flex-1 text-sm bg-[#F4F5F7] border border-[#12233D]/10 rounded-md px-2 py-1 outline-none focus:border-[#C98A2C]"
                    />
                    <button
                      onClick={() => saveWarna(selected._id)}
                      className="text-[#12233D] bg-[#12233D]/10 rounded-md px-2"
                    >
                      <Check size={14} />
                    </button>
                  </div>
                </div>
                <Fact label="Tenor" value={`${selected.tenor || "-"} bln`} />
                <Fact label="Nama BPKB" value={selected.namaBpkb} />
                <Fact label="No Mesin" value={selected.noMesin} mono />
                <Fact label="No Rangka" value={selected.noRangka} mono />
                <div className="col-span-2">
                  <Fact label="Alamat" value={selected.alamat} />
                </div>
              </div>

              {/* contact & collector info */}
              <div className="bg-white rounded-xl p-3 grid grid-cols-2 gap-3">
                <div>
                  <p className="text-[10px] uppercase tracking-wide text-[#12233D]/40 flex items-center gap-1">
                    <Phone size={10} /> No HP
                  </p>
                  {selected.noHp ? (
                    <a
                      href={`tel:${String(selected.noHp).split("//")[0].trim()}`}
                      className="text-sm mt-0.5 font-mono font-semibold text-[#2A6FB0] block"
                    >
                      {selected.noHp}
                    </a>
                  ) : (
                    <p className="text-sm mt-0.5 font-medium">-</p>
                  )}
                </div>
                <Fact
                  icon={Briefcase}
                  label="Pekerjaan"
                  value={selected.pekerjaan}
                />
                <Fact icon={Store} label="Dealer" value={selected.dealer} />
                <Fact label="Kecamatan" value={selected.kecamatan} />
                <Fact
                  icon={UserCog}
                  label="Collector"
                  value={selected.collector}
                />
                <Fact
                  icon={UserCog}
                  label="Recovery Head"
                  value={selected.recoveryHead}
                />
                <Fact
                  label="Status Debitur"
                  value={selected.statusDebitur}
                />
                <Fact
                  label="Status Rumah"
                  value={selected.statusRumah}
                />
                <Fact
                  icon={Calendar}
                  label="Tanggal Jatuh Tempo"
                  value={
                    selected.tanggalJT && selected.tanggalJT !== "00:00:00"
                      ? String(selected.tanggalJT)
                      : "-"
                  }
                />
                <Fact label="Deliquency" value={selected.deliquency} />
                {selected.kronologis && (
                  <div className="col-span-2">
                    <Fact
                      icon={ClipboardList}
                      label="Kronologis"
                      value={selected.kronologis}
                    />
                  </div>
                )}
                {(selected.alamatKtp || selected.alamatTagih) && (
                  <>
                    {selected.alamatKtp && (
                      <div className="col-span-2">
                        <Fact label="Alamat KTP" value={selected.alamatKtp} />
                      </div>
                    )}
                    {selected.alamatTagih && (
                      <div className="col-span-2">
                        <Fact
                          label="Alamat Tagih"
                          value={selected.alamatTagih}
                        />
                      </div>
                    )}
                  </>
                )}
              </div>

              {/* notes */}
              <div>
                <p className="text-[11px] uppercase tracking-wide text-[#12233D]/50 font-semibold mb-1.5">
                  Catatan Kunjungan
                </p>
                <div className="flex gap-2">
                  <textarea
                    value={noteDraft}
                    onChange={(e) => setNoteDraft(e.target.value)}
                    placeholder="Tulis hasil kunjungan / komunikasi…"
                    rows={2}
                    className="flex-1 bg-white border border-[#12233D]/10 rounded-lg px-3 py-2 text-sm outline-none focus:border-[#C98A2C] resize-none"
                  />
                  <button
                    onClick={() => addNote(selected._id)}
                    disabled={!noteDraft.trim()}
                    className="bg-[#12233D] text-white rounded-lg px-3 disabled:opacity-30 shrink-0"
                  >
                    <Plus size={16} />
                  </button>
                </div>

                <div className="mt-3 space-y-2">
                  {(activity[selected._id]?.notes || []).length === 0 && (
                    <p className="text-xs text-[#12233D]/40">
                      Belum ada catatan.
                    </p>
                  )}
                  {(activity[selected._id]?.notes || []).map((n) => (
                    <div
                      key={n.id}
                      className="bg-white rounded-lg px-3 py-2 text-sm border border-[#12233D]/5"
                    >
                      <p className="text-[10px] font-mono text-[#12233D]/40">
                        {new Date(n.ts).toLocaleString("id-ID")}
                      </p>
                      <p className="text-[#12233D] mt-0.5">{n.text}</p>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      <p className="text-center text-[10px] text-[#12233D]/30 py-4">
        © SRISP 2026
      </p>
    </div>
  );
}

function Fact({ icon: Icon, label, value, mono, strong }) {
  return (
    <div>
      <p className="text-[10px] uppercase tracking-wide text-[#12233D]/40 flex items-center gap-1">
        {Icon && <Icon size={10} />} {label}
      </p>
      <p
        className={`text-sm mt-0.5 ${mono ? "font-mono" : ""} ${
          strong ? "font-bold" : "font-medium"
        }`}
      >
        {value || "-"}
      </p>
    </div>
  );
}

function BarRow({
  label,
  value,
  max,
  color,
  formatValue,
  onClick,
  rank,
  highlight,
}) {
  const pct = max > 0 ? Math.max(4, (value / max) * 100) : 0;
  return (
    <button
      type="button"
      disabled={!onClick}
      onClick={onClick}
      className={`w-full text-left mb-3 last:mb-0 rounded-lg -mx-1.5 px-1.5 py-1 transition-colors ${
        onClick ? "active:bg-[#12233D]/[0.04]" : ""
      } ${highlight ? "bg-[#C98A2C]/[0.06]" : ""}`}
    >
      <div className="flex items-center justify-between text-xs mb-1.5">
        <span className="text-[#12233D]/80 font-medium truncate pr-2 flex items-center gap-1.5">
          {rank && (
            <span
              className="w-4 h-4 rounded-full flex items-center justify-center text-[9px] font-bold shrink-0"
              style={{
                backgroundColor: rank <= 3 ? color + "26" : "#12233D0D",
                color: rank <= 3 ? color : "#12233D80",
              }}
            >
              {rank}
            </span>
          )}
          {label}
          {onClick && (
            <ChevronRight size={11} className="text-[#12233D]/25 shrink-0" />
          )}
        </span>
        <span className="font-mono font-bold shrink-0 text-[#12233D]">
          {formatValue ? formatValue(value) : value.toLocaleString("id-ID")}
        </span>
      </div>
      <div className="h-2.5 bg-[#12233D]/[0.06] rounded-full overflow-hidden">
        <div
          className="h-full rounded-full transition-all"
          style={{
            width: `${pct}%`,
            background: `linear-gradient(90deg, ${color}CC, ${color})`,
          }}
        />
      </div>
    </button>
  );
}

function DashboardCard({ title, icon: Icon, accent = "#12233D", children }) {
  return (
    <div
      className="rounded-2xl p-4 shadow-[0_1px_2px_rgba(18,35,61,0.06),0_8px_20px_-6px_rgba(18,35,61,0.10)] border border-[#12233D]/[0.04]"
      style={{
        background: `linear-gradient(160deg, ${accent}0D 0%, #ffffff 22%)`,
      }}
    >
      <div className="flex items-center gap-2 mb-3.5">
        {Icon && (
          <span
            className="w-7 h-7 rounded-lg flex items-center justify-center shrink-0"
            style={{ backgroundColor: accent + "1A", color: accent }}
          >
            <Icon size={14} strokeWidth={2.5} />
          </span>
        )}
        <p className="text-[12px] uppercase tracking-wide text-[#12233D]/60 font-bold">
          {title}
        </p>
      </div>
      {children}
    </div>
  );
}

function KpiCard({ label, value, sub, icon: Icon, gradient }) {
  if (gradient) {
    return (
      <div
        className="rounded-2xl p-3.5 text-white overflow-hidden relative"
        style={{
          background: `linear-gradient(135deg, ${gradient[0]}, ${gradient[1]})`,
          boxShadow: `0 10px 24px -8px ${gradient[1]}80`,
        }}
      >
        {Icon && (
          <Icon
            size={54}
            strokeWidth={1.5}
            className="absolute -right-2 -bottom-2 opacity-20"
          />
        )}
        <p className="text-[10px] uppercase tracking-wide font-bold text-white/70 relative">
          {label}
        </p>
        <p className="text-lg font-mono font-extrabold tabular-nums mt-1 truncate relative">
          {value}
        </p>
        {sub && (
          <p className="text-[10px] mt-0.5 text-white/60 relative">{sub}</p>
        )}
      </div>
    );
  }
  return (
    <div className="rounded-2xl p-3.5 shadow-[0_1px_2px_rgba(18,35,61,0.06),0_8px_20px_-6px_rgba(18,35,61,0.10)] bg-white border border-[#12233D]/[0.04]">
      <p className="text-[10px] uppercase tracking-wide font-bold text-[#12233D]/40">
        {label}
      </p>
      <p className="text-lg font-mono font-extrabold tabular-nums mt-1 truncate">
        {value}
      </p>
      {sub && (
        <p className="text-[10px] mt-0.5 text-[#12233D]/40">
          {sub}
        </p>
      )}
    </div>
  );
}

function DonutChart({ data, total }) {
  let acc = 0;
  const stops = data
    .filter((d) => d.value > 0)
    .map((d) => {
      const pct = total > 0 ? (d.value / total) * 100 : 0;
      const start = acc;
      acc += pct;
      return `${d.color} ${start}% ${acc}%`;
    });
  const bg =
    stops.length > 0 ? `conic-gradient(${stops.join(", ")})` : "#EEF0F3";
  return (
    <div className="flex items-center gap-5">
      <div
        className="relative w-28 h-28 rounded-full shrink-0 shadow-[0_4px_16px_-4px_rgba(18,35,61,0.25)]"
        style={{ background: bg }}
      >
        <div className="absolute inset-2 bg-white rounded-full flex flex-col items-center justify-center">
          <span className="text-xl font-mono font-extrabold text-[#12233D]">
            {total.toLocaleString("id-ID")}
          </span>
          <span className="text-[8px] uppercase tracking-wide text-[#12233D]/40 font-semibold">
            Kontrak
          </span>
        </div>
      </div>
      <div className="flex-1 space-y-2">
        {data.map((d) => {
          const pct = total > 0 ? ((d.value / total) * 100).toFixed(0) : 0;
          return (
            <div key={d.label} className="flex items-center gap-2 text-xs">
              <span
                className="w-2.5 h-2.5 rounded-[3px] shrink-0"
                style={{ backgroundColor: d.color }}
              />
              <span className="text-[#12233D]/70 truncate flex-1">
                {d.label}
              </span>
              <span
                className="font-mono font-bold text-[10px] px-1.5 py-0.5 rounded shrink-0"
                style={{ color: d.color, backgroundColor: d.color + "1A" }}
              >
                {pct}%
              </span>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ---------- branch map ----------
// Approximate coordinates for common Sumbagut (North Sumatra region)
// cities/regencies. Branch names are matched by substring, so entries
// like "KOTA PEKANBARU 2" or "KAB. LABUHANBATU-RANTAU PRAPAT" still
// resolve to the right pin.
const CITY_COORDS = {
  "PEMATANG SIANTAR": [2.9595, 99.0687],
  "TANJUNG PINANG": [0.9186, 104.4562],
  "PADANG SIDEMPUAN": [1.3792, 99.2727],
  "TANJUNG BALAI": [2.9861, 99.8022],
  "RANTAU PRAPAT": [2.0947, 99.8371],
  "LABUHANBATU": [2.0947, 99.8371],
  "TEBING TINGGI": [3.3285, 99.1625],
  "KUALA SIMPANG": [4.3743, 98.0019],
  "GUNUNGSITOLI": [1.2907, 97.6162],
  "LHOKSEUMAWE": [5.1801, 97.1507],
  "BANDA ACEH": [5.5483, 95.3238],
  "MEULABOH": [4.1372, 96.1282],
  "BENGKALIS": [1.4649, 102.0985],
  "TAKENGON": [4.6281, 96.836],
  "SIBOLGA": [1.7427, 98.7793],
  "BIREUEN": [5.2033, 96.7009],
  "PEKANBARU": [0.5071, 101.4478],
  "KISARAN": [2.9847, 99.6134],
  "SUBULUSSALAM": [2.698, 97.941],
  "LANGSA": [4.4683, 97.9683],
  "BINJAI": [3.6001, 98.4854],
  "MEDAN": [3.5952, 98.6722],
  "BATAM": [1.1301, 104.0529],
  "DUMAI": [1.6667, 101.45],
  "SIGLI": [5.3861, 95.96],
  "SIAK": [0.8564, 102.0367],
  "DURI": [1.2657, 101.2557],
};

function findCoords(cabangName) {
  const norm = String(cabangName || "")
    .toUpperCase()
    .replace(/^KOTA\s+/, "")
    .replace(/^KAB\.?\s+/, "")
    .replace(/^KABUPATEN\s+/, "")
    .trim();
  const keys = Object.keys(CITY_COORDS).sort((a, b) => b.length - a.length);
  for (const key of keys) {
    if (norm.includes(key)) return CITY_COORDS[key];
  }
  return null;
}

function MapCard({ cabangStats, onFilter }) {
  const containerRef = useRef(null);
  const mapRef = useRef(null);
  const [available] = useState(
    typeof window !== "undefined" && !!window.L
  );

  const placed = useMemo(() => {
    return cabangStats
      .map((c) => ({ ...c, coords: findCoords(c.label) }))
      .filter((c) => c.coords);
  }, [cabangStats]);

  useEffect(() => {
    if (!available || !containerRef.current || mapRef.current) return;
    const L = window.L;
    const map = L.map(containerRef.current, {
      zoomControl: true,
      attributionControl: true,
    }).setView([2.5, 99.5], 6);
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: "© OpenStreetMap",
      maxZoom: 18,
    }).addTo(map);
    mapRef.current = map;
    return () => {
      map.remove();
      mapRef.current = null;
    };
  }, [available]);

  useEffect(() => {
    if (!mapRef.current || !window.L) return;
    const L = window.L;
    const map = mapRef.current;
    const markers = [];
    const maxVal = Math.max(1, ...placed.map((c) => c.value));

    placed.forEach((c) => {
      const size = Math.round(16 + (c.value / maxVal) * 26);
      const icon = L.divIcon({
        html: `
          <div style="position:relative;width:${size}px;height:${size}px;">
            <span class="animate-ping" style="position:absolute;inset:0;border-radius:9999px;background:#C98A2C;opacity:0.4;"></span>
            <span style="position:absolute;inset:0;border-radius:9999px;background:#C98A2C;border:2px solid white;box-shadow:0 2px 6px rgba(18,35,61,0.4);"></span>
          </div>`,
        className: "",
        iconSize: [size, size],
        iconAnchor: [size / 2, size / 2],
      });
      const marker = L.marker(c.coords, { icon }).addTo(map);
      marker.bindPopup(
        `<div style="font-family:sans-serif;min-width:150px">
          <b>${c.label}</b><br/>
          ${c.count.toLocaleString("id-ID")} kontrak<br/>
          ${formatRp(c.value)}
        </div>`
      );
      marker.on("click", () => onFilter("cabang", c.label));
      markers.push(marker);
    });

    return () => {
      markers.forEach((m) => map.removeLayer(m));
    };
  }, [placed]);

  if (!available) {
    return (
      <DashboardCard title="Peta Sebaran Cabang" icon={MapPin} accent="#C98A2C">
        <p className="text-xs text-[#12233D]/50 py-6 text-center">
          Peta hanya tersedia di versi web yang sudah di-deploy (butuh akses
          internet untuk memuat peta).
        </p>
      </DashboardCard>
    );
  }

  return (
    <DashboardCard title="Peta Sebaran Cabang" icon={MapPin} accent="#C98A2C">
      <div
        ref={containerRef}
        className="w-full h-64 rounded-xl overflow-hidden"
      />
      <p className="text-[10px] text-[#12233D]/40 mt-2">
        {placed.length} dari {cabangStats.length} cabang berhasil dipetakan ·
        ukuran titik = besar sisa hutang · tap titik untuk lihat detail
      </p>
    </DashboardCard>
  );
}

function InsightCard({ records, activity, cabangStats, recoCount, highCount }) {
  const getStatus = (id) => activity[id]?.status || "belum_dihubungi";
  const total = records.length;
  const totalOutstanding = records.reduce(
    (s, r) => s + (Number(r.balPrin) || 0),
    0
  );

  const insights = [];

  if (cabangStats.length > 0) {
    const top = cabangStats[0];
    const share = totalOutstanding > 0 ? (top.value / totalOutstanding) * 100 : 0;
    insights.push(
      `Cabang ${top.label} menyumbang ${share.toFixed(
        0
      )}% dari total sisa hutang (${formatRpCompact(
        top.value
      )}) — konsentrasi risiko terbesar ada di sini.`
    );
  }

  if (highCount > 0) {
    const pct = total > 0 ? (highCount / total) * 100 : 0;
    insights.push(
      `${highCount.toLocaleString(
        "id-ID"
      )} kontrak (${pct.toFixed(
        0
      )}%) masuk Matriks Risiko HIGH — prioritaskan penanganan kelompok ini.`
    );
  }

  const belum = records.filter(
    (r) => getStatus(r._id) === "belum_dihubungi"
  ).length;
  if (belum > 0) {
    const pct = total > 0 ? (belum / total) * 100 : 0;
    insights.push(
      `${pct.toFixed(
        0
      )}% kontrak (${belum.toLocaleString(
        "id-ID"
      )}) masih berstatus "Belum Dihubungi" — potensi kerja lapangan yang belum tersentuh.`
    );
  }

  if (recoCount.length > 0) {
    const top = recoCount[0];
    insights.push(
      `${top.label} memegang portofolio terbanyak dengan ${top.value.toLocaleString(
        "id-ID"
      )} kontrak sebagai Recovery Head.`
    );
  }

  if (cabangStats.length > 1) {
    const bottom = cabangStats[cabangStats.length - 1];
    insights.push(
      `Cabang dengan portofolio paling kecil: ${bottom.label} (${formatRpCompact(
        bottom.value
      )}, ${bottom.count.toLocaleString("id-ID")} kontrak).`
    );
  }

  if (insights.length === 0) return null;

  return (
    <DashboardCard title="Insight" icon={Lightbulb} accent="#8B5CF6">
      <ul className="space-y-2.5">
        {insights.map((text, i) => (
          <li key={i} className="flex gap-2 text-xs leading-relaxed">
            <span
              className="w-1.5 h-1.5 rounded-full mt-1.5 shrink-0"
              style={{ backgroundColor: "#8B5CF6" }}
            />
            <span className="text-[#12233D]/80">{text}</span>
          </li>
        ))}
      </ul>
    </DashboardCard>
  );
}

function Dashboard({ records, activity, onFilter }) {
  const getStatus = (id) => activity[id]?.status || "belum_dihubungi";

  const statusData = useMemo(() => {
    const counts = {};
    STATUS_OPTIONS.forEach((s) => (counts[s.key] = 0));
    records.forEach((r) => {
      const s = getStatus(r._id);
      counts[s] = (counts[s] || 0) + 1;
    });
    return STATUS_OPTIONS.map((s) => ({
      key: s.key,
      label: s.label,
      value: counts[s.key],
      color: s.color,
    }));
  }, [records, activity]);

  const matriksData = useMemo(() => {
    const counts = {};
    records.forEach((r) => {
      const m = String(r.matriks || "").toUpperCase().trim();
      if (!m) return;
      counts[m] = (counts[m] || 0) + 1;
    });
    const order = ["LOW", "MEDIUM", "MEDIUM HIGH", "HIGH"];
    return Object.entries(counts)
      .sort((a, b) => order.indexOf(a[0]) - order.indexOf(b[0]))
      .map(([label, value]) => ({
        label,
        value,
        color: matriksColor(label),
      }));
  }, [records]);

  const cabangStats = useMemo(() => {
    const map = {};
    records.forEach((r) => {
      if (!r.cabang) return;
      if (!map[r.cabang]) map[r.cabang] = { value: 0, count: 0, high: 0 };
      map[r.cabang].value += Number(r.balPrin) || 0;
      map[r.cabang].count += 1;
      if (String(r.matriks || "").toUpperCase().trim() === "HIGH") {
        map[r.cabang].high += 1;
      }
    });
    return Object.entries(map)
      .map(([label, v]) => ({ label, ...v }))
      .sort((a, b) => b.value - a.value);
  }, [records]);

  const cabangOutstanding = useMemo(
    () => cabangStats.slice(0, 10),
    [cabangStats]
  );

  const recoCount = useMemo(() => {
    const map = {};
    records.forEach((r) => {
      if (!r.recoveryHead) return;
      map[r.recoveryHead] = (map[r.recoveryHead] || 0) + 1;
    });
    return Object.entries(map)
      .map(([label, value]) => ({ label, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 10);
  }, [records]);

  const lunasStats = useMemo(() => {
    let count = 0;
    let value = 0;
    records.forEach((r) => {
      if (getStatus(r._id) === "lunas") {
        count++;
        value += Number(r.balPrin) || 0;
      }
    });
    return { count, value };
  }, [records, activity]);

  const totalOutstanding = useMemo(
    () => records.reduce((sum, r) => sum + (Number(r.balPrin) || 0), 0),
    [records]
  );
  const avgOutstanding = records.length
    ? totalOutstanding / records.length
    : 0;

  const highCount = useMemo(
    () =>
      records.filter(
        (r) => String(r.matriks || "").toUpperCase().trim() === "HIGH"
      ).length,
    [records]
  );

  const maxStatus = Math.max(1, ...statusData.map((d) => d.value));
  const maxMatriks = Math.max(1, ...matriksData.map((d) => d.value));
  const maxCabang = Math.max(1, ...cabangOutstanding.map((d) => d.value));
  const maxReco = Math.max(1, ...recoCount.map((d) => d.value));

  return (
    <div className="px-4 mt-4 space-y-3 pb-6">
      {/* colorful KPI cards */}
      <div className="grid grid-cols-2 gap-2.5">
        <KpiCard
          label="Total Sisa Hutang"
          value={formatRpCompact(totalOutstanding)}
          sub={formatRp(totalOutstanding)}
          icon={Wallet}
          gradient={["#2E6BE6", "#1B3FAE"]}
        />
        <KpiCard
          label="Total Kontrak"
          value={records.length.toLocaleString("id-ID")}
          icon={Users}
          gradient={["#8B5CF6", "#5B3EC9"]}
        />
        <KpiCard
          label="Lunas / Selesai"
          value={lunasStats.count.toLocaleString("id-ID")}
          sub={formatRpCompact(lunasStats.value)}
          icon={CheckCircle2}
          gradient={["#22B573", "#0E8A56"]}
        />
        <KpiCard
          label="Rata-rata / Kontrak"
          value={formatRpCompact(avgOutstanding)}
          icon={TrendingUp}
          gradient={["#F0932F", "#D9702E"]}
        />
      </div>

      <DashboardCard
        title="Proporsi Status Collection"
        icon={PieChart}
        accent="#12233D"
      >
        <DonutChart data={statusData} total={records.length} />
      </DashboardCard>

      <MapCard cabangStats={cabangStats} onFilter={onFilter} />

      <InsightCard
        records={records}
        activity={activity}
        cabangStats={cabangStats}
        recoCount={recoCount}
        highCount={highCount}
      />

      <DashboardCard title="Status Collection" icon={ListChecks} accent="#2A6FB0">
        {statusData.map((d) => (
          <BarRow
            key={d.label}
            label={d.label}
            value={d.value}
            max={maxStatus}
            color={d.color}
            onClick={
              d.value > 0 ? () => onFilter("status", d.key) : undefined
            }
          />
        ))}
      </DashboardCard>

      {matriksData.length > 0 && (
        <DashboardCard
          title="Matriks Risiko"
          icon={AlertTriangle}
          accent="#B23A2E"
        >
          {matriksData.map((d) => (
            <BarRow
              key={d.label}
              label={d.label}
              value={d.value}
              max={maxMatriks}
              color={d.color}
              highlight={d.label === "HIGH"}
              onClick={() => onFilter("matriks", d.label)}
            />
          ))}
        </DashboardCard>
      )}

      {cabangOutstanding.length > 0 && (
        <DashboardCard
          title="Top 10 Cabang — Sisa Hutang"
          icon={Building2}
          accent="#C98A2C"
        >
          {cabangOutstanding.map((d, i) => (
            <BarRow
              key={d.label}
              label={d.label}
              value={d.value}
              max={maxCabang}
              color="#C98A2C"
              formatValue={formatRpCompact}
              rank={i + 1}
              highlight={i === 0}
              onClick={() => onFilter("cabang", d.label)}
            />
          ))}
        </DashboardCard>
      )}

      {recoCount.length > 0 && (
        <DashboardCard
          title="Top 10 Recovery Head — Jumlah Kontrak"
          icon={Trophy}
          accent="#2A6FB0"
        >
          {recoCount.map((d, i) => (
            <BarRow
              key={d.label}
              label={d.label}
              value={d.value}
              max={maxReco}
              color="#2A6FB0"
              rank={i + 1}
              highlight={i === 0}
              onClick={() => onFilter("reco", d.label)}
            />
          ))}
        </DashboardCard>
      )}
    </div>
  );
}
