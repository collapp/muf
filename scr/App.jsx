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

function formatRp(n) {
  const v = Number(n);
  if (!n || isNaN(v)) return "-";
  return "Rp " + v.toLocaleString("id-ID", { maximumFractionDigits: 0 });
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
  const [statusFilter, setStatusFilter] = useState("");
  const [page, setPage] = useState(1);
  const [selected, setSelected] = useState(null);
  const [noteDraft, setNoteDraft] = useState("");
  const [confirmReset, setConfirmReset] = useState(false);
  const fileInput = useRef(null);

  useEffect(() => {
    (async () => {
      const r = await loadJSON("wo-records", []);
      const a = await loadJSON("wo-activity", {});
      setRecords(r);
      setActivity(a);
      setReady(true);
    })();
  }, []);

  const cabangList = useMemo(() => {
    const set = new Set(records.map((r) => r.cabang).filter(Boolean));
    return Array.from(set).sort();
  }, [records]);

  const getStatus = (id) => activity[id]?.status || "belum_dihubungi";

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return records.filter((r) => {
      if (cabangFilter && r.cabang !== cabangFilter) return false;
      if (statusFilter && getStatus(r._id) !== statusFilter) return false;
      if (!q) return true;
      return [
        r.noKontrak,
        r.konsumen,
        r.noPolisi,
        r.noMesin,
        r.noRangka,
        r.namaBpkb,
      ]
        .filter(Boolean)
        .some((v) => String(v).toLowerCase().includes(q));
    });
  }, [records, search, cabangFilter, statusFilter, activity]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const pageRecords = filtered.slice(
    (page - 1) * PAGE_SIZE,
    page * PAGE_SIZE
  );

  const totalOutstanding = useMemo(
    () => records.reduce((sum, r) => sum + (Number(r.balPrin) || 0), 0),
    [records]
  );

  useEffect(() => setPage(1), [search, cabangFilter, statusFilter]);

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
          <button
            onClick={() => fileInput.current?.click()}
            className="flex items-center gap-1.5 bg-[#C98A2C] hover:bg-[#B67923] text-[#12233D] font-semibold text-sm px-3 py-2 rounded-lg transition-colors"
          >
            <Upload size={16} />
            {records.length ? "Ganti" : "Upload"}
          </button>
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
          </div>

          {/* list */}
          <div className="px-4 mt-3 space-y-2">
            {pageRecords.map((r) => {
              const st = statusInfo(getStatus(r._id));
              return (
                <button
                  key={r._id}
                  onClick={() => setSelected(r)}
                  className="w-full text-left bg-white rounded-xl p-3 flex items-center gap-3 shadow-sm border border-[#12233D]/5 active:scale-[0.99] transition-transform"
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
                <Fact label="Tenor" value={`${selected.tenor || "-"} bln`} />
                <Fact label="Nama BPKB" value={selected.namaBpkb} />
                <Fact label="No Mesin" value={selected.noMesin} mono />
                <Fact label="No Rangka" value={selected.noRangka} mono />
                <div className="col-span-2">
                  <Fact label="Alamat" value={selected.alamat} />
                </div>
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
