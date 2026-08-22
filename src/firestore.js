import {
  collection,
  doc,
  getDocs,
  onSnapshot,
  query,
  where,
  writeBatch,
  updateDoc,
  arrayUnion,
  addDoc,
  serverTimestamp,
} from "firebase/firestore";
import { db } from "./firebase";

const CONTRACTS = "contracts";
const UPLOADS = "uploads";
const BATCH_SIZE = 400; // Firestore hard limit is 500 ops per batch

function sanitizeId(raw, fallback) {
  const s = String(raw ?? "").trim();
  const cleaned = s.replace(/\//g, "-"); // '/' is not allowed in doc IDs
  return cleaned || fallback;
}

// ---------- live subscriptions ----------
// Active (non-archived) contracts — this is the main "Daftar" data source.
export function subscribeContracts(onData, onError) {
  const q = query(collection(db, CONTRACTS), where("archived", "==", false));
  return onSnapshot(
    q,
    (snap) => onData(snap.docs.map((d) => ({ _id: d.id, ...d.data() }))),
    onError
  );
}

// Archived contracts — shown in the "Riwayat" tab.
export function subscribeArchived(onData, onError) {
  const q = query(collection(db, CONTRACTS), where("archived", "==", true));
  return onSnapshot(
    q,
    (snap) => onData(snap.docs.map((d) => ({ _id: d.id, ...d.data() }))),
    onError
  );
}

// ---------- monthly upload: merge + auto-archive ----------
// - Contracts present in the new file: upsert their raw fields, keep
//   status/notes/warnaKendaraan untouched (merge:true), unarchive them.
// - Contracts that existed before but are missing from the new file:
//   marked archived (moved to "Riwayat"), data is kept, never deleted.
export async function mergeUpload(records, uploadedBy, onProgress) {
  const existingSnap = await getDocs(collection(db, CONTRACTS));
  const existingIds = new Set(existingSnap.docs.map((d) => d.id));

  const newIds = new Set();
  const docsToWrite = records.map((r, i) => {
    const id = sanitizeId(r.noKontrak, `row-${i}`);
    newIds.add(id);
    return { id, data: r };
  });

  const toArchive = [...existingIds].filter((id) => !newIds.has(id));
  const newCount = [...newIds].filter((id) => !existingIds.has(id)).length;
  const totalOps = docsToWrite.length + toArchive.length;
  let done = 0;

  for (let i = 0; i < docsToWrite.length; i += BATCH_SIZE) {
    const batch = writeBatch(db);
    docsToWrite.slice(i, i + BATCH_SIZE).forEach(({ id, data }) => {
      batch.set(
        doc(db, CONTRACTS, id),
        { ...data, archived: false, updatedAt: serverTimestamp() },
        { merge: true }
      );
    });
    await batch.commit();
    done += Math.min(BATCH_SIZE, docsToWrite.length - i);
    onProgress?.(done, totalOps);
  }

  for (let i = 0; i < toArchive.length; i += BATCH_SIZE) {
    const batch = writeBatch(db);
    toArchive.slice(i, i + BATCH_SIZE).forEach((id) => {
      batch.update(doc(db, CONTRACTS, id), {
        archived: true,
        archivedAt: serverTimestamp(),
      });
    });
    await batch.commit();
    done += Math.min(BATCH_SIZE, toArchive.length - i);
    onProgress?.(done, totalOps);
  }

  const totalOutstanding = records.reduce(
    (s, r) => s + (Number(r.balPrin) || 0),
    0
  );

  await addDoc(collection(db, UPLOADS), {
    date: serverTimestamp(),
    uploadedBy: uploadedBy || "unknown",
    totalRecords: records.length,
    newCount,
    archivedCount: toArchive.length,
    totalOutstanding,
  });

  return { total: records.length, newCount, archivedCount: toArchive.length };
}

// ---------- per-contract updates ----------
export async function setContractStatusFS(id, statusKey) {
  await updateDoc(doc(db, CONTRACTS, id), {
    status: statusKey,
    statusHistory: arrayUnion({
      status: statusKey,
      ts: new Date().toISOString(),
    }),
  });
}

export async function addNoteFS(id, text) {
  await updateDoc(doc(db, CONTRACTS, id), {
    notes: arrayUnion({
      id: Date.now(),
      ts: new Date().toISOString(),
      text,
    }),
  });
}

export async function saveWarnaFS(id, value) {
  await updateDoc(doc(db, CONTRACTS, id), { warnaKendaraan: value });
}
