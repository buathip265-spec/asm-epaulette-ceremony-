import {
  collection,
  deleteDoc,
  doc,
  getDoc,
  getDocs,
  limit,
  onSnapshot,
  orderBy,
  query,
  runTransaction,
  serverTimestamp,
  setDoc,
  startAfter,
  updateDoc,
  where,
  writeBatch,
} from 'firebase/firestore';
import { db } from '../firebase/config.js';
import { participantDocId } from '../utils/participantId.js';
import { reserveBadgeNumbers } from './counters.js';
import { logAuditEvent } from './audit.js';
import { snapshotParticipantsBackup } from './backup.js';

export class CheckInError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
  }
}

// The Firestore collection is still named "guests" (kept as-is per an
// explicit decision to avoid an unnecessary, riskier migration) — this
// module is where that fact lives; everything above this line in the app
// talks about "participants" in code and UI, and this is the one place
// that maps that concept onto the real, pre-existing collection name.
export const GUESTS_COLLECTION = 'guests';
const col = () => collection(db, GUESTS_COLLECTION);
const ref = (id) => doc(db, GUESTS_COLLECTION, id);

// ---------------------------------------------------------------------------
// Reads — every query here is deliberately SCOPED (a status filter, a year,
// a prefix, or a single document by id). Nothing in this file holds a live
// listener over the whole collection: at real-event scale that's the
// difference between one screen syncing a few hundred KB and every
// connected device re-downloading thousands of records on every reconnect.
// ---------------------------------------------------------------------------

// PUBLIC-safe: a single get() by student ID. Firestore rules only grant
// `get`, never `list`, to unauthenticated callers — so this is the only
// public lookup shape that exists, by design (see firestore.rules).
export async function getParticipantByStudentId(studentId) {
  const id = (studentId || '').trim();
  if (!id) return null;
  const snap = await getDoc(ref(id));
  return snap.exists() ? { id: snap.id, ...snap.data() } : null;
}

// STAFF-only listeners below (enforced by rules `list: if isStaff()`).
//
// Every query below that combines an equality filter with `orderBy()` on a
// DIFFERENT field needs an explicit composite index — Firestore's automatic
// single-field indexes do NOT cover that combination, only a bare equality
// filter, a bare range filter, or orderBy alone. Each composite this file
// needs is declared in firestore.indexes.json; deploy it with
// `firebase deploy --only firestore:indexes` before relying on these.

export function listenBadgeQueue(onChange, onError) {
  const q = query(col(), where('status', '==', 'checked_in'), orderBy('badgeNumber'), limit(200));
  return onSnapshot(q, (snap) => onChange(snap.docs.map((d) => ({ id: d.id, ...d.data() }))), onError);
}

export function listenCallQueue(onChange, onError) {
  const q = query(
    col(),
    where('arrived', '==', true),
    where('called', '==', false),
    where('skipped', '==', false),
    orderBy('badgeNumber'),
    limit(200)
  );
  return onSnapshot(q, (snap) => onChange(snap.docs.map((d) => ({ id: d.id, ...d.data() }))), onError);
}

export function listenByFlag(field, value, onChange, onError, { limitCount = 200 } = {}) {
  const q = query(col(), where(field, '==', value), orderBy('badgeNumber'), limit(limitCount));
  return onSnapshot(q, (snap) => onChange(snap.docs.map((d) => ({ id: d.id, ...d.data() }))), onError);
}

export function listenByYear(year, onChange, onError, { limitCount = 400 } = {}) {
  const constraints = [orderBy('badgeNumber'), limit(limitCount)];
  const q = year === 'all' ? query(col(), ...constraints) : query(col(), where('year', '==', year), ...constraints);
  return onSnapshot(q, (snap) => onChange(snap.docs.map((d) => ({ id: d.id, ...d.data() }))), onError);
}

export async function searchByNamePrefix(prefix, { limitCount = 25 } = {}) {
  const p = prefix.trim();
  if (!p) return [];
  const snap = await getDocs(
    query(col(), where('name', '>=', p), where('name', '<=', p + ''), limit(limitCount))
  );
  return snap.docs.map((d) => ({ id: d.id, ...d.data() }));
}

export async function searchByStudentIdPrefix(prefix, { limitCount = 25 } = {}) {
  const p = prefix.trim();
  if (!p) return [];
  const snap = await getDocs(
    query(col(), where('studentId', '>=', p), where('studentId', '<=', p + ''), limit(limitCount))
  );
  return snap.docs.map((d) => ({ id: d.id, ...d.data() }));
}

// Full-roster reads exist ONLY for the two rare, deliberate, ADMIN-gated
// operations that genuinely need everyone at once: exporting the report
// and taking a pre-destructive-operation backup (see services/backup.js).
// Neither is a listener and neither runs on a hot path.
export async function getAllParticipantsForExport() {
  const snap = await getDocs(query(col(), orderBy('badgeNumber')));
  return snap.docs.map((d) => ({ id: d.id, ...d.data() }));
}

// Admin table pagination — never loads the full roster at once.
export async function listParticipantsPage({ pageSize = 50, cursor = null, year = 'all' } = {}) {
  const constraints = [orderBy('badgeNumber'), limit(pageSize)];
  if (year !== 'all') constraints.unshift(where('year', '==', year));
  if (cursor) constraints.push(startAfter(cursor));
  const snap = await getDocs(query(col(), ...constraints));
  return {
    items: snap.docs.map((d) => ({ id: d.id, ...d.data() })),
    lastDoc: snap.docs.length > 0 ? snap.docs[snap.docs.length - 1] : null,
    hasMore: snap.docs.length === pageSize,
  };
}

// ---------------------------------------------------------------------------
// Writes
// ---------------------------------------------------------------------------

// The one and only check-in path. Runs as a transaction so that two staff
// devices racing to check in the same person can never both succeed: the
// second one re-reads inside the transaction, sees status is no longer
// 'pending', and rejects with ALREADY_CHECKED_IN before writing anything.
// The caller must only show a success state after this promise resolves —
// never optimistically before.
export async function checkInParticipant(participantId, actorUid) {
  const participantRef = ref(participantId);
  await runTransaction(db, async (tx) => {
    const snap = await tx.get(participantRef);
    if (!snap.exists()) throw new CheckInError('NOT_FOUND', 'ไม่พบข้อมูลผู้เข้าร่วมงานนี้');
    const data = snap.data();
    if (data.status !== 'pending') {
      throw new CheckInError('ALREADY_CHECKED_IN', 'ผู้เข้าร่วมคนนี้เช็กชื่อไปแล้ว');
    }
    tx.update(participantRef, {
      status: 'checked_in',
      arrived: true,
      checkInTime: serverTimestamp(),
      checkInBy: actorUid,
      updatedAt: serverTimestamp(),
    });
    tx.set(doc(collection(db, 'checkInEvents')), {
      participantId,
      action: 'check_in',
      actorUid,
      previousStatus: data.status,
      newStatus: 'checked_in',
      timestamp: serverTimestamp(),
    });
  });
}

export async function handoverBadge(participantId, actorUid) {
  const participantRef = ref(participantId);
  await runTransaction(db, async (tx) => {
    const snap = await tx.get(participantRef);
    if (!snap.exists()) throw new CheckInError('NOT_FOUND', 'ไม่พบข้อมูลผู้เข้าร่วมงานนี้');
    const data = snap.data();
    if (data.status !== 'checked_in') {
      throw new CheckInError('NOT_READY', 'ผู้เข้าร่วมคนนี้ยังไม่ได้เช็กชื่อ หรือรับป้ายไปแล้ว');
    }
    tx.update(participantRef, { status: 'completed', updatedAt: serverTimestamp() });
    tx.set(doc(collection(db, 'checkInEvents')), {
      participantId,
      action: 'handover',
      actorUid,
      previousStatus: data.status,
      newStatus: 'completed',
      timestamp: serverTimestamp(),
    });
  });
}

export async function setCalled(participantId, called, actorUid) {
  const patch = { called, updatedAt: serverTimestamp() };
  if (called) patch.skipped = false; // being called always clears "skipped"
  await updateDoc(ref(participantId), patch);
  if (called) {
    await setDoc(
      doc(db, 'stageState', 'current'),
      { participantId, calledAt: serverTimestamp(), calledBy: actorUid },
      { merge: true }
    );
  }
}

export async function setSkipped(participantId, skipped) {
  const patch = { skipped, updatedAt: serverTimestamp() };
  if (skipped) patch.called = false; // being skipped always clears "called"
  await updateDoc(ref(participantId), patch);
}

// ADMIN only (enforced by rules) — creates one participant with a freshly
// reserved, permanent badge number.
export async function createParticipant(fields, actorUid) {
  const badgeNumber = await reserveBadgeNumbers(1);
  const id = participantDocId(fields.studentId);
  await setDoc(ref(id), {
    studentId: (fields.studentId || '').trim(),
    name: fields.name.trim(),
    year: fields.year,
    note: (fields.note || '').trim(),
    badgeNumber,
    status: 'pending',
    arrived: false,
    checkInTime: null,
    checkInBy: null,
    called: false,
    skipped: false,
    createdAt: serverTimestamp(),
    createdBy: actorUid,
    updatedAt: serverTimestamp(),
  });
  return id;
}

// studentId is intentionally NOT editable here — it's the document id, and
// Firestore document ids can't be renamed in place. Changing a mistyped
// student ID means deleting and re-creating the record (an ADMIN action),
// not an in-place edit.
export async function updateParticipantDetails(id, { name, year, note }, actorUid) {
  await updateDoc(ref(id), {
    name: name.trim(),
    year,
    note: (note || '').trim(),
    updatedAt: serverTimestamp(),
  });
  await logAuditEvent({ actorUid, action: 'edit_participant', detail: { participantId: id } });
}

export async function deleteParticipant(id, actorUid) {
  await deleteDoc(ref(id));
  await logAuditEvent({ actorUid, action: 'delete_participant', detail: { participantId: id } });
}

// Resets every participant's operational state back to 'pending'. Snapshots
// a full backup first (see services/backup.js) so this is recoverable if
// it's triggered by mistake.
export async function resetAllStatuses(actorUid) {
  const backup = await snapshotParticipantsBackup({ actorUid, reason: 'reset_all_statuses' });

  const snap = await getDocs(col());
  const docs = snap.docs;
  const CHUNK_SIZE = 400;
  for (let i = 0; i < docs.length; i += CHUNK_SIZE) {
    const batch = writeBatch(db);
    docs.slice(i, i + CHUNK_SIZE).forEach((d) => {
      batch.update(d.ref, {
        status: 'pending',
        arrived: false,
        checkInTime: null,
        checkInBy: null,
        called: false,
        skipped: false,
        updatedAt: serverTimestamp(),
      });
    });
    await batch.commit();
  }
  await setDoc(doc(db, 'stageState', 'current'), { participantId: null, calledAt: null }, { merge: true });

  await logAuditEvent({
    actorUid,
    action: 'reset_all_statuses',
    detail: { participantCount: docs.length, backupSnapshotId: backup.snapshotId },
  });
  return { participantCount: docs.length, backupSnapshotId: backup.snapshotId };
}
