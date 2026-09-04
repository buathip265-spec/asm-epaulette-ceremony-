import { addDoc, collection, serverTimestamp } from 'firebase/firestore';
import { db } from '../firebase/config.js';

// A single append-only trail for admin-level actions (import, delete,
// bulk reset, role changes) — separate from checkInEvents, which tracks
// per-participant status transitions. Both collections are write-once by
// rule (see firestore.rules): nothing, including an admin, can edit or
// delete an entry after the fact.
export async function logAuditEvent({ actorUid, action, detail = {} }) {
  await addDoc(collection(db, 'auditLogs'), {
    actorUid,
    action,
    detail,
    timestamp: serverTimestamp(),
  });
}
