import { collection, doc, getDocs, setDoc, writeBatch, serverTimestamp } from 'firebase/firestore';
import { db } from '../firebase/config.js';
import { GUESTS_COLLECTION } from './participants.js';

const CHUNK_SIZE = 400; // stay comfortably under Firestore's 500-write batch limit

// Snapshots the entire current guests collection into
// backups/{snapshotId}/guests/* — mirroring the live collection name
// exactly, so restoring by hand is a plain "copy this subcollection back
// over guests/" with no name translation to get wrong — before a
// destructive operation (Excel replace-import, full status reset). This is
// the recovery path while staying on the Spark plan — no Cloud Storage
// scheduled export required. It costs one full collection read, which is
// acceptable because it only runs on a deliberate, rare, admin-confirmed
// action, never on a hot path or a live listener.
export async function snapshotParticipantsBackup({ actorUid, reason }) {
  const snap = await getDocs(collection(db, GUESTS_COLLECTION));
  const docs = snap.docs.map((d) => ({ id: d.id, data: d.data() }));

  const snapshotId = new Date().toISOString().replace(/[:.]/g, '-');
  const backupRef = doc(db, 'backups', snapshotId);
  await setDoc(backupRef, {
    createdAt: serverTimestamp(),
    createdBy: actorUid,
    reason,
    participantCount: docs.length,
  });

  for (let i = 0; i < docs.length; i += CHUNK_SIZE) {
    const batch = writeBatch(db);
    docs.slice(i, i + CHUNK_SIZE).forEach(({ id, data }) => {
      batch.set(doc(db, 'backups', snapshotId, GUESTS_COLLECTION, id), data);
    });
    await batch.commit();
  }

  return { snapshotId, participantCount: docs.length };
}
