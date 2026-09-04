#!/usr/bin/env node
// ============================================================================
// One-time migration: run ONCE, against the live "guests" collection, before
// switching the app over to the new rules/data model. Safe to re-run (it's
// idempotent for docs it has already migrated), but it is NOT a routine
// script — read this whole comment block before running it.
//
// What it does, and why each part is needed:
//
// 1. Re-keys every guest document so its Firestore document ID equals its
//    studentId. The old app assigned arbitrary ids ('h01', 'guest_171...'),
//    but the new security rules give the public a `get()`-only lookup by
//    student ID (see firestore.rules) — that only works if the id IS the
//    student ID. A doc with no studentId keeps a generated id, unchanged.
//
// 2. Assigns a PERMANENT badgeNumber to every guest, computed ONCE here
//    using the exact same sort (year, then Thai name) the old app used to
//    compute badge numbers live on every render. After this script runs,
//    badgeNumber is frozen data — nothing recomputes it from a sort again.
//
// 3. Seeds counters/badgeSequence so the app's badge-numbering continues
//    from the right place for anyone added after this migration.
//
// 4. Adds the fields the new code expects that don't exist on old docs yet:
//    arrived, createdAt, createdBy, updatedAt.
//
// SAFETY: before changing anything, this script checks for two studentId
// collisions and REFUSES to proceed if either is1 found, printing exactly
// what to fix by hand first:
//   - two existing docs sharing the same non-empty studentId (the old app
//     never prevented this; the new one structurally can't allow it)
//   - a collision between what a re-keyed doc's new id would be and an
//     existing doc's, current id
//
// USAGE
//   1. Deploy firestore.rules first (this script authenticates as an admin
//      and writes through the SAME rules the app uses — see firestore.rules
//      for why counters/create is deliberately rules-denied outside this
//      script's one-time seed below).
//   2. Create the first admin account by hand in the Firebase console
//      (Authentication tab), then create its staffUsers/{uid} doc with
//      { role: 'admin' } by hand too (also via the console — this is the
//      one bootstrap step nothing in the app can do for itself).
//   3. Set MIGRATION_ADMIN_EMAIL / MIGRATION_ADMIN_PASSWORD env vars to
//      that account's credentials (plus the usual VITE_FIREBASE_* vars —
//      this script reads the same .env as the app).
//   4. Run against a STAGING project first: npm run migrate:badges
//   5. Review the printed summary. Only then run it against production.
// ============================================================================

import { initializeApp } from 'firebase/app';
import { getAuth, signInWithEmailAndPassword } from 'firebase/auth';
import { collection, doc, getDocs, getFirestore, serverTimestamp, setDoc, writeBatch } from 'firebase/firestore';

// Env vars come from `node --env-file=.env` (see the "migrate:badges" script
// in package.json) — Node 20.6+'s built-in loader, deliberately not the
// `dotenv` package, so this one script doesn't add a new dependency for
// something the runtime already does.

const GUESTS_COLLECTION = 'guests';

const YEAR_WEIGHTS = {
  'ปี 1': 1,
  'ปี 2': 2,
  'ปี 3': 3,
  'ปี 4': 4,
  บัณฑิต: 5,
  อาจารย์: 6,
  แขกผู้มีเกียรติ: 7,
};
function getYearOrderWeight(yearStr) {
  if (!yearStr) return 99;
  for (const [key, weight] of Object.entries(YEAR_WEIGHTS)) {
    if (yearStr.includes(key)) return weight;
  }
  return 50;
}
const TITLE_PREFIX = /^(นาย|นางสาว|นาง|ด\.ช\.|ด\.ญ\.|ผศ\.|รศ\.|ดร\.)\s*/;
function getSortableName(fullName) {
  if (!fullName) return '';
  return fullName.replace(TITLE_PREFIX, '').trim();
}
function participantDocId(studentId) {
  const trimmed = (studentId || '').trim();
  return trimmed || null; // null => keep the existing generated id, unchanged
}

async function main() {
  const required = [
    'VITE_FIREBASE_API_KEY',
    'VITE_FIREBASE_AUTH_DOMAIN',
    'VITE_FIREBASE_PROJECT_ID',
    'VITE_FIREBASE_APP_ID',
    'MIGRATION_ADMIN_EMAIL',
    'MIGRATION_ADMIN_PASSWORD',
  ];
  const missing = required.filter((k) => !process.env[k]);
  if (missing.length > 0) {
    console.error(`Missing required env vars: ${missing.join(', ')}`);
    process.exit(1);
  }

  const app = initializeApp({
    apiKey: process.env.VITE_FIREBASE_API_KEY,
    authDomain: process.env.VITE_FIREBASE_AUTH_DOMAIN,
    projectId: process.env.VITE_FIREBASE_PROJECT_ID,
    appId: process.env.VITE_FIREBASE_APP_ID,
  });
  const auth = getAuth(app);
  const db = getFirestore(app);

  console.log(`Signing in as ${process.env.MIGRATION_ADMIN_EMAIL}...`);
  await signInWithEmailAndPassword(auth, process.env.MIGRATION_ADMIN_EMAIL, process.env.MIGRATION_ADMIN_PASSWORD);

  console.log(`Reading ${GUESTS_COLLECTION}/ ...`);
  const snap = await getDocs(collection(db, GUESTS_COLLECTION));
  const existing = snap.docs.map((d) => ({ id: d.id, data: d.data() }));
  console.log(`Found ${existing.length} existing documents.`);

  // ---- safety check 1: duplicate studentId among existing docs ----------
  const byStudentId = new Map();
  existing.forEach((doc_) => {
    const sid = (doc_.data.studentId || '').trim();
    if (!sid) return;
    byStudentId.set(sid, [...(byStudentId.get(sid) || []), doc_.id]);
  });
  const dupes = Array.from(byStudentId.entries()).filter(([, ids]) => ids.length > 1);
  if (dupes.length > 0) {
    console.error('\nABORTING — duplicate studentId found among EXISTING documents:');
    dupes.forEach(([sid, ids]) => console.error(`  studentId ${sid}: documents ${ids.join(', ')}`));
    console.error('\nResolve these by hand (merge or delete the extras) before re-running.');
    process.exit(1);
  }

  // ---- compute the permanent sort/badge order, exactly as the old app did
  const sorted = [...existing].sort((a, b) => {
    const wa = getYearOrderWeight(a.data.year);
    const wb = getYearOrderWeight(b.data.year);
    if (wa !== wb) return wa - wb;
    return getSortableName(a.data.name).localeCompare(getSortableName(b.data.name), 'th');
  });

  // ---- safety check 2: a re-keyed id colliding with a DIFFERENT doc's current id
  const targetIds = sorted.map((doc_) => participantDocId(doc_.data.studentId) || doc_.id);
  const currentIdSet = new Set(existing.map((doc_) => doc_.id));
  const collisions = targetIds.filter((tid, i) => {
    const sourceId = sorted[i].id;
    return tid !== sourceId && currentIdSet.has(tid);
  });
  if (collisions.length > 0) {
    console.error('\nABORTING — re-keying would collide with an existing document id:');
    collisions.forEach((c) => console.error(`  target id ${c} already exists as a different document`));
    process.exit(1);
  }

  console.log('\nNo collisions found. Proceeding with migration...\n');

  const CHUNK = 400;
  let rekeyed = 0;
  let updatedInPlace = 0;

  for (let i = 0; i < sorted.length; i += CHUNK) {
    const chunk = sorted.slice(i, i + CHUNK);
    const batch = writeBatch(db);
    for (let j = 0; j < chunk.length; j++) {
      const item = chunk[j];
      const globalIndex = i + j;
      const badgeNumber = globalIndex + 1;
      const targetId = participantDocId(item.data.studentId) || item.id;

      const nextData = {
        ...item.data,
        badgeNumber,
        arrived: item.data.status === 'checked_in' || item.data.status === 'completed',
        checkInTime: item.data.checkInTime ?? null,
        checkInBy: item.data.checkInBy ?? null,
        createdAt: item.data.createdAt ?? serverTimestamp(),
        createdBy: item.data.createdBy ?? 'migration',
        updatedAt: serverTimestamp(),
      };

      if (targetId !== item.id) {
        batch.set(doc(db, GUESTS_COLLECTION, targetId), nextData);
        batch.delete(doc(db, GUESTS_COLLECTION, item.id));
        rekeyed++;
      } else {
        batch.set(doc(db, GUESTS_COLLECTION, item.id), nextData);
        updatedInPlace++;
      }
    }
    await batch.commit();
    console.log(`  committed ${Math.min(i + CHUNK, sorted.length)}/${sorted.length}`);
  }

  await setDoc(doc(db, 'counters', 'badgeSequence'), {
    nextBadgeNumber: sorted.length + 1,
    updatedAt: serverTimestamp(),
  });

  console.log('\nDone.');
  console.log(`  Re-keyed to studentId (delete + recreate): ${rekeyed}`);
  console.log(`  Updated in place (id unchanged):            ${updatedInPlace}`);
  console.log(`  Badge numbers assigned:                     1..${sorted.length}`);
  console.log(`  counters/badgeSequence.nextBadgeNumber set to ${sorted.length + 1}`);
  process.exit(0);
}

main().catch((err) => {
  console.error('Migration failed:', err);
  process.exit(1);
});
