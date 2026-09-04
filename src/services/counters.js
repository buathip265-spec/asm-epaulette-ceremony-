import { doc, runTransaction, serverTimestamp } from 'firebase/firestore';
import { db } from '../firebase/config.js';

const BADGE_COUNTER_REF = () => doc(db, 'counters', 'badgeSequence');

// Reserves `count` consecutive badge numbers and returns the first one.
// This is the ONLY place badge numbers are minted — once assigned, a
// badgeNumber is never recomputed or reassigned (see src/utils/thaiName.js
// for why the old "resort the whole roster" approach broke printed badges).
//
// Wrapped in its own small transaction (touches one document) so concurrent
// admins creating participants at the same time can never receive the same
// badge number, without needing every participant write to go through a
// single giant transaction.
export async function reserveBadgeNumbers(count, { reset = false } = {}) {
  if (count <= 0) return 1;
  return runTransaction(db, async (tx) => {
    const ref = BADGE_COUNTER_REF();
    const snap = await tx.get(ref);
    const current = reset ? 1 : (snap.exists() ? snap.data().nextBadgeNumber : 1);
    tx.set(ref, { nextBadgeNumber: current + count, updatedAt: serverTimestamp() }, { merge: true });
    return current;
  });
}
