// Firestore Security Rules test suite.
//
// This is the single highest-value test in the project — it's the actual
// access-control boundary (see the audit's findings C-1/C-2). It runs
// against the LOCAL Firestore emulator, never a real project, and needs
// the Firebase CLI + a JRE installed. Run it with:
//
//   npm run test:rules
//
// which wraps `firebase emulators:exec` around a normal vitest run so the
// emulator starts, the tests run against it, and it shuts down again
// automatically — nothing here touches production data.
import { readFileSync } from 'node:fs';
import { afterAll, beforeAll, beforeEach, describe, it } from 'vitest';
import {
  assertFails,
  assertSucceeds,
  initializeTestEnvironment,
} from '@firebase/rules-unit-testing';
import { doc, getDoc, getDocs, collection, setDoc, updateDoc, deleteDoc } from 'firebase/firestore';

const PROJECT_ID = 'ceremony-rules-test';
let testEnv;

const seedGuest = async (id, data) => {
  await testEnv.withSecurityRulesDisabled(async (ctx) => {
    await setDoc(doc(ctx.firestore(), 'guests', id), {
      studentId: id,
      name: 'ทดสอบ ระบบ',
      year: 'ปี 1',
      note: '',
      badgeNumber: 1,
      status: 'pending',
      arrived: false,
      checkInTime: null,
      checkInBy: null,
      called: false,
      skipped: false,
      ...data,
    });
  });
};

const seedStaffUser = async (uid, role) => {
  await testEnv.withSecurityRulesDisabled(async (ctx) => {
    await setDoc(doc(ctx.firestore(), 'staffUsers', uid), { role });
  });
};

beforeAll(async () => {
  testEnv = await initializeTestEnvironment({
    projectId: PROJECT_ID,
    firestore: { rules: readFileSync('firestore.rules', 'utf8') },
  });
});

afterAll(async () => {
  await testEnv.cleanup();
});

beforeEach(async () => {
  await testEnv.clearFirestore();
});

describe('PUBLIC access', () => {
  it('can get() a single guest by id', async () => {
    await seedGuest('69014522', {});
    const publicDb = testEnv.unauthenticatedContext().firestore();
    await assertSucceeds(getDoc(doc(publicDb, 'guests', '69014522')));
  });

  it('cannot list() the guests collection', async () => {
    await seedGuest('69014522', {});
    const publicDb = testEnv.unauthenticatedContext().firestore();
    await assertFails(getDocs(collection(publicDb, 'guests')));
  });

  it('cannot write to a guest document', async () => {
    await seedGuest('69014522', {});
    const publicDb = testEnv.unauthenticatedContext().firestore();
    await assertFails(updateDoc(doc(publicDb, 'guests', '69014522'), { status: 'checked_in' }));
  });
});

describe('STAFF access', () => {
  it('can list() the roster', async () => {
    await seedGuest('69014522', {});
    await seedStaffUser('staff-1', 'staff');
    const staffDb = testEnv.authenticatedContext('staff-1').firestore();
    await assertSucceeds(getDocs(collection(staffDb, 'guests')));
  });

  it('can transition a guest from pending to checked_in', async () => {
    await seedGuest('69014522', { status: 'pending' });
    await seedStaffUser('staff-1', 'staff');
    const staffDb = testEnv.authenticatedContext('staff-1').firestore();
    await assertSucceeds(
      updateDoc(doc(staffDb, 'guests', '69014522'), { status: 'checked_in', arrived: true, updatedAt: new Date() })
    );
  });

  it('cannot set status to an arbitrary, non-operational value alongside identity fields', async () => {
    await seedGuest('69014522', { status: 'pending' });
    await seedStaffUser('staff-1', 'staff');
    const staffDb = testEnv.authenticatedContext('staff-1').firestore();
    // Attempting to rename the participant through a "status update" —
    // touches a frozen field (name), which onlyOperationalFieldsChanged()
    // must reject regardless of what else is in the payload.
    await assertFails(updateDoc(doc(staffDb, 'guests', '69014522'), { status: 'checked_in', name: 'ชื่อปลอม' }));
  });

  it('cannot change badgeNumber or studentId', async () => {
    await seedGuest('69014522', {});
    await seedStaffUser('staff-1', 'staff');
    const staffDb = testEnv.authenticatedContext('staff-1').firestore();
    await assertFails(updateDoc(doc(staffDb, 'guests', '69014522'), { badgeNumber: 999 }));
    await assertFails(updateDoc(doc(staffDb, 'guests', '69014522'), { studentId: '00000000' }));
  });

  it('cannot delete a guest', async () => {
    await seedGuest('69014522', {});
    await seedStaffUser('staff-1', 'staff');
    const staffDb = testEnv.authenticatedContext('staff-1').firestore();
    await assertFails(deleteDoc(doc(staffDb, 'guests', '69014522')));
  });

  it('cannot create a new guest', async () => {
    await seedStaffUser('staff-1', 'staff');
    const staffDb = testEnv.authenticatedContext('staff-1').firestore();
    await assertFails(setDoc(doc(staffDb, 'guests', '69099999'), { studentId: '69099999', name: 'ใหม่', year: 'ปี 1', badgeNumber: 2, status: 'pending' }));
  });
});

describe('a signed-in user with no staffUsers document', () => {
  it('is treated as PUBLIC, not STAFF — fails closed', async () => {
    await seedGuest('69014522', {});
    // Signed in (has a uid), but never given a staffUsers doc.
    const noRoleDb = testEnv.authenticatedContext('random-visitor').firestore();
    await assertFails(getDocs(collection(noRoleDb, 'guests')));
    await assertFails(updateDoc(doc(noRoleDb, 'guests', '69014522'), { status: 'checked_in' }));
  });
});

describe('ADMIN access', () => {
  it('can create, edit, and delete guests', async () => {
    await seedStaffUser('admin-1', 'admin');
    const adminDb = testEnv.authenticatedContext('admin-1').firestore();
    await assertSucceeds(
      setDoc(doc(adminDb, 'guests', '69099999'), { studentId: '69099999', name: 'ใหม่', year: 'ปี 1', badgeNumber: 2, status: 'pending' })
    );
    await assertSucceeds(updateDoc(doc(adminDb, 'guests', '69099999'), { name: 'แก้ไขแล้ว' }));
    await assertSucceeds(deleteDoc(doc(adminDb, 'guests', '69099999')));
  });

  it('cannot modify their own staffUsers document (self-privilege-escalation guard)', async () => {
    await seedStaffUser('admin-1', 'admin');
    const adminDb = testEnv.authenticatedContext('admin-1').firestore();
    await assertFails(updateDoc(doc(adminDb, 'staffUsers', 'admin-1'), { role: 'staff' }));
  });

  it('can modify a DIFFERENT staff member\'s role', async () => {
    await seedStaffUser('admin-1', 'admin');
    await seedStaffUser('staff-2', 'staff');
    const adminDb = testEnv.authenticatedContext('admin-1').firestore();
    await assertSucceeds(updateDoc(doc(adminDb, 'staffUsers', 'staff-2'), { role: 'admin' }));
  });
});

describe('counters/badgeSequence bootstrap', () => {
  it('ADMIN can create it from scratch (a brand-new project with nothing to migrate)', async () => {
    await seedStaffUser('admin-1', 'admin');
    const adminDb = testEnv.authenticatedContext('admin-1').firestore();
    await assertSucceeds(setDoc(doc(adminDb, 'counters', 'badgeSequence'), { nextBadgeNumber: 1 }));
  });

  it('STAFF cannot create or update it', async () => {
    await seedStaffUser('staff-1', 'staff');
    const staffDb = testEnv.authenticatedContext('staff-1').firestore();
    await assertFails(setDoc(doc(staffDb, 'counters', 'badgeSequence'), { nextBadgeNumber: 1 }));
  });
});

describe('checkInEvents (append-only audit trail)', () => {
  it('STAFF can create an entry', async () => {
    await seedStaffUser('staff-1', 'staff');
    const staffDb = testEnv.authenticatedContext('staff-1').firestore();
    await assertSucceeds(
      setDoc(doc(collection(staffDb, 'checkInEvents')), { participantId: '69014522', action: 'check_in', actorUid: 'staff-1' })
    );
  });

  it('cannot be updated or deleted by anyone, including ADMIN', async () => {
    await seedStaffUser('admin-1', 'admin');
    let eventId;
    await testEnv.withSecurityRulesDisabled(async (ctx) => {
      const ref = doc(collection(ctx.firestore(), 'checkInEvents'));
      await setDoc(ref, { participantId: '69014522', action: 'check_in' });
      eventId = ref.id;
    });
    const adminDb = testEnv.authenticatedContext('admin-1').firestore();
    await assertFails(updateDoc(doc(adminDb, 'checkInEvents', eventId), { action: 'tampered' }));
    await assertFails(deleteDoc(doc(adminDb, 'checkInEvents', eventId)));
  });
});
