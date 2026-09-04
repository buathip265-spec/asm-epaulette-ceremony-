# ระบบเช็คชื่อพิธีวันเกียรติยศ (Epaulette Ceremony Check-In)

A React + Vite + Firebase check-in system: participant search, staff-verified
check-in, badge hand-out queue, MC call queue, a public LED stage display,
Excel import/export, and role-based admin tools.

This README covers what a small, non-web-specialist team needs to run,
deploy, and operate the system safely. For the full technical audit this
rework was based on (findings, architecture rationale, data model, and the
security-rules reasoning), see the audit document delivered alongside it.

## Roles

| Role | Can | Enforced by |
|---|---|---|
| **PUBLIC** | Look up one participant by exact student ID; view the LED stage display | Firestore rules — no login needed |
| **STAFF** | Search/browse participants, check them in, hand out badges, run the MC call queue | Firebase Auth + `staffUsers/{uid}.role` |
| **ADMIN** | Everything STAFF can, plus add/edit/delete participants, Excel import/export, reset check-in status | Firebase Auth + `staffUsers/{uid}.role == 'admin'` |

There is no self-service kiosk — every check-in is performed by a signed-in
staff member searching for and confirming the participant, not the
participant tapping their own name.

## Local setup

```bash
npm install
cp .env.example .env   # fill in Firebase config from the console, or ask
                        # whoever ran the initial setup for the values
npm run dev
```

## Before the first real use: bootstrap one admin account

Nothing in the app can create the first admin — this is the one manual step,
done once, directly in the Firebase console:

1. **Authentication → Users → Add user** — create the admin's email/password
   account. Note the generated UID.
2. **Firestore → staffUsers collection → Add document** — document ID = that
   UID, with one field: `role` (string) = `admin`.

From then on, that admin can add more `staffUsers` documents through the
app's own tooling if you build a management screen for it later — for now,
every additional staff account is created the same way: an account in
Authentication, then a `staffUsers/{uid}` document with `role: 'staff'` or
`role: 'admin'`.

## Deploying Firestore rules & indexes

```bash
npm install -g firebase-tools   # once
firebase login
firebase deploy --only firestore:rules,firestore:indexes
```

Do this **before** the app goes live with real data — the rules are what
actually protects the roster; nothing in the client enforces access.

## Migrating existing data (one time, if the `guests` collection already has data)

If there's already data in the `guests` collection from before this rework,
run the migration script once — it re-keys documents by student ID, assigns
permanent badge numbers, and seeds the badge counter. **Read the comment
block at the top of `scripts/migrate-badge-numbers.mjs` in full before
running it** — it explains exactly what it does and the safety checks it
runs first.

```bash
# .env needs MIGRATION_ADMIN_EMAIL / MIGRATION_ADMIN_PASSWORD added
# (the bootstrap admin account's credentials) on top of the usual
# VITE_FIREBASE_* values.
npm run migrate:badges
```

Run this against a **staging** Firebase project first if you have one.
There is no undo button for a migration script — that's why it refuses to
proceed at all if it finds duplicate student IDs or an id collision, rather
than guessing.

## Testing

```bash
npm run test         # unit tests — pure logic, no Firebase needed
npm run test:rules    # Firestore Security Rules test suite — needs the
                       # Firebase CLI and a JRE installed locally, since it
                       # runs against the local emulator
```

## Building & deploying the app itself

```bash
npm run build
firebase deploy --only hosting   # if using Firebase Hosting — firebase.json
                                  # is already configured for it
```

Any static host works, not just Firebase Hosting — `npm run build` produces
a plain `dist/` folder.

## Production checklist (do this before a real event, not after)

- [ ] `.env` filled in with the **production** Firebase project's config
- [ ] `firestore.rules` deployed and the rules test suite passing
- [ ] Firestore composite index deployed (`firestore.indexes.json`) —
      Firestore will otherwise reject the MC/stage queries with a "missing
      index" error the first time they run
- [ ] At least one admin account bootstrapped (see above)
- [ ] All staff accounts created ahead of the event — not during it
- [ ] Real participant data imported via **Admin → นำเข้ารายชื่อจาก Excel**
      (review the preview counts before confirming)
- [ ] A test check-in performed end-to-end on the actual devices that will
      be used (tablet for check-in, phone for badge queue, whatever drives
      the stage display)
- [ ] Confirm the venue network can reach `firestore.googleapis.com` — if
      it's a locked-down guest network, test this *before* event day
- [ ] `npm run test` passing
- [ ] `npm run test:rules` passing

## Project structure

```
src/
  firebase/       Firebase app/auth/db initialization
  services/       All Firestore reads/writes — the only layer that talks
                  to Firestore directly
  features/       One folder per screen (checkin, badges, mc, stage,
                  publicLookup, admin, auth)
  components/     Shared UI (modals, toasts, nav, status pills)
  hooks/          Small reusable hooks (online status, sound preference)
  utils/          Pure functions — no Firebase, no DOM (unit-tested)
scripts/          One-time operational scripts (badge/id migration)
tests/            Firestore Security Rules tests (needs the emulator)
firestore.rules            The actual access-control boundary
firestore.indexes.json     Composite indexes the queries need
firebase.json               Firebase CLI project config
```
