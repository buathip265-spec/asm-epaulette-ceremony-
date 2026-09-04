# Project Report

**System:** Epaulette Ceremony Check-In Website (ระบบเช็คชื่อพิธีวันเกียรติยศ)
**Date:** 2026-09-04

---

## 1. System Overview

This website is used to manage check-in for the Epaulette Ceremony event. It replaces a paper sign-in sheet with a live, shared digital system that everyone on the team sees update in real time.

**Who uses it:**
- **Staff** — check participants in, hand out badges, and run the stage-calling queue during the event
- **Admin** — manage the participant list, import/export data, and handle setup before the event
- **Public / guests** — can look themselves up to check their queue status, and view the live stage display screen

**Main functions:** search and check in a participant, track badge hand-out, manage who gets called to the stage next, show a live "now calling" screen, and manage the participant list (add, edit, import from Excel, export a report).

---

## 2. Old Version vs New Version

| Topic | Old Version | New Version |
|---|---|---|
| **Login / Staff access** | One shared 4-digit PIN, typed into the app. The same PIN for every staff member, visible in the website's code. | Each staff member has their own personal login (email + password), created individually. No shared PIN. |
| **Check-in** | Any visitor could tap their own name on the kiosk screen and check themselves in — no staff involved. | Only a logged-in Staff member can search for a participant and confirm their check-in. Visitors can no longer check themselves in. |
| **Data security** | The database had no real protection underneath — the PIN only hid buttons in the app, but did not actually stop someone from reading or changing data directly. | Real protection rules are enforced at the database level. The public can only look up **one participant at a time**, by their exact student ID — never the full list. Only Staff/Admin accounts can see everyone or make changes. |
| **Participant data** | The **entire** participant list (names + student IDs) was sent to every visitor's device, even people who never logged in. | The full list is only ever sent to logged-in Staff/Admin. The public can only look up one specific person they already know the ID for. |
| **Badge numbers** | Recalculated every time the list changed. Adding or editing one person could shift everyone else's badge number — a problem if badges were already printed. | Assigned once and never change afterward. Safe to print badges ahead of time. |
| **Excel import** | No preview before importing. Could accidentally wipe the whole list with no way to undo. Didn't check for duplicate or broken rows. | Shows a preview first ("Found X people, Y invalid, Z duplicates") before anything is saved. Automatically backs up the current list before replacing it. Import is blocked until reviewed. |
| **Data backup** | None. Deleting or replacing data was permanent — no way to recover. | A backup snapshot is automatically saved before any action that replaces or resets participant data. |
| **Multiple staff devices** | If two staff checked in the same person at almost the same moment, there was no protection — could cause confusing results. | Built so that only **one** check-in can ever succeed if two staff try the same person at the same time. The second staff member is told "already checked in" instead of a confusing or duplicate result. |
| **Error handling** | If saving to the system failed (e.g. lost internet for a moment), the screen could still show "success" even though nothing was actually saved. | The screen only shows success **after** the system confirms the save worked. If something fails, staff sees a clear failure message and can try again — never a false "success." |
| **Code maintainability** | Nearly the entire website was written in one very large file, making it risky for anyone to change safely. | Reorganized into smaller, clearly separated pieces by function — safer to update in the future. |
| **Testing** | No automated checks existed at all. | 24 automated checks now cover the core logic, plus a dedicated test suite for the database's security rules. |

---

## 3. Current Features

- **Login** — Staff and Admin sign in with a personal email + password.
- **Staff check-in** — search for a participant, verify their name, and confirm check-in.
- **Participant search** — Staff can search by name or student ID and browse by year group. The public can only look up one exact student ID at a time (no browsing).
- **Badge hand-out queue** — shows everyone who has checked in and is waiting for their badge; staff mark it as handed over.
- **Stage-calling queue (MC)** — staff/MC can call a participant to the stage or skip them temporarily.
- **Public stage display** — a screen (e.g. an LED display) that shows who is currently being called, viewable by anyone with no login.
- **Public queue lookup** — anyone can check their own (or someone's) queue status if they know the exact student ID.
- **Excel import** — Admin can add participants in bulk from an Excel file, either adding to the existing list or replacing it entirely, with a review step first.
- **Excel export** — Admin can download a full report of everyone's status.
- **Admin participant management** — add, edit, or delete individual participants; reset everyone's check-in status back to "not arrived" (e.g. to rehearse).
- **QR codes** — quick way to open a specific page (e.g. the stage display, or the lookup page) on another device.

---

## 4. How Staff Use the System

1. Log in with your personal Staff account
2. Go to the **"เช็คชื่อ" (Check-in)** tab
3. Search for the participant by name or student ID
4. Verify the name shown matches the person in front of you
5. Confirm check-in — wait for the "เช็กชื่อสำเร็จ" success message before moving on
6. When it's time to hand out badges, go to the **"รับป้าย" (Badge)** tab, find the participant, hand them their badge, and mark it as handed over
7. If you're running the stage queue, use the **"ขึ้นเวที" (Call queue)** tab to call the next participant — this automatically updates the public stage display

**Important:** always wait for the on-screen confirmation. If the system shows an error instead of a success message, the check-in did **not** happen — try again rather than assuming it worked.

---

## 5. How Admin Uses the System

- **Managing participants** — add new people one at a time, edit details, or delete a record from the **"จัดการ" (Admin)** tab
- **Importing Excel** — upload a file, review the preview (how many will be added, how many have problems), choose whether to add to the existing list or replace it entirely, then confirm
- **Exporting data** — download a full Excel report of the current status of everyone
- **Resetting check-in status** — clears everyone back to "not arrived" (useful for a rehearsal run) — the system takes an automatic backup first
- **Managing Staff accounts** — this is **not** done inside the website. Creating a new Staff or Admin login is done manually through the Firebase Console (the setup tool used when this system was first configured) — see the checklist below.

---

## 6. Important Changes and Improvements

1. **Security** — Removed the shared PIN entirely. Staff/Admin access is now based on real personal logins, and the underlying database itself refuses unauthorized access — not just the visible screen.
2. **Check-in reliability** — Check-in only ever shows "success" after it's actually confirmed saved. No more silent failures that look like a success.
3. **Preventing duplicate/wrong check-ins** — If two staff members try to check in the same participant at the same moment, the system guarantees only one succeeds and clearly tells the second person it was already done.
4. **Excel safety** — Every import now shows a preview and checks for problems (missing names, duplicate IDs, bad formats) before saving anything, and a backup is taken automatically before replacing the list.
5. **Badge number stability** — Numbers are now permanent once assigned, so printed badges will always match the system.
6. **Handling many users** — The system no longer forces every device to download the entire participant list; each screen only loads the specific data it actually needs, which keeps things fast even with a large number of participants and many staff devices at once.
7. **Error handling** — Every screen now clearly shows one of: waiting, saving, success, or error — so staff always know the true state of what they just did.

---

## 7. Things the Team Needs to Be Careful About

- **Do not share Staff or Admin accounts** between multiple people — each person should have their own login.
- **Do not give Admin access unnecessarily** — most day-of-event staff only need a Staff account, not Admin.
- **Check the Excel preview carefully before confirming an import** — especially the "invalid rows" and "duplicates" counts.
- **Be extra careful with "Replace all"** when importing — it removes the current list (a backup is taken automatically, but it's still a big action; don't use it casually).
- **Make sure staff are actually logged in** before relying on the check-in tab — if a page seems to be missing options, it's likely a login issue, not a bug.
- **If the system shows an error message, assume the action did NOT happen** — never assume success just because you clicked the button. Try again.
- **Don't rely on memory for backups** — the system automatically backs up before replacing or resetting data, but this only covers those specific actions, not every single click.

---

## 8. Before the Event Checklist

### ✅ Already Completed
- [x] Firebase project created and connected to the website
- [x] Database access rules set up and deployed
- [x] Database performance settings (indexes) deployed and confirmed working
- [x] First Admin account created
- [x] Login tested for Staff, Admin, and public (no-login) access
- [x] Check-in, badge hand-out, and stage-calling queue tested end-to-end
- [x] Confirmed the public cannot browse the full participant list or make changes
- [x] Confirmed Staff cannot access Admin-only functions

### ⏳ Still Needs to Be Done
- [ ] **Create real Staff accounts** for everyone working the event (currently only 1 admin + 1 test staff account exist)
- [ ] **Delete or repurpose the test Staff account** used during setup testing
- [ ] **Remove the test participant data** (10 test entries currently in the system) — the easiest way is to import the real participant list using "Replace all," which clears test data automatically
- [ ] **Import the real participant Excel list**
- [ ] **Do a final test with real (or realistic) data** on the actual devices that will be used on event day (tablet for check-in, phone for badges, screen for stage display)
- [ ] **Confirm the event venue's internet/wifi** can reach the system before the event — test on-site if possible

---

## 9. Current Limitations

To be transparent about what has **not** yet been fully verified:

- **Not tested at real event scale.** The system has only been tested with 10 test participants and a couple of people using it at once — not with hundreds/thousands of participants or many staff devices simultaneously. It was *designed* to handle that scale, but this has not been proven with a real load test.
- **No self-service password reset.** If a staff member forgets their password, an Admin must reset it manually through the Firebase Console — there is no "forgot password" button in the app yet.
- **No in-app screen for managing Staff accounts.** Every new Staff/Admin account must currently be created manually through the Firebase Console, not through the website itself.
- **No barcode/QR scanning for check-in.** Looking up a participant during check-in is done by typing their name or student ID — there is no card/badge scanner built in.
- **The automated security-rule test suite has not been run in this environment**, though it has been written and reviewed. It requires additional local setup (a testing tool) that hasn't been confirmed on a team member's machine yet.

---

## 10. Final Status

🟡 **Needs attention**

The system itself is built, deployed, and has passed manual testing for logins, permissions, check-in, badge hand-out, and the stage queue. It is technically working.

**Before it can be used at the real event**, the team still needs to: create real Staff accounts, remove the test data, import the real participant list, and do one final run-through with real data on the actual devices that will be used on event day. None of these are code changes — they are setup/data steps the team can do directly.
