import * as XLSX from 'xlsx';
import { collection, doc, getDocs, writeBatch, serverTimestamp } from 'firebase/firestore';
import { db } from '../firebase/config.js';
import { extractParticipantRows, validateParticipantRows } from '../utils/excelValidation.js';
import { sortByYearThenName } from '../utils/thaiName.js';
import { participantDocId } from '../utils/participantId.js';
import { reserveBadgeNumbers } from './counters.js';
import { snapshotParticipantsBackup } from './backup.js';
import { logAuditEvent } from './audit.js';
import { GUESTS_COLLECTION } from './participants.js';

const CHUNK_SIZE = 400;

// Reads a File (from an <input type="file">) and returns SheetJS's raw
// array-of-objects for the first sheet. This is the only place FileReader
// is used — everything downstream (validation, mapping) is plain data.
export function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('ไม่สามารถอ่านไฟล์ได้'));
    reader.onload = (evt) => {
      try {
        const workbook = XLSX.read(evt.target.result, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rawJson = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        resolve(extractParticipantRows(rawJson));
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsBinaryString(file);
  });
}

export function downloadSampleExcel() {
  const sampleRows = [
    { ชั้นปี: 'ปี 1', รหัสนักศึกษา: '69014522', 'ชื่อ-นามสกุล': 'นายกิตติกร บุญมี (กิต)', หมายเหตุ: '' },
    { ชั้นปี: 'ปี 2', รหัสนักศึกษา: '68023567', 'ชื่อ-นามสกุล': 'นางสาวจิราภรณ์ ทัดศรี (เจอาร์)', หมายเหตุ: '' },
    { ชั้นปี: 'ปี 3', รหัสนักศึกษา: '67037256', 'ชื่อ-นามสกุล': 'นางสาววิมลรัตน์ บุญชู (คิบิ)', หมายเหตุ: 'สโมสรนักศึกษา' },
    { ชั้นปี: 'ปี 4', รหัสนักศึกษา: '66045914', 'ชื่อ-นามสกุล': 'นางสาวบัวทิพย์ วัฒนเกษมสกุล (ทิพย์)', หมายเหตุ: 'สตาฟฝ่ายพิธีการ' },
  ];
  const worksheet = XLSX.utils.json_to_sheet(sampleRows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'รายชื่อผู้เข้าร่วม');
  XLSX.writeFile(workbook, 'ตัวอย่างไฟล์รายชื่อ_พิธีวันเกียรติยศ_69-66.xlsx');
}

function statusLabel(g) {
  if (g.status === 'completed') return 'รับป้ายชื่อแล้ว';
  if (g.status === 'checked_in') return 'เช็คชื่อแล้ว (รอรับป้าย)';
  return 'ยังไม่มารายงานตัว';
}

function stageLabel(g) {
  if (g.called) return 'ขานชื่อประดับบ่าแล้ว';
  if (g.skipped) return 'ข้ามคิว/สแตนด์บาย';
  return g.status === 'pending' ? 'ยังไม่มา' : 'รอขึ้นเวที';
}

// Export is ADMIN-only (enforced in the UI layer + not reachable without
// STAFF/ADMIN auth to read the roster at all, per firestore.rules `list`).
export function downloadSummaryReport(participants) {
  const rows = participants.map((g) => ({
    ลำดับป้าย: g.badgeNumber,
    ชั้นปี: g.year,
    รหัสนักศึกษา: g.studentId || '-',
    'ชื่อ-นามสกุล': g.name,
    สถานะการเช็คชื่อ: statusLabel(g),
    เวลาที่เช็คชื่อ: g.checkInTime ? g.checkInTime.toDate().toLocaleString('th-TH') : '-',
    สถานะบนเวที: stageLabel(g),
    หมายเหตุ: g.note || '',
  }));
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'รายงานพิธีวันเกียรติยศ');
  XLSX.writeFile(workbook, `รายงานสรุปพิธีวันเกียรติยศ_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// Builds the confirmation preview the admin sees before anything is
// written — "Found N participants. M invalid rows. K duplicates in file."
// plus, in 'append' mode specifically, rows that collide with a student ID
// already in the LIVE roster (which validateParticipantRows alone can't
// know about, since it never touches Firestore). A collision here is not
// safe to import: participants are keyed by student ID, so writing it
// would silently overwrite — and un-check-in — an existing record.
export async function buildImportPreview(rows, mode) {
  const base = validateParticipantRows(rows);
  if (mode !== 'append') {
    return { ...base, existingCollisions: [] };
  }

  const existingSnap = await getDocs(collection(db, GUESTS_COLLECTION));
  const existingIds = new Set(existingSnap.docs.map((d) => d.id));

  const collisions = base.validRows.filter((r) => r.studentId && existingIds.has(r.studentId));
  const collisionRows = new Set(collisions.map((r) => r.sourceRow));
  const validRows = base.validRows.filter((r) => !collisionRows.has(r.sourceRow));

  return {
    ...base,
    validRows,
    existingCollisions: collisions,
    summary: {
      ...base.summary,
      valid: validRows.length,
      existingCollisions: collisions.length,
    },
  };
}

// The actual import write path. Only ever called after the admin has seen
// and confirmed the validation preview above — this function assumes
// `validRows` has already been reviewed and excludes anything invalid,
// duplicate-in-file, or colliding with the existing roster.
//
// Sequencing matters here: for 'replace' mode we snapshot a backup FIRST,
// then delete, then insert — never delete before the new data exists
// somewhere recoverable. If a network drop happens mid-import, the worst
// case is a partially-applied import, not a wiped, unrecoverable roster.
export async function importParticipants({ validRows, mode, actorUid }) {
  if (mode === 'replace') {
    await snapshotParticipantsBackup({ actorUid, reason: 'excel_import_replace' });
  }

  const sorted = sortByYearThenName(validRows);
  const startingBadge = await reserveBadgeNumbers(sorted.length, { reset: mode === 'replace' });
  const newDocs = sorted.map((row, index) => ({
    id: participantDocId(row.studentId),
    data: {
      studentId: row.studentId,
      name: row.name,
      year: row.year,
      note: row.note || '',
      badgeNumber: startingBadge + index,
      status: 'pending',
      arrived: false,
      checkInTime: null,
      checkInBy: null,
      called: false,
      skipped: false,
      createdAt: serverTimestamp(),
      createdBy: actorUid,
      updatedAt: serverTimestamp(),
    },
  }));

  if (mode === 'replace') {
    const existing = await getDocs(collection(db, GUESTS_COLLECTION));
    for (let i = 0; i < existing.docs.length; i += CHUNK_SIZE) {
      const batch = writeBatch(db);
      existing.docs.slice(i, i + CHUNK_SIZE).forEach((d) => batch.delete(d.ref));
      await batch.commit();
    }
  }

  for (let i = 0; i < newDocs.length; i += CHUNK_SIZE) {
    const batch = writeBatch(db);
    newDocs.slice(i, i + CHUNK_SIZE).forEach(({ id, data }) => batch.set(doc(db, GUESTS_COLLECTION, id), data));
    await batch.commit();
  }

  await logAuditEvent({
    actorUid,
    action: mode === 'replace' ? 'import_replace' : 'import_append',
    detail: { importedCount: newDocs.length },
  });

  return { importedCount: newDocs.length };
}
