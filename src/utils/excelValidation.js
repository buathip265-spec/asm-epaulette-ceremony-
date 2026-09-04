import { detectYearFromStudentId } from './yearDetect.js';

const HEADER_ALIASES = {
  year: ['ชั้นปี', 'ปี', 'year', 'ระดับ'],
  studentId: ['รหัสนักศึกษา', 'รหัสประจำตัว', 'รหัส', 'student_id', 'id'],
  name: ['ชื่อ-นามสกุล', 'ชื่อ นามสกุล', 'ชื่อสกุล', 'ชื่อ', 'name'],
  note: ['หมายเหตุ', 'note', 'remark', 'ตำแหน่ง'],
};

function findValue(row, keys, aliases) {
  const key = keys.find((k) => aliases.some((alias) => k.toLowerCase().trim().includes(alias.toLowerCase())));
  return key ? String(row[key]).trim() : '';
}

// Pure function: takes the raw array-of-objects SheetJS produces and turns
// it into participant-shaped rows, with no Firestore or File I/O — this is
// what makes it unit-testable without a browser or a live project.
export function extractParticipantRows(rawJson) {
  return rawJson.map((row, index) => {
    const keys = Object.keys(row);
    const rawYear = findValue(row, keys, HEADER_ALIASES.year);
    const studentId = findValue(row, keys, HEADER_ALIASES.studentId);
    const name = findValue(row, keys, HEADER_ALIASES.name);
    const note = findValue(row, keys, HEADER_ALIASES.note);
    const autoYear = detectYearFromStudentId(studentId);

    return {
      sourceRow: index + 2, // +1 for header row, +1 for 1-based row numbers
      year: rawYear || autoYear || '',
      studentId,
      name,
      note,
    };
  });
}

const VALID_STUDENT_ID = /^\d{6,10}$/;

// Validates the parsed rows and produces the counts the admin sees in the
// import preview ("Found N participants. M invalid rows. K duplicates.")
// before anything is written. Nothing here touches Firestore — import is
// only allowed to proceed once the caller has reviewed this summary.
export function validateParticipantRows(rows) {
  const errors = [];
  const seenStudentIds = new Map(); // studentId -> [sourceRow, ...]

  rows.forEach((row) => {
    const rowErrors = [];
    if (!row.name) rowErrors.push('ไม่มีชื่อ-นามสกุล');
    if (!row.year) rowErrors.push('ไม่สามารถระบุชั้นปีได้ (โปรดระบุคอลัมน์ชั้นปี)');
    if (row.studentId && !VALID_STUDENT_ID.test(row.studentId)) {
      rowErrors.push('รูปแบบรหัสนักศึกษาไม่ถูกต้อง (ต้องเป็นตัวเลข 6-10 หลัก)');
    }
    if (rowErrors.length > 0) {
      errors.push({ sourceRow: row.sourceRow, name: row.name, studentId: row.studentId, reasons: rowErrors });
    }
    if (row.studentId) {
      const existing = seenStudentIds.get(row.studentId) || [];
      seenStudentIds.set(row.studentId, [...existing, row.sourceRow]);
    }
  });

  const duplicates = Array.from(seenStudentIds.entries())
    .filter(([, sourceRows]) => sourceRows.length > 1)
    .map(([studentId, sourceRows]) => ({ studentId, sourceRows }));

  const invalidSourceRows = new Set(errors.map((e) => e.sourceRow));
  const duplicateSourceRows = new Set(duplicates.flatMap((d) => d.sourceRows.slice(1))); // keep first occurrence
  const rejectedSourceRows = new Set([...invalidSourceRows, ...duplicateSourceRows]);

  const validRows = rows.filter((r) => !rejectedSourceRows.has(r.sourceRow));

  return {
    totalRows: rows.length,
    validRows,
    invalidRows: errors,
    duplicates,
    summary: {
      total: rows.length,
      valid: validRows.length,
      invalid: errors.length,
      duplicateStudentIds: duplicates.length,
    },
  };
}
