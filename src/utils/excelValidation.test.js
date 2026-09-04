import { describe, expect, it } from 'vitest';
import { extractParticipantRows, validateParticipantRows } from './excelValidation.js';

describe('extractParticipantRows', () => {
  it('maps Thai headers to the participant shape', () => {
    const raw = [{ ชั้นปี: 'ปี 1', รหัสนักศึกษา: '69014522', 'ชื่อ-นามสกุล': 'นายกิตติกร บุญมี', หมายเหตุ: 'สโมสร' }];
    const [row] = extractParticipantRows(raw);
    expect(row).toMatchObject({ year: 'ปี 1', studentId: '69014522', name: 'นายกิตติกร บุญมี', note: 'สโมสร' });
  });

  it('auto-detects year from student id when no year column matches', () => {
    const raw = [{ รหัสนักศึกษา: '68023567', ชื่อ: 'จิราภรณ์' }];
    const [row] = extractParticipantRows(raw);
    expect(row.year).toBe('ปี 2');
  });

  it('assigns 1-based, header-offset source row numbers', () => {
    const raw = [{ ชื่อ: 'A' }, { ชื่อ: 'B' }];
    const rows = extractParticipantRows(raw);
    expect(rows.map((r) => r.sourceRow)).toEqual([2, 3]);
  });
});

describe('validateParticipantRows', () => {
  const row = (overrides) => ({ sourceRow: 2, year: 'ปี 1', studentId: '69014522', name: 'ทดสอบ', note: '', ...overrides });

  it('accepts a clean row', () => {
    const result = validateParticipantRows([row({})]);
    expect(result.summary).toMatchObject({ total: 1, valid: 1, invalid: 0, duplicateStudentIds: 0 });
  });

  it('rejects a row with no name', () => {
    const result = validateParticipantRows([row({ name: '' })]);
    expect(result.summary.valid).toBe(0);
    expect(result.invalidRows[0].reasons).toContain('ไม่มีชื่อ-นามสกุล');
  });

  it('rejects a malformed student id but keeps validating other fields', () => {
    const result = validateParticipantRows([row({ studentId: 'abc' })]);
    expect(result.invalidRows[0].reasons.some((r) => r.includes('รหัสนักศึกษา'))).toBe(true);
  });

  it('allows an empty student id (e.g. an honoured guest with none)', () => {
    const result = validateParticipantRows([row({ studentId: '' })]);
    expect(result.summary.valid).toBe(1);
  });

  it('flags duplicate student ids and keeps only the first occurrence as valid', () => {
    const rows = [
      row({ sourceRow: 2, studentId: '69014522' }),
      row({ sourceRow: 3, studentId: '69014522' }),
    ];
    const result = validateParticipantRows(rows);
    expect(result.summary.duplicateStudentIds).toBe(1);
    expect(result.summary.valid).toBe(1);
    expect(result.validRows[0].sourceRow).toBe(2);
  });

  it('does not import a genuinely invalid row even if its student id is unique', () => {
    const result = validateParticipantRows([row({ name: '', studentId: '69099999' })]);
    expect(result.validRows).toHaveLength(0);
  });
});
