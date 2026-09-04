#!/usr/bin/env node
// Generates a small, clearly-marked Excel file of TEST participants — for
// verifying the check-in flow works end-to-end before real event data goes
// in. Every name starts with "ทดสอบ" (test) and every row's note flags it,
// so it's unmistakable in the admin table and safe to tell apart from real
// participants. Purely local — writes a file, never touches Firestore.
//
// Usage: node scripts/generate-test-data.mjs
// Output: test-data-DO-NOT-USE-FOR-REAL-EVENT.xlsx (project root)

import * as XLSX from 'xlsx';

const TEST_NOTE = 'TEST DATA - ลบก่อนใช้งานจริง';

const rows = [
  { ชั้นปี: 'ปี 1', รหัสนักศึกษา: '69000001', 'ชื่อ-นามสกุล': 'ทดสอบ หนึ่ง (Test)', หมายเหตุ: TEST_NOTE },
  { ชั้นปี: 'ปี 1', รหัสนักศึกษา: '69000002', 'ชื่อ-นามสกุล': 'ทดสอบ สอง (Test)', หมายเหตุ: TEST_NOTE },
  { ชั้นปี: 'ปี 1', รหัสนักศึกษา: '69000003', 'ชื่อ-นามสกุล': 'ทดสอบ สาม (Test)', หมายเหตุ: TEST_NOTE },
  { ชั้นปี: 'ปี 2', รหัสนักศึกษา: '68000001', 'ชื่อ-นามสกุล': 'ทดสอบ สี่ (Test)', หมายเหตุ: TEST_NOTE },
  { ชั้นปี: 'ปี 2', รหัสนักศึกษา: '68000002', 'ชื่อ-นามสกุล': 'ทดสอบ ห้า (Test)', หมายเหตุ: TEST_NOTE },
  { ชั้นปี: 'ปี 2', รหัสนักศึกษา: '68000003', 'ชื่อ-นามสกุล': 'ทดสอบ หก (Test)', หมายเหตุ: TEST_NOTE },
  { ชั้นปี: 'ปี 3', รหัสนักศึกษา: '67000001', 'ชื่อ-นามสกุล': 'ทดสอบ เจ็ด (Test)', หมายเหตุ: TEST_NOTE },
  { ชั้นปี: 'ปี 3', รหัสนักศึกษา: '67000002', 'ชื่อ-นามสกุล': 'ทดสอบ แปด (Test)', หมายเหตุ: TEST_NOTE },
  { ชั้นปี: 'ปี 4', รหัสนักศึกษา: '66000001', 'ชื่อ-นามสกุล': 'ทดสอบ เก้า (Test)', หมายเหตุ: TEST_NOTE },
  { ชั้นปี: 'ปี 4', รหัสนักศึกษา: '66000002', 'ชื่อ-นามสกุล': 'ทดสอบ สิบ (Test)', หมายเหตุ: TEST_NOTE },
];

const worksheet = XLSX.utils.json_to_sheet(rows);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'TEST DATA');
XLSX.writeFile(workbook, 'test-data-DO-NOT-USE-FOR-REAL-EVENT.xlsx');

console.log(`Wrote ${rows.length} test participants to test-data-DO-NOT-USE-FOR-REAL-EVENT.xlsx`);
console.log('Student IDs used: 69000001-69000003, 68000001-68000003, 67000001-67000002, 66000001-66000002');
console.log('(all in a 000xxx range that will never collide with a real student ID)');
