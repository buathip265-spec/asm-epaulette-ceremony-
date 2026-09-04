// รหัสนักศึกษาปีล่าสุด 2 หลักแรก -> ชั้นปี (69=ปี1, 68=ปี2, 67=ปี3, 66=ปี4, <=65=บัณฑิต)
export function detectYearFromStudentId(studentId) {
  if (!studentId || studentId.trim().length < 2) return '';
  const prefix = studentId.trim().substring(0, 2);
  if (prefix === '69') return 'ปี 1';
  if (prefix === '68') return 'ปี 2';
  if (prefix === '67') return 'ปี 3';
  if (prefix === '66') return 'ปี 4';
  const num = parseInt(prefix, 10);
  if (!isNaN(num) && num <= 65) return 'บัณฑิต';
  return '';
}

export const YEAR_WEIGHTS = {
  'ปี 1': 1,
  'ปี 2': 2,
  'ปี 3': 3,
  'ปี 4': 4,
  บัณฑิต: 5,
  อาจารย์: 6,
  แขกผู้มีเกียรติ: 7,
};

export function getYearOrderWeight(yearStr) {
  if (!yearStr) return 99;
  for (const [key, weight] of Object.entries(YEAR_WEIGHTS)) {
    if (yearStr.includes(key)) return weight;
  }
  return 50;
}

export const YEAR_CODE_LABELS = {
  'ปี 1': '69',
  'ปี 2': '68',
  'ปี 3': '67',
  'ปี 4': '66',
};
