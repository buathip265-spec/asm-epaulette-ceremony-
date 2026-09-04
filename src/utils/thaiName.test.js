import { describe, expect, it } from 'vitest';
import { getSortableName, sortByYearThenName } from './thaiName.js';

describe('getSortableName', () => {
  it('strips a leading Thai honorific', () => {
    expect(getSortableName('นายสมชาย ใจดี')).toBe('สมชาย ใจดี');
    expect(getSortableName('นางสาวจิราภรณ์ ทัดศรี')).toBe('จิราภรณ์ ทัดศรี');
    expect(getSortableName('ดร.สมหญิง รักเรียน')).toBe('สมหญิง รักเรียน');
  });

  it('leaves a name with no recognized prefix untouched', () => {
    expect(getSortableName('Somchai Jaidee')).toBe('Somchai Jaidee');
  });

  it('handles empty input', () => {
    expect(getSortableName('')).toBe('');
    expect(getSortableName(null)).toBe('');
  });
});

describe('sortByYearThenName', () => {
  it('groups by year first, then sorts names within a year', () => {
    // Plain leading consonants (ก, ข) — unlike a leading sara-e vowel (เ),
    // their Thai collation order is unambiguous across ICU builds, which
    // is what makes this a stable assertion rather than one tied to a
    // specific runtime's locale tailoring.
    const input = [
      { name: 'นายขาว ขาว', year: 'ปี 2' },
      { name: 'นายขจร ขจร', year: 'ปี 1' },
      { name: 'นายกล้า กล้า', year: 'ปี 1' },
    ];
    const result = sortByYearThenName(input);
    expect(result.map((r) => r.name)).toEqual(['นายกล้า กล้า', 'นายขจร ขจร', 'นายขาว ขาว']);
  });

  it('does not mutate the input array', () => {
    const input = [{ name: 'บี', year: 'ปี 2' }, { name: 'เอ', year: 'ปี 1' }];
    const copy = [...input];
    sortByYearThenName(input);
    expect(input).toEqual(copy);
  });
});
