import { describe, expect, it } from 'vitest';
import { detectYearFromStudentId, getYearOrderWeight } from './yearDetect.js';

describe('detectYearFromStudentId', () => {
  it('maps the current four academic-year prefixes', () => {
    expect(detectYearFromStudentId('69014522')).toBe('ปี 1');
    expect(detectYearFromStudentId('68023567')).toBe('ปี 2');
    expect(detectYearFromStudentId('67037256')).toBe('ปี 3');
    expect(detectYearFromStudentId('66045914')).toBe('ปี 4');
  });

  it('treats anything 65 or below as a graduate', () => {
    expect(detectYearFromStudentId('65103113')).toBe('บัณฑิต');
    expect(detectYearFromStudentId('60000001')).toBe('บัณฑิต');
  });

  it('returns empty for missing or too-short input', () => {
    expect(detectYearFromStudentId('')).toBe('');
    expect(detectYearFromStudentId(null)).toBe('');
    expect(detectYearFromStudentId('6')).toBe('');
  });

  it('returns empty for a non-numeric, unrecognized prefix', () => {
    expect(detectYearFromStudentId('ab014522')).toBe('');
  });
});

describe('getYearOrderWeight', () => {
  it('orders the four academic years ascending', () => {
    expect(getYearOrderWeight('ปี 1')).toBeLessThan(getYearOrderWeight('ปี 2'));
    expect(getYearOrderWeight('ปี 2')).toBeLessThan(getYearOrderWeight('ปี 3'));
    expect(getYearOrderWeight('ปี 3')).toBeLessThan(getYearOrderWeight('ปี 4'));
  });

  it('places graduates after all four years', () => {
    expect(getYearOrderWeight('บัณฑิต')).toBeGreaterThan(getYearOrderWeight('ปี 4'));
  });

  it('falls back for an unrecognized value', () => {
    expect(getYearOrderWeight('')).toBe(99);
    expect(getYearOrderWeight('some other label')).toBe(50);
  });
});
