import { describe, expect, it } from 'vitest';
import { participantDocId } from './participantId.js';

describe('participantDocId', () => {
  it('uses the trimmed student id when present', () => {
    expect(participantDocId('  69014522  ')).toBe('69014522');
  });

  it('generates a fallback id when there is no student id', () => {
    const id = participantDocId('');
    expect(id).toMatch(/^manual_/);
  });

  it('generates distinct fallback ids on repeated calls', () => {
    expect(participantDocId('')).not.toBe(participantDocId(''));
  });
});
