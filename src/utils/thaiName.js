import { getYearOrderWeight } from './yearDetect.js';

const TITLE_PREFIX = /^(นาย|นางสาว|นาง|ด\.ช\.|ด\.ญ\.|ผศ\.|รศ\.|ดร\.)\s*/;

export function getSortableName(fullName) {
  if (!fullName) return '';
  return fullName.replace(TITLE_PREFIX, '').trim();
}

// Used for DISPLAY ordering only (e.g. sorting a page of results in the
// admin table, or an already-fetched search result). This never assigns or
// mutates badgeNumber — see src/utils/badges.js for that. Badge numbers are
// permanent and are not derived from this sort.
export function sortByYearThenName(participants) {
  return [...participants].sort((a, b) => {
    const weightA = getYearOrderWeight(a.year);
    const weightB = getYearOrderWeight(b.year);
    if (weightA !== weightB) return weightA - weightB;
    return getSortableName(a.name).localeCompare(getSortableName(b.name), 'th');
  });
}
