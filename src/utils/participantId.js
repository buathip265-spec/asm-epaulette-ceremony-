// Participants are keyed by their student ID whenever one exists. This is
// what lets an unauthenticated visitor look up their own single record with
// a plain document get() (see firestore.rules) without ever being able to
// list/enumerate the roster, and it structurally prevents two rows from
// ever sharing the same student ID (a second write with the same id just
// overwrites the same document instead of creating a duplicate).
//
// Guests with no student ID (e.g. an honoured faculty guest added by hand)
// fall back to a generated id, since a Firestore document id can't be empty.
export function participantDocId(studentId) {
  const trimmed = (studentId || '').trim();
  if (trimmed) return trimmed;
  return `manual_${crypto.randomUUID()}`;
}
