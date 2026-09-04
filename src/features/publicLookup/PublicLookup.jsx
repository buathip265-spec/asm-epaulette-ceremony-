import { useState } from 'react';
import { Loader2, SearchCheck } from 'lucide-react';
import { getParticipantByStudentId } from '../../services/participants.js';

// PUBLIC-safe by construction: this looks up exactly one record, by the
// exact student ID the visitor types in, via Firestore's get() — never a
// list()/query over the roster (see firestore.rules). There is no way to
// browse or search by name here; that requires STAFF and lives in
// features/checkin/CheckInStation.jsx instead. This is a deliberate
// trade-off from the audit: it keeps the full roster from ever being
// enumerable by an unauthenticated visitor.
export default function PublicLookup() {
  const [studentId, setStudentId] = useState('');
  const [result, setResult] = useState(undefined); // undefined = not searched yet, null = not found
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const handleSearch = async (e) => {
    e.preventDefault();
    if (!studentId.trim()) return;
    setLoading(true);
    setError('');
    try {
      const found = await getParticipantByStudentId(studentId);
      setResult(found);
    } catch {
      setError('ค้นหาไม่สำเร็จ กรุณาลองใหม่อีกครั้ง');
      setResult(undefined);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="max-w-lg mx-auto space-y-5">
      <div className="bg-gradient-to-r from-blue-900 to-sky-900 text-white p-6 rounded-3xl shadow-xl border-2 border-sky-400">
        <div className="flex items-center gap-2 text-sky-300 text-xs font-bold uppercase mb-2">
          <SearchCheck className="w-4 h-4" /> ตรวจสอบลำดับคิวขึ้นเวที
        </div>
        <h2 className="text-xl sm:text-2xl font-black">ค้นหาด้วยรหัสนักศึกษา</h2>

        <form onSubmit={handleSearch} className="mt-4 flex gap-2">
          <input
            type="text"
            value={studentId}
            onChange={(e) => setStudentId(e.target.value)}
            placeholder="กรอกรหัสนักศึกษา เช่น 69014522"
            className="flex-1 px-4 py-3 bg-white text-slate-900 rounded-2xl text-sm font-bold outline-none shadow-sm"
          />
          <button
            type="submit"
            disabled={loading}
            className="px-5 py-3 bg-sky-400 hover:bg-sky-300 disabled:bg-sky-200 text-slate-950 font-black rounded-2xl text-sm flex items-center gap-2"
          >
            {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : <SearchCheck className="w-4 h-4" />}
            <span>ค้นหา</span>
          </button>
        </form>
      </div>

      {error && <p className="text-xs font-bold text-red-600 text-center">{error}</p>}

      {result === null && (
        <div className="bg-white p-5 rounded-2xl shadow border text-center text-sm text-slate-500">
          ไม่พบรหัสนักศึกษานี้ในระบบ กรุณาตรวจสอบอีกครั้ง
        </div>
      )}

      {result && (
        <div className="bg-white text-slate-900 p-5 rounded-2xl shadow-lg flex items-center justify-between gap-3">
          <div className="flex items-center gap-3 min-w-0">
            <div className="w-14 h-14 rounded-xl bg-slate-900 text-sky-300 font-black flex flex-col items-center justify-center shrink-0">
              <span className="text-[8px] opacity-70">ป้าย</span>
              <span className="text-lg leading-none">#{result.badgeNumber}</span>
            </div>
            <div className="min-w-0">
              <p className="font-extrabold text-sm sm:text-base truncate">{result.name}</p>
              <p className="text-xs text-slate-500 font-mono">
                รหัส: {result.studentId} • {result.year}
              </p>
            </div>
          </div>
          <div className="shrink-0">
            {result.called ? (
              <span className="px-3 py-1 bg-purple-100 text-purple-800 font-bold text-xs rounded-xl border border-purple-200">✓ ขานชื่อแล้ว</span>
            ) : result.status === 'completed' || result.status === 'checked_in' ? (
              <span className="px-3 py-1 bg-emerald-100 text-emerald-800 font-bold text-xs rounded-xl border border-emerald-200">
                รอขึ้นเวที (ลำดับ #{result.badgeNumber})
              </span>
            ) : (
              <span className="px-3 py-1 bg-slate-100 text-slate-600 font-bold text-xs rounded-xl">ยังไม่ถึงจุดลงทะเบียน</span>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
