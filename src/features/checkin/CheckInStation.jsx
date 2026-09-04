import { useEffect, useMemo, useState } from 'react';
import { Award, CheckCheck, GraduationCap, Loader2, Search } from 'lucide-react';
import { useAuth } from '../auth/AuthContext.jsx';
import { useToast } from '../../components/ToastContext.jsx';
import StatusPill from '../../components/StatusPill.jsx';
import {
  CheckInError,
  checkInParticipant,
  listenByYear,
  searchByNamePrefix,
  searchByStudentIdPrefix,
} from '../../services/participants.js';
import { triggerHaptic } from '../../utils/device.js';

const YEAR_TABS = [
  { value: 'all', label: 'ทุกชั้นปี' },
  { value: 'ปี 1', label: 'ปี 1 (69)' },
  { value: 'ปี 2', label: 'ปี 2 (68)' },
  { value: 'ปี 3', label: 'ปี 3 (67)' },
  { value: 'ปี 4', label: 'ปี 4 (66)' },
];

const BROWSE_LIMIT = 300;

export default function CheckInStation() {
  const { user } = useAuth();
  const { notify } = useToast();

  const [yearTab, setYearTab] = useState('all');
  const [query, setQuery] = useState('');
  const [browseList, setBrowseList] = useState([]);
  const [searchResults, setSearchResults] = useState(null); // null = not searching
  const [searching, setSearching] = useState(false);
  const [selected, setSelected] = useState(null);
  const [checkingIn, setCheckingIn] = useState(false);

  // Live, but scoped to one year at a time (or capped at BROWSE_LIMIT for
  // "all") — never the whole roster. See the audit's scalability section
  // for why an unscoped listener doesn't hold up at real-event scale.
  useEffect(() => {
    const unsub = listenByYear(
      yearTab,
      setBrowseList,
      (err) => {
        console.error('listenByYear failed:', err);
        notify({ type: 'error', message: 'ไม่สามารถโหลดรายชื่อได้ กรุณาตรวจสอบการเชื่อมต่อ' });
      },
      { limitCount: BROWSE_LIMIT }
    );
    return unsub;
  }, [yearTab, notify]);

  useEffect(() => {
    const q = query.trim();
    if (!q) {
      setSearchResults(null);
      return;
    }
    setSearching(true);
    const timer = setTimeout(async () => {
      try {
        const [byName, byId] = await Promise.all([searchByNamePrefix(q), searchByStudentIdPrefix(q)]);
        const merged = new Map();
        [...byId, ...byName].forEach((p) => merged.set(p.id, p));
        setSearchResults(Array.from(merged.values()));
      } catch {
        notify({ type: 'error', message: 'ค้นหาไม่สำเร็จ กรุณาลองใหม่' });
        setSearchResults([]);
      } finally {
        setSearching(false);
      }
    }, 300);
    return () => clearTimeout(timer);
  }, [query, notify]);

  const visibleList = useMemo(() => {
    if (searchResults !== null) {
      return yearTab === 'all' ? searchResults : searchResults.filter((g) => g.year === yearTab);
    }
    return browseList;
  }, [searchResults, browseList, yearTab]);

  const handleConfirmCheckIn = async () => {
    if (!selected) return;
    setCheckingIn(true);
    try {
      await checkInParticipant(selected.id, user.uid);
      triggerHaptic([60, 50, 60]);
      notify({ type: 'success', message: `เช็กชื่อสำเร็จ — ${selected.name}` });
      setSelected(null);
    } catch (err) {
      if (err instanceof CheckInError && err.code === 'ALREADY_CHECKED_IN') {
        notify({ type: 'info', message: err.message });
      } else {
        notify({ type: 'error', message: 'ไม่สามารถบันทึกการเช็กชื่อได้ กรุณาลองใหม่' });
      }
      // Deliberately does NOT close the modal or mutate any local status —
      // the live listener is the only thing allowed to move a participant
      // out of 'pending'. If the write failed, they're still pending here.
    } finally {
      setCheckingIn(false);
    }
  };

  return (
    <div className="max-w-3xl mx-auto space-y-4">
      <div className="bg-gradient-to-br from-slate-950 via-slate-900 to-blue-950 text-white rounded-3xl p-5 sm:p-7 text-center border-2 border-sky-500/40 shadow-xl">
        <div className="inline-flex items-center gap-2 bg-sky-400/20 text-sky-300 px-3.5 py-1 rounded-full font-bold text-xs border border-sky-500/30 mb-2">
          <Award className="w-3.5 h-3.5" /> พิธีวันเกียรติยศ • จุดเช็คชื่อ (เจ้าหน้าที่)
        </div>
        <h2 className="text-xl sm:text-2xl font-black text-sky-300">ค้นหาและยืนยันตัวตนผู้เข้าร่วมงาน</h2>
        <p className="text-xs text-slate-300 mt-1">พิมพ์ชื่อหรือรหัสนักศึกษา ตรวจสอบตัวตน แล้วกดยืนยันเช็กชื่อ</p>
      </div>

      <div className="bg-white p-1.5 rounded-2xl border-2 border-sky-200 shadow-sm flex items-center gap-1.5 overflow-x-auto">
        {YEAR_TABS.map((t) => (
          <button
            key={t.value}
            onClick={() => setYearTab(t.value)}
            className={`flex-1 min-w-[85px] py-2 px-2.5 rounded-xl font-black text-xs transition-all ${
              yearTab === t.value ? 'bg-sky-500 text-slate-950 shadow-md' : 'text-slate-600 hover:bg-slate-100'
            }`}
          >
            {t.label}
          </button>
        ))}
      </div>

      <div className="relative shadow-sm">
        <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 w-5 h-5" />
        <input
          type="text"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="พิมพ์ชื่อ หรือรหัสนักศึกษา..."
          className="w-full pl-12 pr-4 py-3 rounded-2xl border-2 border-slate-200 bg-white text-sm focus:ring-4 focus:ring-sky-500/20 focus:border-sky-400 outline-none shadow-sm"
        />
        {searching && <Loader2 className="absolute right-4 top-1/2 -translate-y-1/2 w-4 h-4 animate-spin text-sky-500" />}
      </div>

      {searchResults !== null && (
        <p className="text-[11px] text-slate-400 px-1">
          ผลการค้นหา {visibleList.length} รายการ{' '}
          <span className="text-slate-300">(แสดงสูงสุด 25 รายการต่อประเภทการค้นหา)</span>
        </p>
      )}
      {searchResults === null && visibleList.length >= BROWSE_LIMIT && (
        <p className="text-[11px] text-amber-600 px-1 font-bold">
          แสดง {BROWSE_LIMIT} รายชื่อแรกของ{yearTab === 'all' ? 'ทั้งหมด' : yearTab} — พิมพ์ค้นหาเพื่อดูรายชื่ออื่น
        </p>
      )}

      <div className="space-y-2.5">
        {visibleList.length === 0 && (
          <div className="py-14 text-center text-slate-400 text-xs bg-white rounded-2xl border border-dashed">
            ไม่พบรายชื่อ
          </div>
        )}
        {visibleList.map((guest) => {
          const isCheckedIn = guest.status === 'checked_in';
          const isCompleted = guest.status === 'completed';
          return (
            <div
              key={guest.id}
              onClick={() => {
                if (guest.status === 'pending') {
                  triggerHaptic();
                  setSelected(guest);
                }
              }}
              className={`p-3 sm:p-3.5 rounded-2xl border-2 flex items-center justify-between gap-3 transition-all ${
                isCompleted
                  ? 'bg-slate-50 border-slate-200 opacity-60'
                  : isCheckedIn
                  ? 'bg-sky-50/90 border-sky-400 shadow-md'
                  : 'bg-white hover:border-sky-400 border-slate-200 cursor-pointer shadow-sm hover:shadow-md'
              }`}
            >
              <div className="flex items-center gap-3 min-w-0">
                <div
                  className={`w-11 h-11 sm:w-12 sm:h-12 rounded-2xl font-black flex flex-col items-center justify-center shrink-0 ${
                    isCompleted ? 'bg-emerald-100 text-emerald-800' : isCheckedIn ? 'bg-sky-500 text-slate-950' : 'bg-slate-900 text-sky-300'
                  }`}
                >
                  <span className="text-[7px] uppercase font-bold opacity-70">ลำดับ</span>
                  <span className="text-base font-extrabold">#{guest.badgeNumber}</span>
                </div>
                <div className="min-w-0">
                  <div className="font-extrabold text-slate-900 text-xs sm:text-sm truncate">{guest.name}</div>
                  <div className="flex flex-wrap items-center gap-1.5 mt-0.5">
                    <span className="text-[11px] font-mono font-bold text-blue-700 bg-blue-50 px-1.5 py-0.2 rounded border border-blue-100">
                      รหัส {guest.studentId || '-'}
                    </span>
                    <span className="text-[10px] font-bold text-slate-600 bg-slate-100 px-1.5 py-0.2 rounded flex items-center gap-1">
                      <GraduationCap className="w-3 h-3" /> {guest.year}
                    </span>
                  </div>
                </div>
              </div>
              <div className="shrink-0">
                {isCompleted ? (
                  <StatusPill state="success" label="รับป้ายแล้ว" />
                ) : isCheckedIn ? (
                  <StatusPill state="checked_in" label="รอรับป้าย..." />
                ) : (
                  <button className="bg-slate-900 hover:bg-sky-500 hover:text-slate-950 text-sky-300 px-4 py-2 rounded-xl font-black text-xs shadow-md transition-colors">
                    เช็คชื่อ
                  </button>
                )}
              </div>
            </div>
          );
        })}
      </div>

      {selected && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs">
          <div className="bg-white rounded-3xl p-6 sm:p-7 max-w-sm w-full shadow-2xl text-center border border-slate-100">
            <div className="w-16 h-16 bg-slate-900 text-sky-400 rounded-2xl flex flex-col items-center justify-center mx-auto mb-3">
              <span className="text-[9px] text-slate-400 font-bold uppercase">ป้ายลำดับ</span>
              <span className="text-2xl font-black">#{selected.badgeNumber}</span>
            </div>
            <h3 className="text-xl font-black text-slate-900">ยืนยันตัวตน</h3>
            <p className="text-xs text-slate-500 mt-1">โปรดตรวจสอบชื่อให้ถูกต้องก่อนกดยืนยัน</p>

            <div className="bg-sky-50 rounded-2xl p-4 my-4 border-2 border-sky-200 text-center">
              <div className="text-base sm:text-lg font-black text-slate-900">{selected.name}</div>
              <div className="flex justify-center items-center gap-2 mt-2">
                <span className="font-mono text-xs font-bold text-blue-700 bg-white px-2 py-0.5 rounded border border-blue-100">
                  รหัส: {selected.studentId || '-'}
                </span>
                <span className="text-xs font-bold text-sky-950 bg-sky-200/80 px-2 py-0.5 rounded">{selected.year}</span>
              </div>
            </div>

            <div className="flex gap-3">
              <button
                onClick={() => setSelected(null)}
                disabled={checkingIn}
                className="flex-1 py-3 border-2 border-slate-200 hover:bg-slate-50 rounded-2xl font-bold text-xs sm:text-sm text-slate-600 disabled:opacity-50"
              >
                ยกเลิก
              </button>
              <button
                onClick={handleConfirmCheckIn}
                disabled={checkingIn}
                className="flex-1 py-3 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-300 text-white font-black rounded-2xl text-xs sm:text-sm shadow-md flex items-center justify-center gap-1.5"
              >
                {checkingIn ? <Loader2 className="w-4 h-4 animate-spin" /> : <CheckCheck className="w-4 h-4" />}
                <span>{checkingIn ? 'กำลังบันทึก...' : 'ยืนยันเช็กชื่อ'}</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
