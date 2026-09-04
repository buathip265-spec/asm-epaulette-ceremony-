import { useEffect, useState } from 'react';
import { Mic2, ShieldCheck, SkipForward } from 'lucide-react';
import { useAuth } from '../auth/AuthContext.jsx';
import { useToast } from '../../components/ToastContext.jsx';
import { listenByFlag, listenCallQueue, setCalled, setSkipped } from '../../services/participants.js';
import { triggerHaptic } from '../../utils/device.js';

const FILTERS = [
  { value: 'calling', label: 'กำลังเรียก' },
  { value: 'skipped', label: 'ข้ามคิว' },
  { value: 'called', label: 'ขานแล้ว' },
];

export default function CallQueue() {
  const { user } = useAuth();
  const { notify } = useToast();
  const [filter, setFilter] = useState('calling');
  const [callingPool, setCallingPool] = useState([]);
  const [skippedPool, setSkippedPool] = useState([]);
  const [calledPool, setCalledPool] = useState([]);
  const [busyId, setBusyId] = useState(null);

  useEffect(() => {
    const onError = (err) => {
      console.error('CallQueue listener failed:', err);
      notify({ type: 'error', message: 'ไม่สามารถโหลดคิวได้ กรุณาตรวจสอบการเชื่อมต่อ' });
    };
    const unsub1 = listenCallQueue(setCallingPool, onError);
    const unsub2 = listenByFlag('skipped', true, setSkippedPool, onError);
    const unsub3 = listenByFlag('called', true, setCalledPool, onError);
    return () => {
      unsub1();
      unsub2();
      unsub3();
    };
  }, [notify]);

  const list = filter === 'calling' ? callingPool : filter === 'skipped' ? skippedPool : calledPool;

  const handleCall = async (guest) => {
    setBusyId(guest.id);
    triggerHaptic();
    try {
      await setCalled(guest.id, true, user.uid);
    } catch {
      notify({ type: 'error', message: 'ไม่สามารถขานชื่อได้ กรุณาลองใหม่' });
    } finally {
      setBusyId(null);
    }
  };

  const handleSkip = async (guest) => {
    setBusyId(guest.id);
    triggerHaptic();
    try {
      await setSkipped(guest.id, !guest.skipped);
    } catch {
      notify({ type: 'error', message: 'ไม่สามารถปรับคิวได้ กรุณาลองใหม่' });
    } finally {
      setBusyId(null);
    }
  };

  return (
    <div className="space-y-4 max-w-4xl mx-auto">
      <div className="bg-slate-950 text-white rounded-3xl p-5 sm:p-6 border-2 border-slate-800 shadow-xl flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
        <div>
          <div className="flex items-center gap-2">
            <h2 className="text-xl sm:text-2xl font-black text-sky-300 flex items-center gap-2">
              <Mic2 className="w-6 h-6 text-sky-400" /> ลำดับขึ้นเวที
            </h2>
            <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 text-xs font-bold rounded-md flex items-center gap-1">
              <ShieldCheck className="w-3.5 h-3.5" /> พิธีกร
            </span>
          </div>
          <p className="text-xs text-slate-400 mt-1">เรียงตามลำดับป้าย — กดขานชื่อหรือข้ามคิวได้</p>
        </div>
        <div className="flex items-center gap-1 bg-slate-900 p-1 rounded-2xl border border-slate-800 text-xs">
          {FILTERS.map((f) => (
            <button
              key={f.value}
              onClick={() => setFilter(f.value)}
              className={`px-2.5 py-1.5 rounded-xl font-bold transition-all ${
                filter === f.value ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
              }`}
            >
              {f.label} ({f.value === 'calling' ? callingPool.length : f.value === 'skipped' ? skippedPool.length : calledPool.length})
            </button>
          ))}
        </div>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        {list.length === 0 && (
          <div className="col-span-full py-14 text-center text-slate-400 text-xs bg-white rounded-2xl border border-dashed">
            ไม่มีรายชื่อในคิวนี้
          </div>
        )}
        {list.map((guest) => (
          <div
            key={guest.id}
            className={`p-3.5 sm:p-4 rounded-3xl border-2 flex items-center justify-between gap-3 ${
              guest.called ? 'bg-slate-100 border-slate-200 opacity-60' : guest.skipped ? 'bg-amber-50 border-amber-300' : 'bg-sky-50/90 border-sky-300'
            }`}
          >
            <div className="flex items-center gap-3 min-w-0">
              <div className={`w-11 h-11 sm:w-12 sm:h-12 rounded-2xl font-black flex flex-col items-center justify-center shrink-0 ${guest.skipped ? 'bg-amber-400 text-slate-950' : 'bg-slate-900 text-sky-300'}`}>
                <span className="text-[7px] leading-none opacity-80">ลำดับ</span>
                <span className="text-sm sm:text-base font-extrabold">#{guest.badgeNumber}</span>
              </div>
              <div className="min-w-0">
                <div className={`font-black text-xs sm:text-sm truncate ${guest.called ? 'line-through text-slate-400' : 'text-slate-900'}`}>{guest.name}</div>
                <div className="text-[11px] sm:text-xs text-slate-500 flex items-center gap-1.5 mt-0.5">
                  <span className="bg-blue-100 text-blue-800 px-1.5 py-0.2 rounded text-[10px] font-bold">{guest.year}</span>
                  <span className="font-mono text-blue-700 font-bold">{guest.studentId}</span>
                </div>
              </div>
            </div>
            <div className="flex items-center gap-1.5 shrink-0">
              {!guest.called && (
                <button
                  onClick={() => handleSkip(guest)}
                  disabled={busyId === guest.id}
                  className={`p-2 rounded-xl text-xs font-bold transition-colors disabled:opacity-50 ${
                    guest.skipped ? 'bg-amber-200 hover:bg-amber-300 text-amber-900' : 'bg-slate-200 hover:bg-slate-300 text-slate-700'
                  }`}
                >
                  <SkipForward className="w-4 h-4" />
                </button>
              )}
              <button
                onClick={() => handleCall(guest)}
                disabled={guest.called || busyId === guest.id}
                className={`px-3 py-1.5 sm:px-4 sm:py-2 rounded-2xl text-xs font-black shrink-0 transition-colors shadow-sm disabled:opacity-60 ${
                  guest.called ? 'bg-slate-300 text-slate-700' : 'bg-sky-500 hover:bg-sky-400 text-slate-950 shadow-md'
                }`}
              >
                {guest.called ? '✓ ขานแล้ว' : '🎙️ ขานชื่อ'}
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
