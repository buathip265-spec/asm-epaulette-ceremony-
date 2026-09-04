import { useEffect, useRef, useState } from 'react';
import { Bell, CheckCircle2, Loader2, Printer, ShieldCheck, UserCheck } from 'lucide-react';
import { useAuth } from '../auth/AuthContext.jsx';
import { useToast } from '../../components/ToastContext.jsx';
import { handoverBadge, listenBadgeQueue } from '../../services/participants.js';
import { triggerHaptic, playArrivalChime } from '../../utils/device.js';
import { useSound } from '../../components/SoundContext.jsx';
import BadgePrintModal from '../../components/BadgePrintModal.jsx';

export default function BadgeQueue() {
  const { user } = useAuth();
  const { notify } = useToast();
  const { enabled: soundEnabled } = useSound();
  const [queue, setQueue] = useState([]);
  const [printGuest, setPrintGuest] = useState(null);
  const [handingOver, setHandingOver] = useState(null); // participant id in flight
  const knownIds = useRef(null); // null until the first snapshot lands
  // Read via a ref inside the listener so toggling sound never has to tear
  // down and reattach the Firestore listener itself — only whether the
  // chime plays should change, not the subscription.
  const soundEnabledRef = useRef(soundEnabled);
  soundEnabledRef.current = soundEnabled;

  useEffect(() => {
    const unsub = listenBadgeQueue((next) => {
      if (knownIds.current) {
        const isNewArrival = next.some((g) => !knownIds.current.has(g.id));
        if (isNewArrival) {
          playArrivalChime(soundEnabledRef.current);
          triggerHaptic([60, 50, 60]);
        }
      }
      knownIds.current = new Set(next.map((g) => g.id));
      setQueue(next);
    }, (err) => {
      console.error('listenBadgeQueue failed:', err);
      notify({ type: 'error', message: 'ไม่สามารถโหลดคิวรับป้ายได้ กรุณาตรวจสอบการเชื่อมต่อ' });
    });
    return unsub;
  }, [notify]);

  const handleHandover = async (guest) => {
    setHandingOver(guest.id);
    try {
      await handoverBadge(guest.id, user.uid);
      triggerHaptic(40);
      notify({ type: 'success', message: `มอบป้าย #${guest.badgeNumber} เรียบร้อย` });
    } catch {
      notify({ type: 'error', message: 'ไม่สามารถบันทึกการมอบป้ายได้ กรุณาลองใหม่' });
    } finally {
      setHandingOver(null);
    }
  };

  return (
    <div className="space-y-4">
      <div className="bg-white rounded-3xl p-5 sm:p-6 border-2 border-sky-400 shadow-xl relative overflow-hidden">
        <div className="absolute top-0 left-0 w-full h-2 bg-gradient-to-r from-blue-600 to-sky-400"></div>
        <div className="flex items-center gap-2 pb-4 border-b border-slate-100">
          <h2 className="text-xl sm:text-2xl font-black text-slate-900 flex items-center gap-2">
            <Bell className="w-6 h-6 text-sky-500" /> คิวรับป้ายชื่อ
          </h2>
          <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 text-xs font-bold rounded-md flex items-center gap-1">
            <ShieldCheck className="w-3.5 h-3.5" /> เจ้าหน้าที่
          </span>
        </div>

        {queue.length === 0 ? (
          <div className="py-20 flex flex-col items-center justify-center text-center bg-slate-50 rounded-2xl border border-dashed border-slate-300 mt-4">
            <div className="w-16 h-16 bg-slate-200/80 rounded-full flex items-center justify-center mb-3 text-slate-400">
              <UserCheck className="w-8 h-8" />
            </div>
            <h3 className="text-lg font-bold text-slate-700">ไม่มีคิวค้างในขณะนี้</h3>
            <p className="text-xs text-slate-400 mt-1">เมื่อมีคนเช็กชื่อ รายชื่อจะขึ้นมาที่นี่ทันที</p>
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 sm:gap-5 mt-4">
            {queue.map((guest) => (
              <div key={guest.id} className="bg-white border-2 border-sky-400 rounded-3xl p-4 sm:p-5 shadow-lg flex flex-col justify-between">
                <div>
                  <div className="flex items-center gap-3.5">
                    <div className="w-14 h-14 sm:w-16 sm:h-16 bg-slate-900 text-sky-300 rounded-2xl flex flex-col items-center justify-center font-black shrink-0">
                      <span className="text-[8px] text-slate-400 font-bold">ป้ายลำดับ</span>
                      <span className="text-xl sm:text-2xl font-black">#{guest.badgeNumber}</span>
                    </div>
                    <div className="min-w-0 flex-1">
                      <span className="inline-block px-2 py-0.5 bg-sky-50 text-sky-800 border border-sky-200 text-[10px] font-black rounded mb-1">
                        {guest.year}
                      </span>
                      <div className="font-black text-slate-900 text-sm sm:text-base truncate">{guest.name}</div>
                      <div className="text-xs font-mono font-bold text-blue-700 mt-0.5">รหัส: {guest.studentId || '-'}</div>
                    </div>
                  </div>
                  <div className="mt-3.5 p-2.5 bg-sky-50 rounded-xl border border-sky-200 text-xs font-bold text-sky-900 flex items-center justify-between">
                    <span>
                      เช็กชื่อเมื่อ: {guest.checkInTime ? guest.checkInTime.toDate().toLocaleTimeString('th-TH') : '—'}
                    </span>
                    <button onClick={() => setPrintGuest(guest)} className="text-blue-600 hover:text-blue-800 font-bold flex items-center gap-1 text-[11px]">
                      <Printer className="w-3.5 h-3.5" /> พิมพ์ป้าย
                    </button>
                  </div>
                </div>
                <button
                  onClick={() => handleHandover(guest)}
                  disabled={handingOver === guest.id}
                  className="mt-4 w-full py-3 bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-300 text-white font-black rounded-2xl text-xs sm:text-sm flex items-center justify-center gap-2 shadow-md transition-colors"
                >
                  {handingOver === guest.id ? <Loader2 className="w-5 h-5 animate-spin" /> : <CheckCircle2 className="w-5 h-5" />}
                  <span>มอบป้ายชื่อ #{guest.badgeNumber} เรียบร้อย</span>
                </button>
              </div>
            ))}
          </div>
        )}
      </div>

      <BadgePrintModal guest={printGuest} onClose={() => setPrintGuest(null)} />
    </div>
  );
}
