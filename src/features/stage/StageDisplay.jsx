import { useEffect, useState } from 'react';
import { doc, getDoc, onSnapshot } from 'firebase/firestore';
import { Award, Clock, Maximize2 } from 'lucide-react';
import { db } from '../../firebase/config.js';
import { GUESTS_COLLECTION } from '../../services/participants.js';

// Public, unauthenticated display. It deliberately reads only ONE document
// (stageState/current) rather than the roster — Firestore rules give
// PUBLIC a plain get() on stageState and on a single participant, never a
// list() over participants. That also means this screen can only ever show
// the CURRENT guest, not a "next 5 up" preview — showing a queue of
// upcoming names to an unauthenticated audience would mean either opening
// up roster enumeration again, or exposing who's arrived before they're
// called, neither of which this design allows for PUBLIC.
export default function StageDisplay() {
  const [current, setCurrent] = useState(null);

  useEffect(() => {
    const unsub = onSnapshot(doc(db, 'stageState', 'current'), async (snap) => {
      const participantId = snap.exists() ? snap.data().participantId : null;
      if (!participantId) {
        setCurrent(null);
        return;
      }
      try {
        const pSnap = await getDoc(doc(db, GUESTS_COLLECTION, participantId));
        setCurrent(pSnap.exists() ? { id: pSnap.id, ...pSnap.data() } : null);
      } catch {
        setCurrent(null);
      }
    });
    return unsub;
  }, []);

  return (
    <div className="max-w-5xl mx-auto">
      <div className="bg-gradient-to-br from-slate-950 via-slate-900 to-blue-950 text-white rounded-3xl p-6 sm:p-12 text-center border-4 border-sky-400 shadow-2xl relative overflow-hidden">
        <div className="flex justify-between items-center mb-6">
          <span className="inline-flex items-center gap-1.5 text-xs font-bold text-sky-300 bg-sky-500/20 px-3 py-1 rounded-full border border-sky-400/30">
            <Award className="w-4 h-4" /> พิธีประดับบ่า • กำลังอยู่บนเวที
          </span>
          <button
            onClick={() => {
              if (!document.fullscreenElement) document.documentElement.requestFullscreen().catch(() => {});
              else document.exitFullscreen().catch(() => {});
            }}
            className="p-2 bg-slate-800 hover:bg-slate-700 rounded-xl text-slate-300 transition-colors"
          >
            <Maximize2 className="w-4 h-4" />
          </button>
        </div>

        {current ? (
          <div className="space-y-4 py-4 animate-in fade-in zoom-in duration-300">
            <div className="inline-block px-5 py-2 rounded-2xl bg-sky-500 text-slate-950 font-black text-xl sm:text-2xl shadow-lg">
              ป้ายลำดับ #{current.badgeNumber}
            </div>
            <h1 className="text-3xl sm:text-6xl font-black text-white tracking-tight leading-tight">{current.name}</h1>
            <div className="flex justify-center items-center gap-3 text-sm sm:text-lg text-sky-300 font-mono font-bold">
              <span>รหัสนักศึกษา: {current.studentId || '-'}</span>
              <span>•</span>
              <span className="bg-blue-800/80 px-3 py-1 rounded-xl text-white font-sans">{current.year}</span>
            </div>
          </div>
        ) : (
          <div className="py-12 text-slate-400">
            <Clock className="w-16 h-16 mx-auto mb-3 opacity-40 animate-spin" style={{ animationDuration: '6s' }} />
            <h3 className="text-2xl font-bold text-slate-300">กำลังเตรียมเริ่มการขานชื่อ</h3>
            <p className="text-xs text-slate-500 mt-1">รายชื่อจะปรากฏที่นี่ทันทีเมื่อพิธีกรเริ่มกดขานชื่อ</p>
          </div>
        )}
      </div>
    </div>
  );
}
