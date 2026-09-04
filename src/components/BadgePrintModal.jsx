import { Printer, X } from 'lucide-react';

export default function BadgePrintModal({ guest, onClose }) {
  if (!guest) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-xs">
      <div className="bg-white rounded-3xl p-6 max-w-sm w-full shadow-2xl border text-center">
        <div className="flex justify-between items-center mb-4">
          <h3 className="font-black text-slate-900 text-base">ตัวอย่างป้ายชื่อสำหรับพิมพ์</h3>
          <button onClick={onClose} className="text-slate-400 p-1 print:hidden">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div
          data-print-badge
          className="p-6 bg-gradient-to-b from-sky-50 to-white rounded-2xl border-2 border-blue-600 shadow-md text-center space-y-2"
        >
          <div className="text-[10px] font-black tracking-widest text-blue-800 uppercase">
            พิธีวันเกียรติยศ • ประดับบ่า
          </div>
          <div className="w-16 h-16 rounded-2xl bg-slate-900 text-sky-300 font-black text-2xl flex flex-col items-center justify-center mx-auto shadow-inner">
            <span className="text-[8px] opacity-70">ลำดับ</span>
            #{guest.badgeNumber}
          </div>
          <h2 className="font-black text-lg text-slate-900 pt-1 leading-snug">{guest.name}</h2>
          <p className="text-xs font-mono font-bold text-blue-700">{guest.studentId || '-'}</p>
          <div className="inline-block px-3 py-0.5 bg-blue-100 text-blue-900 text-xs font-bold rounded-full">
            {guest.year}
          </div>
        </div>

        <div className="flex gap-2 mt-5 print:hidden">
          <button onClick={onClose} className="flex-1 py-2.5 border-2 rounded-xl text-xs font-bold text-slate-600">
            ปิด
          </button>
          <button
            onClick={() => window.print()}
            className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold flex items-center justify-center gap-1.5 shadow-md"
          >
            <Printer className="w-4 h-4" /> พิมพ์ป้ายนี้
          </button>
        </div>
      </div>
    </div>
  );
}
