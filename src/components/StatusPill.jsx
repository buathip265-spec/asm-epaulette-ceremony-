import { AlertCircle, Check, Clock, Loader2, WifiOff } from 'lucide-react';

// The six states called out in the brief, given one consistent visual
// vocabulary used everywhere a participant's status is shown, instead of
// each screen inventing its own badge styling.
const VARIANTS = {
  waiting: { label: 'รอเรียก', icon: Clock, cls: 'bg-slate-100 text-slate-600 border-slate-200' },
  checked_in: { label: 'เช็กชื่อแล้ว', icon: Check, cls: 'bg-sky-100 text-sky-800 border-sky-300' },
  error: { label: 'เกิดข้อผิดพลาด', icon: AlertCircle, cls: 'bg-red-50 text-red-700 border-red-200' },
  offline: { label: 'ออฟไลน์', icon: WifiOff, cls: 'bg-slate-800 text-slate-200 border-slate-700' },
  saving: { label: 'กำลังบันทึก...', icon: Loader2, cls: 'bg-amber-50 text-amber-700 border-amber-200', spin: true },
  success: { label: 'สำเร็จ', icon: Check, cls: 'bg-emerald-100 text-emerald-800 border-emerald-300' },
};

export default function StatusPill({ state, label }) {
  const variant = VARIANTS[state] || VARIANTS.waiting;
  const Icon = variant.icon;
  return (
    <span
      className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-xl border text-[11px] font-bold ${variant.cls}`}
    >
      <Icon className={`w-3.5 h-3.5 ${variant.spin ? 'animate-spin' : ''}`} />
      <span>{label || variant.label}</span>
    </span>
  );
}
