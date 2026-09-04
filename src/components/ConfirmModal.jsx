import { useState } from 'react';
import { AlertTriangle, Loader2 } from 'lucide-react';

// Generic confirm dialog. For destructive/irreversible-feeling actions
// (Excel replace, full status reset, delete), pass `requireTypedPhrase` —
// the confirm button stays disabled until the admin types it exactly,
// which is a deliberate extra beat beyond a single "OK" click for actions
// covered in the backup & recovery plan.
export default function ConfirmModal({
  isOpen,
  title,
  message,
  confirmText = 'ยืนยัน',
  confirmColor = 'bg-red-600 hover:bg-red-700',
  requireTypedPhrase,
  busy = false,
  onConfirm,
  onCancel,
}) {
  const [typed, setTyped] = useState('');
  if (!isOpen) return null;

  const canConfirm = !requireTypedPhrase || typed.trim() === requireTypedPhrase;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs">
      <div className="bg-white rounded-3xl p-6 max-w-sm w-full shadow-2xl text-center border">
        <div className="w-12 h-12 bg-sky-100 text-sky-600 rounded-2xl flex items-center justify-center mx-auto mb-3">
          <AlertTriangle className="w-6 h-6" />
        </div>
        <h3 className="text-lg font-black text-slate-900">{title}</h3>
        <p className="text-xs text-slate-500 mt-2 mb-4 leading-relaxed">{message}</p>

        {requireTypedPhrase && (
          <div className="mb-4 text-left">
            <label className="text-[11px] font-bold text-slate-600 block mb-1">
              พิมพ์ <span className="font-mono bg-slate-100 px-1 rounded">{requireTypedPhrase}</span> เพื่อยืนยัน
            </label>
            <input
              type="text"
              value={typed}
              onChange={(e) => setTyped(e.target.value)}
              className="w-full px-3 py-2 border-2 border-slate-200 rounded-xl text-sm font-mono outline-none focus:border-sky-400"
              autoFocus
            />
          </div>
        )}

        <div className="flex gap-3">
          <button
            onClick={() => {
              setTyped('');
              onCancel();
            }}
            disabled={busy}
            className="flex-1 py-2.5 border-2 rounded-xl font-bold text-xs text-slate-600 disabled:opacity-50"
          >
            ยกเลิก
          </button>
          <button
            onClick={onConfirm}
            disabled={!canConfirm || busy}
            className={`flex-1 py-2.5 text-white font-bold rounded-xl text-xs shadow-md flex items-center justify-center gap-1.5 disabled:opacity-50 disabled:cursor-not-allowed ${confirmColor}`}
          >
            {busy && <Loader2 className="w-3.5 h-3.5 animate-spin" />}
            <span>{confirmText}</span>
          </button>
        </div>
      </div>
    </div>
  );
}
