import { useMemo, useState } from 'react';
import { Copy, Edit3, QrCode, X } from 'lucide-react';
import { copyTextSafely } from '../utils/device.js';
import { useToast } from './ToastContext.jsx';

const TARGET_TABS = [
  { value: 'stage', label: 'จอ LED เวที' },
  { value: 'lookup', label: 'ค้นหาลำดับคิว (ผู้ปกครอง)' },
  { value: 'checkin', label: 'จุดเช็คชื่อ (เจ้าหน้าที่)' },
];

export default function QrModal({ onClose }) {
  const { notify } = useToast();
  const [target, setTarget] = useState('stage');
  const [customBaseUrl, setCustomBaseUrl] = useState(
    typeof window !== 'undefined' ? window.location.origin + window.location.pathname : ''
  );
  const [isEditingUrl, setIsEditingUrl] = useState(false);

  const url = useMemo(() => {
    const base = (customBaseUrl.trim() || '').split('?')[0];
    return `${base}?tab=${target}`;
  }, [customBaseUrl, target]);

  const handleCopy = async () => {
    try {
      await copyTextSafely(url);
      notify({ type: 'success', message: 'คัดลอกลิงก์แล้ว' });
    } catch {
      notify({ type: 'error', message: 'คัดลอกลิงก์ไม่สำเร็จ' });
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/75 backdrop-blur-xs">
      <div className="bg-white rounded-3xl p-6 sm:p-7 max-w-md w-full shadow-2xl border flex flex-col">
        <div className="flex justify-between items-center mb-4 pb-3 border-b">
          <div className="flex items-center gap-2.5">
            <div className="w-10 h-10 rounded-2xl bg-sky-100 text-sky-700 flex items-center justify-center">
              <QrCode className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-base font-black text-slate-900">QR Code เข้าใช้งานระบบ</h3>
              <p className="text-[11px] text-slate-500">สแกนเพื่อเปิดใช้งานบนอุปกรณ์อื่น</p>
            </div>
          </div>
          <button onClick={onClose} className="text-slate-400 p-1">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="space-y-3.5 text-xs">
          <div>
            <label className="font-bold text-slate-700 block mb-1.5">เลือกหน้าที่ต้องการให้เปิด:</label>
            <div className="grid grid-cols-1 gap-1.5 p-1 bg-slate-100 rounded-2xl">
              {TARGET_TABS.map((t) => (
                <button
                  key={t.value}
                  onClick={() => setTarget(t.value)}
                  className={`py-2 px-2 rounded-xl font-bold text-center transition-all ${
                    target === t.value ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-600'
                  }`}
                >
                  {t.label}
                </button>
              ))}
            </div>
          </div>

          <div className="bg-slate-50 p-4 rounded-3xl border text-center flex flex-col items-center justify-center">
            <div className="bg-white p-3 rounded-2xl border-2 border-sky-300 shadow-md mb-2">
              <img
                src={`https://api.qrserver.com/v1/create-qr-code/?size=190x190&data=${encodeURIComponent(url)}&color=0284c7&bgcolor=ffffff`}
                alt="QR Code"
                className="w-44 h-44 object-contain"
              />
            </div>
          </div>

          <div>
            <div className="flex justify-between items-center mb-1">
              <label className="font-bold text-slate-700">ลิงก์เว็บไซต์:</label>
              <button
                type="button"
                onClick={() => setIsEditingUrl(!isEditingUrl)}
                className="text-sky-600 font-bold flex items-center gap-1 text-[11px]"
              >
                <Edit3 className="w-3 h-3" />
                <span>{isEditingUrl ? 'ซ่อนแก้ไข' : 'แก้ไข URL'}</span>
              </button>
            </div>
            {isEditingUrl && (
              <input
                type="text"
                value={customBaseUrl}
                onChange={(e) => setCustomBaseUrl(e.target.value)}
                className="w-full mb-2 px-2.5 py-1.5 bg-white rounded-lg border border-sky-300 text-xs font-mono outline-none"
              />
            )}
            <div className="flex gap-2">
              <input
                type="text"
                readOnly
                value={url}
                className="flex-1 px-3 py-2 bg-slate-100 rounded-xl border border-slate-200 text-slate-600 text-xs font-mono outline-none select-all"
              />
              <button
                onClick={handleCopy}
                className="px-4 py-2 bg-sky-500 hover:bg-sky-400 active:scale-95 text-slate-950 font-black rounded-xl text-xs flex items-center gap-1.5 shrink-0"
              >
                <Copy className="w-3.5 h-3.5" />
                <span>คัดลอก</span>
              </button>
            </div>
          </div>
        </div>

        <div className="mt-5 pt-3 border-t">
          <button
            onClick={onClose}
            className="w-full py-2.5 bg-slate-950 hover:bg-slate-800 text-sky-300 font-bold rounded-xl text-xs"
          >
            ปิดหน้าต่าง
          </button>
        </div>
      </div>
    </div>
  );
}
