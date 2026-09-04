import { WifiOff } from 'lucide-react';
import { useOnlineStatus } from '../hooks/useOnlineStatus.js';

// Shown whenever the browser itself reports no network — a distinct,
// higher-confidence signal than a single Firestore listener error, which
// can also mean "permission denied" or "briefly reconnecting."
export default function OfflineBanner() {
  const isOnline = useOnlineStatus();
  if (isOnline) return null;

  return (
    <div className="bg-red-500 text-white text-xs py-1.5 px-4 text-center font-bold flex items-center justify-center gap-2">
      <WifiOff className="w-4 h-4 shrink-0" />
      <span>ไม่มีการเชื่อมต่ออินเทอร์เน็ต — ระบบจะซิงค์ข้อมูลอัตโนมัติเมื่อเชื่อมต่อได้อีกครั้ง</span>
    </div>
  );
}
