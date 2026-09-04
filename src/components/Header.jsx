import { Award, Bell, LogOut, Mic2, MonitorPlay, QrCode, Search, Settings, ShieldCheck, Smartphone, Volume2, VolumeX } from 'lucide-react';
import { useAuth } from '../features/auth/AuthContext.jsx';
import { useSound } from './SoundContext.jsx';

export const TABS = [
  { value: 'checkin', label: 'เช็คชื่อ', icon: Smartphone, require: 'staff' },
  { value: 'badges', label: 'รับป้าย', icon: Bell, require: 'staff' },
  { value: 'mc', label: 'ขึ้นเวที', icon: Mic2, require: 'staff' },
  { value: 'stage', label: 'จอ LED', icon: MonitorPlay, require: null },
  { value: 'lookup', label: 'ค้นหาคิว', icon: Search, require: null },
  { value: 'admin', label: 'จัดการ', icon: Settings, require: 'admin' },
];

export default function Header({ activeTab, setActiveTab, onOpenQr, syncStatus }) {
  const { user, isStaff, isAdmin, logout } = useAuth();
  const { enabled: soundEnabled, toggle: toggleSound } = useSound();

  // Every tab is always listed — clicking one that needs a role the visitor
  // doesn't have lands on StaffGate's login/no-access screen, never a
  // silent no-op. Hiding the tab entirely would just be UI dressing, since
  // the real enforcement is Firestore rules either way.
  return (
    <header className="bg-slate-950 text-white sticky top-0 z-40 shadow-xl border-b border-sky-500/20">
      <div className="max-w-7xl mx-auto px-3 sm:px-4 py-2.5 flex flex-col md:flex-row items-center justify-between gap-2.5">
        <div className="flex items-center justify-between w-full md:w-auto">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-gradient-to-tr from-blue-700 via-blue-600 to-sky-400 text-white p-2 rounded-2xl flex items-center justify-center shadow-md shrink-0">
              <Award className="w-6 h-6" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-sm sm:text-base font-black text-sky-300 tracking-wide">ระบบเช็คชื่อพิธีวันเกียรติยศ</h1>
                <span
                  className={`inline-flex items-center gap-1 text-[10px] px-2.5 py-0.5 rounded-full font-bold border ${
                    syncStatus === 'connected' ? 'bg-sky-500/20 text-sky-300 border-sky-500/40' : 'bg-blue-900/40 text-blue-300 border-blue-500/40'
                  }`}
                >
                  <span className={`w-1.5 h-1.5 rounded-full ${syncStatus === 'connected' ? 'bg-sky-400 animate-ping' : 'bg-blue-400'}`}></span>
                  {syncStatus === 'connected' ? 'Online' : 'Local'}
                </span>
              </div>
              {user && (
                <div className="text-[10px] text-emerald-400 font-bold flex items-center gap-1">
                  <ShieldCheck className="w-3 h-3" /> {isAdmin ? 'ผู้ดูแลระบบ' : isStaff ? 'เจ้าหน้าที่' : user.email}
                </div>
              )}
            </div>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2 w-full md:w-auto justify-end">
          <div className="flex items-center bg-slate-900 p-1 rounded-xl border border-slate-800 overflow-x-auto max-w-full">
            {TABS.map((t) => {
              const Icon = t.icon;
              return (
                <button
                  key={t.value}
                  onClick={() => setActiveTab(t.value)}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 ${
                    activeTab === t.value ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                  }`}
                >
                  <Icon className="w-3.5 h-3.5" />
                  <span>{t.label}</span>
                </button>
              );
            })}
          </div>

          <button onClick={onOpenQr} className="px-3.5 py-1.5 rounded-xl border border-sky-400 bg-sky-500 hover:bg-sky-400 text-slate-950 font-black text-xs flex items-center gap-1.5 shadow-md">
            <QrCode className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">QR Code</span>
          </button>

          <button onClick={toggleSound} className={`p-2 rounded-xl border text-xs ${soundEnabled ? 'border-sky-400 text-sky-300 bg-sky-400/10' : 'border-slate-800 text-slate-500 bg-slate-900'}`}>
            {soundEnabled ? <Volume2 className="w-4 h-4" /> : <VolumeX className="w-4 h-4" />}
          </button>

          {user && (
            <button onClick={logout} className="px-2.5 py-1.5 rounded-xl border border-red-500/40 text-red-300 bg-red-500/10 hover:bg-red-500/20 text-xs font-bold flex items-center gap-1">
              <LogOut className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">ออกจากระบบ</span>
            </button>
          )}
        </div>
      </div>
    </header>
  );
}
