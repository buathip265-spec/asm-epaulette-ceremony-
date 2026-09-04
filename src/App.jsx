import { lazy, Suspense, useEffect, useState } from 'react';
import { Award, Loader2 } from 'lucide-react';
import { AuthProvider, useAuth } from './features/auth/AuthContext.jsx';
import { ToastProvider } from './components/ToastContext.jsx';
import { SoundProvider } from './components/SoundContext.jsx';
import Header from './components/Header.jsx';
import BottomNav from './components/BottomNav.jsx';
import OfflineBanner from './components/OfflineBanner.jsx';
import QrModal from './components/QrModal.jsx';
import StaffGate from './features/auth/StaffGate.jsx';
import { useOnlineStatus } from './hooks/useOnlineStatus.js';

// Each tab is its own chunk. This matters most for AdminPanel, which pulls
// in the xlsx parsing/writing library — there's no reason a kiosk device
// that only ever opens the check-in tab should download that on first
// load. React.lazy + Suspense gives us this split for free per-route.
const CheckInStation = lazy(() => import('./features/checkin/CheckInStation.jsx'));
const BadgeQueue = lazy(() => import('./features/badges/BadgeQueue.jsx'));
const CallQueue = lazy(() => import('./features/mc/CallQueue.jsx'));
const StageDisplay = lazy(() => import('./features/stage/StageDisplay.jsx'));
const PublicLookup = lazy(() => import('./features/publicLookup/PublicLookup.jsx'));
const AdminPanel = lazy(() => import('./features/admin/AdminPanel.jsx'));

function TabLoading() {
  return (
    <div className="flex items-center justify-center py-24">
      <Loader2 className="w-8 h-8 text-sky-500 animate-spin" />
    </div>
  );
}

const TAB_REQUIRES = {
  checkin: 'staff',
  badges: 'staff',
  mc: 'staff',
  admin: 'admin',
};

function readTabFromUrl() {
  if (typeof window === 'undefined') return null;
  const tab = new URLSearchParams(window.location.search).get('tab');
  const known = ['checkin', 'badges', 'mc', 'stage', 'lookup', 'admin'];
  return known.includes(tab) ? tab : null;
}

function Shell() {
  const { loading } = useAuth();
  const isOnline = useOnlineStatus();
  const [activeTab, setActiveTab] = useState(() => readTabFromUrl() || 'stage');
  const [qrOpen, setQrOpen] = useState(false);

  useEffect(() => {
    const url = new URL(window.location.href);
    url.searchParams.set('tab', activeTab);
    window.history.replaceState({}, '', url);
  }, [activeTab]);

  if (loading) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center bg-slate-950 text-white p-4">
        <div className="relative flex items-center justify-center mb-6">
          <div className="w-16 h-16 rounded-3xl bg-blue-500/20 border border-blue-500/40 flex items-center justify-center animate-pulse">
            <Award className="w-8 h-8 text-sky-400" />
          </div>
          <Loader2 className="w-20 h-20 text-sky-400 animate-spin absolute" />
        </div>
        <h2 className="text-xl font-bold text-sky-300">กำลังเปิดระบบพิธีวันเกียรติยศ...</h2>
      </div>
    );
  }

  const requiredRole = TAB_REQUIRES[activeTab];

  return (
    <div className="min-h-screen flex flex-col bg-slate-100 text-slate-800 font-sans pb-24 md:pb-0">
      <Header activeTab={activeTab} setActiveTab={setActiveTab} onOpenQr={() => setQrOpen(true)} syncStatus={isOnline ? 'connected' : 'offline'} />
      <OfflineBanner />

      <main className="flex-1 max-w-7xl w-full mx-auto p-3 sm:p-5">
        <Suspense fallback={<TabLoading />}>
          {requiredRole ? (
            <StaffGate require={requiredRole} title={requiredRole === 'admin' ? 'เข้าสู่ระบบผู้ดูแล' : 'เข้าสู่ระบบเจ้าหน้าที่'}>
              {activeTab === 'checkin' && <CheckInStation />}
              {activeTab === 'badges' && <BadgeQueue />}
              {activeTab === 'mc' && <CallQueue />}
              {activeTab === 'admin' && <AdminPanel />}
            </StaffGate>
          ) : (
            <>
              {activeTab === 'stage' && <StageDisplay />}
              {activeTab === 'lookup' && <PublicLookup />}
            </>
          )}
        </Suspense>
      </main>

      <BottomNav activeTab={activeTab} setActiveTab={setActiveTab} />
      {qrOpen && <QrModal onClose={() => setQrOpen(false)} />}
    </div>
  );
}

export default function App() {
  return (
    <AuthProvider>
      <ToastProvider>
        <SoundProvider>
          <Shell />
        </SoundProvider>
      </ToastProvider>
    </AuthProvider>
  );
}
