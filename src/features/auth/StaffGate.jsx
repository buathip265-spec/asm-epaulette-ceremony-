import { Loader2, ShieldAlert } from 'lucide-react';
import { useAuth } from './AuthContext.jsx';
import LoginScreen from './LoginScreen.jsx';

// Wraps any staff-only or admin-only screen. This is a UX convenience —
// the actual enforcement is Firestore Security Rules (firestore.rules).
// Hiding a button here never substitutes for the rule; it just avoids
// showing someone a screen that would fail anyway.
export default function StaffGate({ require = 'staff', title, children }) {
  const { user, isStaff, isAdmin, loading } = useAuth();

  if (loading) {
    return (
      <div className="min-h-[50vh] flex items-center justify-center">
        <Loader2 className="w-8 h-8 text-sky-500 animate-spin" />
      </div>
    );
  }

  if (!user) {
    return <LoginScreen title={title} />;
  }

  const allowed = require === 'admin' ? isAdmin : isStaff;
  if (!allowed) {
    return (
      <div className="min-h-[50vh] flex flex-col items-center justify-center text-center p-6">
        <ShieldAlert className="w-10 h-10 text-amber-500 mb-3" />
        <h3 className="font-black text-slate-900">บัญชีนี้ไม่มีสิทธิ์เข้าถึงหน้านี้</h3>
        <p className="text-xs text-slate-500 mt-1 max-w-xs">
          กรุณาติดต่อผู้ดูแลระบบเพื่อขอสิทธิ์{require === 'admin' ? 'ผู้ดูแลระบบ' : 'เจ้าหน้าที่'}
        </p>
      </div>
    );
  }

  return children;
}
