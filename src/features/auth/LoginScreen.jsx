import { useState } from 'react';
import { KeyRound, Loader2, LogIn } from 'lucide-react';
import { useAuth } from './AuthContext.jsx';

const ERROR_MESSAGES = {
  'auth/invalid-credential': 'อีเมลหรือรหัสผ่านไม่ถูกต้อง',
  'auth/invalid-email': 'รูปแบบอีเมลไม่ถูกต้อง',
  'auth/too-many-requests': 'ลองเข้าสู่ระบบผิดหลายครั้งเกินไป กรุณาลองใหม่ภายหลัง',
  'auth/user-disabled': 'บัญชีนี้ถูกระงับการใช้งาน',
};

export default function LoginScreen({ title = 'เข้าสู่ระบบเจ้าหน้าที่' }) {
  const { login } = useAuth();
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState('');

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setSubmitting(true);
    try {
      await login(email.trim(), password);
    } catch (err) {
      setError(ERROR_MESSAGES[err.code] || 'ไม่สามารถเข้าสู่ระบบได้ กรุณาลองใหม่อีกครั้ง');
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="min-h-[70vh] flex items-center justify-center p-4">
      <div className="bg-white rounded-3xl p-6 sm:p-8 max-w-sm w-full shadow-xl border">
        <div className="w-14 h-14 bg-sky-100 text-sky-700 rounded-2xl flex items-center justify-center mx-auto mb-3">
          <KeyRound className="w-7 h-7" />
        </div>
        <h2 className="text-lg font-black text-slate-900 text-center">{title}</h2>
        <p className="text-xs text-slate-500 mt-1 text-center">
          ใช้บัญชีที่ผู้ดูแลระบบสร้างให้ท่าน
        </p>

        <form onSubmit={handleSubmit} className="mt-5 space-y-3">
          <div>
            <label className="font-bold text-slate-700 block mb-1 text-xs">อีเมล</label>
            <input
              type="email"
              required
              autoComplete="username"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl text-sm focus:border-sky-400 outline-none"
              placeholder="staff@example.com"
            />
          </div>
          <div>
            <label className="font-bold text-slate-700 block mb-1 text-xs">รหัสผ่าน</label>
            <input
              type="password"
              required
              autoComplete="current-password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl text-sm focus:border-sky-400 outline-none"
              placeholder="••••••••"
            />
          </div>

          {error && (
            <p className="text-xs font-bold text-red-600" role="alert">
              {error}
            </p>
          )}

          <button
            type="submit"
            disabled={submitting}
            className="w-full py-3 bg-blue-600 hover:bg-blue-700 disabled:bg-slate-300 text-white font-black rounded-2xl text-sm shadow-md flex items-center justify-center gap-2 transition-colors"
          >
            {submitting ? <Loader2 className="w-4 h-4 animate-spin" /> : <LogIn className="w-4 h-4" />}
            <span>เข้าสู่ระบบ</span>
          </button>
        </form>
      </div>
    </div>
  );
}
