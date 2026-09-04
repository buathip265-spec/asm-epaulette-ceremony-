import { createContext, useContext, useEffect, useState } from 'react';
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from 'firebase/auth';
import { doc, onSnapshot } from 'firebase/firestore';
import { auth, db } from '../../firebase/config.js';

// Roles are resolved from staffUsers/{uid}.role, not from a custom claim —
// see firestore.rules for why (no Cloud Functions / Admin SDK needed).
// A signed-in user with no staffUsers document has role `null`, which the
// rest of the app treats exactly like PUBLIC: staff-only screens simply
// don't open for them.
const AuthContext = createContext(null);

export function AuthProvider({ children }) {
  const [user, setUser] = useState(null);
  const [role, setRole] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [roleLoading, setRoleLoading] = useState(false);

  useEffect(() => {
    const unsubAuth = onAuthStateChanged(auth, (nextUser) => {
      setUser(nextUser);
      setAuthLoading(false);
    });
    return unsubAuth;
  }, []);

  useEffect(() => {
    if (!user) {
      setRole(null);
      return undefined;
    }
    setRoleLoading(true);
    const unsubRole = onSnapshot(
      doc(db, 'staffUsers', user.uid),
      (snap) => {
        setRole(snap.exists() ? snap.data().role : null);
        setRoleLoading(false);
      },
      () => {
        // Read denied or failed — treat as no role rather than crashing.
        setRole(null);
        setRoleLoading(false);
      }
    );
    return unsubRole;
  }, [user]);

  const login = (email, password) => signInWithEmailAndPassword(auth, email, password);
  const logout = () => signOut(auth);

  const value = {
    user,
    role, // 'staff' | 'admin' | null
    isStaff: role === 'staff' || role === 'admin',
    isAdmin: role === 'admin',
    loading: authLoading || roleLoading,
    login,
    logout,
  };

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

export function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error('useAuth must be used within AuthProvider');
  return ctx;
}
