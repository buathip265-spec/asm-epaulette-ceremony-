import { createContext, useContext, useState } from 'react';

const KEY = 'ceremony.soundEnabled';
const SoundContext = createContext(null);

// A shared, tab-wide preference — deliberately a Context, not a
// per-component localStorage hook. The header's toggle button and the
// badge queue's arrival chime are two different components that both need
// to see the SAME value change the instant it's toggled; localStorage's
// own "storage" event only fires in OTHER tabs, never the one that made
// the change, which would leave them silently out of sync in the same tab.
export function SoundProvider({ children }) {
  const [enabled, setEnabled] = useState(() => {
    if (typeof localStorage === 'undefined') return true;
    const stored = localStorage.getItem(KEY);
    return stored === null ? true : stored === 'true';
  });

  const toggle = () => {
    setEnabled((prev) => {
      const next = !prev;
      localStorage.setItem(KEY, String(next));
      return next;
    });
  };

  return <SoundContext.Provider value={{ enabled, toggle }}>{children}</SoundContext.Provider>;
}

export function useSound() {
  const ctx = useContext(SoundContext);
  if (!ctx) throw new Error('useSound must be used within SoundProvider');
  return ctx;
}
