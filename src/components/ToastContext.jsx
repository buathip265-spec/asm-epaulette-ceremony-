import { createContext, useCallback, useContext, useRef, useState } from 'react';
import { AlertCircle, Check, Info } from 'lucide-react';

// A single, consistent way to surface SAVING/SUCCESS/ERROR outcomes to the
// person using the device. This exists specifically so no write handler
// ever has to choose between "silently fail" and "silently pretend it
// worked" — every mutation in this app calls notify() with a real outcome.
const ToastContext = createContext(null);

let idCounter = 0;

export function ToastProvider({ children }) {
  const [toasts, setToasts] = useState([]);
  const timers = useRef({});

  const dismiss = useCallback((id) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
    clearTimeout(timers.current[id]);
    delete timers.current[id];
  }, []);

  const notify = useCallback(
    ({ type = 'info', message, duration = 4000 }) => {
      const id = ++idCounter;
      setToasts((prev) => [...prev, { id, type, message }]);
      timers.current[id] = setTimeout(() => dismiss(id), duration);
      return id;
    },
    [dismiss]
  );

  return (
    <ToastContext.Provider value={{ notify }}>
      {children}
      <div className="fixed top-4 left-1/2 -translate-x-1/2 z-[100] flex flex-col gap-2 items-center px-4 w-full max-w-md">
        {toasts.map((t) => (
          <div
            key={t.id}
            role="status"
            className={`w-full px-4 py-2.5 rounded-2xl shadow-xl flex items-center gap-2 text-xs sm:text-sm font-bold text-white animate-in fade-in slide-in-from-top-4 duration-200 ${
              t.type === 'success' ? 'bg-emerald-600' : t.type === 'error' ? 'bg-red-600' : 'bg-slate-800'
            }`}
          >
            {t.type === 'success' ? (
              <Check className="w-4 h-4 shrink-0" />
            ) : t.type === 'error' ? (
              <AlertCircle className="w-4 h-4 shrink-0" />
            ) : (
              <Info className="w-4 h-4 shrink-0" />
            )}
            <span>{t.message}</span>
          </div>
        ))}
      </div>
    </ToastContext.Provider>
  );
}

export function useToast() {
  const ctx = useContext(ToastContext);
  if (!ctx) throw new Error('useToast must be used within ToastProvider');
  return ctx;
}
