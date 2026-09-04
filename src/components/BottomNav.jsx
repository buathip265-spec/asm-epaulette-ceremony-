import { TABS } from './Header.jsx';

export default function BottomNav({ activeTab, setActiveTab }) {
  return (
    <nav className="md:hidden fixed bottom-0 left-0 right-0 z-40 bg-slate-950/95 backdrop-blur-lg border-t border-sky-900/50 px-1 py-1.5 flex items-center justify-around shadow-2xl">
      {TABS.map((t) => {
        const Icon = t.icon;
        return (
          <button
            key={t.value}
            onClick={() => setActiveTab(t.value)}
            className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl transition-all ${
              activeTab === t.value ? 'text-sky-400 font-bold' : 'text-slate-400'
            }`}
          >
            <Icon className="w-4 h-4 mb-0.5" />
            <span className="text-[9px]">{t.label}</span>
          </button>
        );
      })}
    </nav>
  );
}
