import { useState } from 'react';
import { LayoutGrid, FlaskConical, CircleDot, Languages } from 'lucide-react';
import HomePanel from './components/HomePanel.jsx';
import NcCalculator from './components/NcCalculator.jsx';
import { useLanguage } from './lib/i18n.jsx';

export default function App() {
  const [view, setView] = useState('home');
  const { lang, setLang, t } = useLanguage();

  const NAV_ITEMS = [
    { key: 'home', label: t.navHome, icon: LayoutGrid },
    { key: 'nc-calc', label: t.navNcCalc, hint: t.navNcCalcHint, icon: FlaskConical }
  ];

  return (
    <div className="flex min-h-screen">
      <aside className="w-64 shrink-0 bg-sidebar text-slate-200 flex flex-col p-5">
        <div className="flex items-center justify-between gap-3 pb-5 mb-4 border-b border-white/10">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-lg bg-accent text-[#241605] font-mono font-bold text-sm flex items-center justify-center">
              NC
            </div>
            <span className="font-semibold text-[15px] text-slate-50">{t.appName}</span>
          </div>
        </div>

        <nav className="flex flex-col gap-1">
          {NAV_ITEMS.map(({ key, label, hint, icon: Icon }) => (
            <button
              key={key}
              onClick={() => setView(key)}
              className={`flex items-start gap-2.5 text-left px-3 py-2.5 rounded-lg transition-colors ${
                view === key ? 'bg-accent/15 text-slate-50' : 'text-slate-300 hover:bg-white/5'
              }`}
            >
              <Icon size={16} className={`mt-0.5 shrink-0 ${view === key ? 'text-accent' : 'text-slate-400'}`} />
              <span className="flex flex-col">
                <span className="text-[13.5px] font-semibold">{label}</span>
                {hint && <span className="text-[11.5px] text-slate-400">{hint}</span>}
              </span>
            </button>
          ))}
        </nav>

        <div className="mt-auto flex flex-col gap-3">
          <div className="flex items-center gap-1.5 bg-white/5 rounded-lg p-1">
            <Languages size={13} className="text-slate-400 ml-1.5 shrink-0" />
            <button
              onClick={() => setLang('vi')}
              className={`flex-1 text-[12px] font-semibold py-1.5 rounded-md transition-colors ${lang === 'vi' ? 'bg-accent text-[#241605]' : 'text-slate-300 hover:text-slate-50'}`}
            >
              Tiếng Việt
            </button>
            <button
              onClick={() => setLang('zh')}
              className={`flex-1 text-[12px] font-semibold py-1.5 rounded-md transition-colors ${lang === 'zh' ? 'bg-accent text-[#241605]' : 'text-slate-300 hover:text-slate-50'}`}
            >
              中文
            </button>
          </div>

          <div className="flex items-center gap-2 text-[11.5px] text-slate-400">
            <CircleDot size={10} className="text-emerald-400 shrink-0" />
            {/* {t.footerOffline} */}
          </div>
        </div>
      </aside>

      <main className="flex-1 px-10 py-9 max-w-5xl">
        {view === 'home' && <HomePanel onNavigate={setView} />}
        {view === 'nc-calc' && <NcCalculator />}
      </main>
    </div>
  );
}
