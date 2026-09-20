import { useState } from 'react';
import {
  LayoutGrid,
  FlaskConical,
  CircleDot,
  Languages,
  BookOpen,
  Package,
  PanelLeftClose,
  PanelLeftOpen
} from 'lucide-react';
import HomePanel from './components/HomePanel.jsx';
import NcCalculator from './components/NcCalculator.jsx';
import RubberChinese from './components/RubberChinese.jsx';
import RubberInventory from './components/RubberInventory.jsx';
import { useLanguage } from './lib/i18n.jsx';

export default function App() {
  const [view, setView] = useState('home');
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const { lang, setLang, t } = useLanguage();

  const NAV_ITEMS = [
    { key: 'home', label: t.navHome, icon: LayoutGrid },
    { key: 'nc-calc', label: t.navNcCalc, hint: t.navNcCalcHint, icon: FlaskConical },
    { key: 'rubber-chinese', label: t.navRubberChinese, hint: t.navRubberChineseHint, icon: BookOpen },
    {
      key: 'rubber-inventory',
      label: t.navRubberInventory,
      hint: t.navRubberInventoryHint,
      icon: Package
    }
  ];

  return (
    <div className="flex min-h-screen">
      <aside
        className={`
    relative
    shrink-0
    bg-sidebar
    text-slate-200
    flex
    flex-col
    p-5
    transition-all
    duration-300
    ease-in-out
    ${sidebarOpen ? 'w-64' : 'w-[76px]'}
  `}
      >
        <div
          className={`
    flex
    items-center
    pb-5
    mb-4
    border-b
    border-white/10
    ${sidebarOpen ? 'gap-3' : 'justify-center'}
  `}
        >
          <div className="w-9 h-9 shrink-0 rounded-lg font-mono font-bold text-sm flex items-center justify-center">
            <img
              src="logo.png"
              alt=""
              className="w-full h-full object-contain"
            />
          </div>

          {sidebarOpen && (
            <span className="font-semibold text-[15px] text-slate-50 whitespace-nowrap overflow-hidden">
              {t.appName}
            </span>
          )}
        </div>
        <button
          type="button"
          onClick={() => setSidebarOpen(open => !open)}
          className="
    absolute
    -right-3
    top-20
    z-20
    w-6
    h-6
    rounded-full
    bg-sidebar
    border
    border-white/10
    text-slate-400
    flex
    items-center
    justify-center
    hover:text-white
    hover:bg-accent
    transition-colors
    shadow-md
  "
          aria-label={sidebarOpen ? 'Thu gọn menu' : 'Mở rộng menu'}
        >
          {sidebarOpen ? (
            <PanelLeftClose size={14} />
          ) : (
            <PanelLeftOpen size={14} />
          )}
        </button>

        <nav className="flex flex-col gap-1">
          {NAV_ITEMS.map(({ key, label, hint, icon: Icon }) => (
            <button
              key={key}
              onClick={() => setView(key)}
              title={!sidebarOpen ? label : undefined}
              className={`
    flex
    items-start
    rounded-lg
    transition-colors
    ${sidebarOpen
                  ? 'gap-2.5 text-left px-3 py-2.5'
                  : 'justify-center px-2 py-3'
                }
    ${view === key
                  ? 'bg-accent/15 text-slate-50'
                  : 'text-slate-300 hover:bg-white/5'
                }
  `}
            >
              <Icon size={16} className={`mt-0.5 shrink-0 ${view === key ? 'text-accent' : 'text-slate-400'}`} />
              {sidebarOpen && (
                <span className="flex flex-col overflow-hidden">
                  <span className="text-[13.5px] font-semibold whitespace-nowrap">
                    {label}
                  </span>

                  {hint && (
                    <span className="text-[11.5px] text-slate-400 whitespace-nowrap">
                      {hint}
                    </span>
                  )}
                </span>
              )}
            </button>
          ))}
        </nav>

      </aside>

      <main
        className={`
    flex-1
    min-w-0
    px-6
    sm:px-8
    lg:px-10
    py-6
    lg:py-9
    transition-all
    duration-300
  `}
      >
        <div className="flex justify-end mb-5">
          <div className="flex items-center gap-1.5 bg-canvas border border-line rounded-lg p-1">
            <Languages size={13} className="text-ink-faint ml-1.5 shrink-0" />
            <button
              onClick={() => setLang('vi')}
              className={`text-[12px] font-semibold px-3 py-1.5 rounded-md transition-colors ${lang === 'vi' ? 'bg-accent text-[#241605]' : 'text-ink-soft hover:text-ink'}`}
            >
              Tiếng Việt
            </button>
            <button
              onClick={() => setLang('zh')}
              className={`text-[12px] font-semibold px-3 py-1.5 rounded-md transition-colors ${lang === 'zh' ? 'bg-accent text-[#241605]' : 'text-ink-soft hover:text-ink'}`}
            >
              中文
            </button>
          </div>
        </div>

        {view === 'home' && <HomePanel onNavigate={setView} />}
        {view === 'nc-calc' && <NcCalculator />}
        {view === 'rubber-chinese' && <RubberChinese />}
        {view === 'rubber-inventory' && <RubberInventory />}
      </main>
    </div>
  );
}
