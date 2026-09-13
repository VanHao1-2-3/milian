import { FlaskConical, ArrowRight, Sparkles } from 'lucide-react';
import { useLanguage } from '../lib/i18n.jsx';

export default function HomePanel({ onNavigate }) {
  const { t } = useLanguage();

  return (
    <div>
      <header className="mb-7">
        <h1 className="text-[22px] font-bold tracking-tight">{t.homeTitle}</h1>
        <p className="text-ink-soft mt-1.5 max-w-[62ch]">{t.homeSubtitle}</p>
      </header>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <button
          onClick={() => onNavigate('nc-calc')}
          className="group text-left bg-surface border border-line rounded-xl2 p-5 shadow-card hover:border-accent transition-colors flex flex-col gap-3"
        >
          <div className="w-9 h-9 rounded-lg bg-accent-soft flex items-center justify-center">
            <FlaskConical size={18} className="text-accent-dark" />
          </div>
          <div>
            <h3 className="font-semibold text-[15px]">{t.featureNcTitle}</h3>
            <p className="text-[13px] text-ink-soft mt-1 leading-relaxed">{t.featureNcDesc}</p>
          </div>
          <span className="mt-auto text-[13px] font-semibold text-accent-dark flex items-center gap-1 pt-1">
            {t.featureOpen} <ArrowRight size={14} className="group-hover:translate-x-0.5 transition-transform" />
          </span>
        </button>

        <div className="text-left bg-surface/60 border border-dashed border-line rounded-xl2 p-5 flex flex-col gap-3 text-ink-faint">
          <div className="w-9 h-9 rounded-lg bg-canvas flex items-center justify-center">
            <Sparkles size={18} />
          </div>
          <div>
            <h3 className="font-semibold text-[15px] text-ink-faint">{t.comingSoonTitle}</h3>
            <p className="text-[13px] mt-1 leading-relaxed">{t.comingSoonDesc}</p>
          </div>
        </div>
      </div>
    </div>
  );
}
