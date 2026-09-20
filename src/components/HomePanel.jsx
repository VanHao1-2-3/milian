import { useEffect, useMemo, useState } from 'react';
import {
  FlaskConical,
  BookOpen,
  Boxes,
  ArrowRight,
  Sparkles,
  Clock3,
  CalendarDays
} from 'lucide-react';
import { useLanguage } from '../lib/i18n.jsx';

export default function HomePanel({ onNavigate }) {
  const { lang, t } = useLanguage();

  const [now, setNow] = useState(new Date());

  /*
   * ==============================
   * CẬP NHẬT THỜI GIAN THỰC
   * ==============================
   */
  useEffect(() => {
    const timer = setInterval(() => {
      setNow(new Date());
    }, 1000);

    return () => clearInterval(timer);
  }, []);

  /*
   * ==============================
   * THỜI GIAN
   * ==============================
   */
  const timeText = useMemo(() => {
    return new Intl.DateTimeFormat(
      lang === 'zh' ? 'zh-CN' : 'vi-VN',
      {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
      }
    ).format(now);
  }, [now, lang]);

  const dateText = useMemo(() => {
    return new Intl.DateTimeFormat(
      lang === 'zh' ? 'zh-CN' : 'vi-VN',
      {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
      }
    ).format(now);
  }, [now, lang]);

  /*
   * ==============================
   * LỜI CHÀO THEO GIỜ
   * ==============================
   */
  const greeting = useMemo(() => {
    const hour = now.getHours();

    if (lang === 'zh') {
      if (hour < 6) return '夜深了';
      if (hour < 12) return '早上好';
      if (hour < 14) return '中午好';
      if (hour < 18) return '下午好';
      return '晚上好';
    }

    if (hour < 6) return 'Khuya rồi';
    if (hour < 12) return 'Chào buổi sáng';
    if (hour < 14) return 'Chào buổi trưa';
    if (hour < 18) return 'Chào buổi chiều';

    return 'Chào buổi tối';
  }, [now, lang]);

  /*
   * ==============================
   * DANH SÁCH MODULE (CẤU HÌNH TỰ ĐỘNG)
   * Khi thêm tính năng mới chỉ cần thêm 1 object vào mảng này.
   * ==============================
   */
  const features = useMemo(() => [
    {
      key: 'nc-calc',
      title: t.featureNcTitle || (lang === 'zh' ? 'NC数据计算' : 'Tính toán NC'),
      description: t.featureNcDesc || (lang === 'zh' ? '按配方汇总，扣除掺用比例' : 'Tổng hợp theo công thức, trừ tỷ lệ dùng lại'),
      icon: FlaskConical,
      available: true
    },
    {
      key: 'rubber-chinese',
      title: t.navRubberChinese || (lang === 'zh' ? '密炼车间中文' : 'Tiếng Trung Xưởng Luyện'),
      description: t.navRubberChineseHint || (lang === 'zh' ? '密炼车间专业词汇' : 'Từ vựng chuyên ngành xưởng luyện'),
      icon: BookOpen,
      available: true
    },
    {
      key: 'rubber-inventory',
      title: t.navRubberInventory || (lang === 'zh' ? '胶料库存' : 'Tồn kho su'),
      description: t.navRubberInventoryHint || (lang === 'zh' ? '批次库存与FIFO检查' : 'Tồn kho theo lô và kiểm tra FIFO'),
      icon: Boxes,
      available: true
    }
  ], [lang, t]);

  const availableCount = features.filter(item => item.available).length;

  return (
    <div className="w-full">
      {/* HEADER */}
      <header className="mb-6 sm:mb-7">
        <p className="text-sm text-ink-soft mb-1">{greeting}</p>
        <h1 className="text-xl sm:text-[22px] font-bold tracking-tight">{t.homeTitle}</h1>
        <p className="text-ink-soft mt-1.5 max-w-[62ch] text-sm sm:text-base">{t.homeSubtitle}</p>
      </header>

      {/* TIME & DATE */}
      <section className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-6">
        <div className="bg-surface border border-line rounded-xl p-5 shadow-card flex items-center gap-4">
          <div className="w-11 h-11 shrink-0 rounded-xl bg-accent-soft flex items-center justify-center">
            <Clock3 size={20} className="text-accent-dark" />
          </div>
          <div className="min-w-0">
            <p className="text-xs text-ink-faint mb-1">{lang === 'zh' ? '当前时间' : 'Thời gian hiện tại'}</p>
            <div className="text-2xl sm:text-[26px] font-bold tracking-tight tabular-nums">{timeText}</div>
          </div>
        </div>

        <div className="bg-surface border border-line rounded-xl p-5 shadow-card flex items-center gap-4">
          <div className="w-11 h-11 shrink-0 rounded-xl bg-canvas border border-line flex items-center justify-center">
            <CalendarDays size={20} className="text-ink-soft" />
          </div>
          <div className="min-w-0">
            <p className="text-xs text-ink-faint mb-1">{lang === 'zh' ? '今天' : 'Hôm nay'}</p>
            <div className="text-sm sm:text-[15px] font-semibold capitalize">{dateText}</div>
          </div>
        </div>
      </section>

      {/* MODULE HEADER */}
      <div className="flex flex-col sm:flex-row sm:items-end sm:justify-between gap-2 mb-4">
        <div>
          <h2 className="font-semibold text-[15px]">{lang === 'zh' ? '功能模块' : 'Chức năng'}</h2>
          <p className="text-xs text-ink-faint mt-1">
            {lang === 'zh' ? `${availableCount} 个模块可用` : `${availableCount} module đang hoạt động`}
          </p>
        </div>
      </div>

      {/* GRID HÀNG NGANG TỰ ĐỘNG THEO SỐ LƯỢNG MODULE */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
        {features.map(feature => {
          const Icon = feature.icon;

          return (
            <button
              key={feature.key}
              type="button"
              onClick={() => {
                if (feature.available) {
                  onNavigate(feature.key);
                }
              }}
              disabled={!feature.available}
              className="group text-left bg-surface border border-line rounded-xl p-5 shadow-card hover:border-accent hover:-translate-y-0.5 transition-all flex flex-col justify-between gap-4 disabled:cursor-default disabled:hover:border-line disabled:hover:translate-y-0"
            >
              <div className="space-y-3">
                <div className="w-10 h-10 rounded-lg bg-accent-soft flex items-center justify-center">
                  <Icon size={19} className="text-accent-dark" />
                </div>
                <div>
                  <h3 className="font-semibold text-[15px]">{feature.title}</h3>
                  <p className="text-[13px] text-ink-soft mt-1 leading-relaxed">{feature.description}</p>
                </div>
              </div>

              <span className="text-[13px] font-semibold text-accent-dark flex items-center gap-1 pt-1">
                {t.featureOpen || (lang === 'zh' ? '打开功能' : 'Mở chức năng')}
                <ArrowRight size={14} className="group-hover:translate-x-0.5 transition-transform" />
              </span>
            </button>
          );
        })}

        {/* THẺ TÍNH NĂNG SẮP RA MẮT */}
        <div className="bg-surface/60 border border-dashed border-line rounded-xl p-5 flex flex-col gap-3 text-ink-faint">
          <div className="w-10 h-10 rounded-lg bg-canvas flex items-center justify-center">
            <Sparkles size={19} />
          </div>
          <div>
            <h3 className="font-semibold text-[15px] text-ink-faint">
              {t.comingSoonTitle || (lang === 'zh' ? '即将推出' : 'Sắp ra mắt')}
            </h3>
            <p className="text-[13px] mt-1 leading-relaxed">
              {t.comingSoonDesc || (lang === 'zh' ? '以后会在此处添加其他功能。' : 'Các tính năng mới sẽ được cập nhật tại đây.')}
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}