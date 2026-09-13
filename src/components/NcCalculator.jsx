import { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  UploadCloud, RotateCcw, Download, Copy, AlertTriangle,
  Loader2, CheckCircle2, XCircle, Clock, X, Plus
} from 'lucide-react';
import DatePicker from 'react-datepicker';
import { vi } from 'date-fns/locale/vi';
import { zhCN } from 'date-fns/locale/zh-CN';
import 'react-datepicker/dist/react-datepicker.css';
import { readWorkbookFile, toDateInputValue, formatVNDateTime } from '../lib/utils.js';
import { parseProductionRows, groupProduction } from '../lib/rules.js';
import { useReferenceData } from '../lib/useReferenceData.js';
import { useExcludedCodes } from '../lib/useExcludedCodes.js';
import { useLanguage } from '../lib/i18n.jsx';

const DATE_LOCALES = { vi, zh: zhCN };

export default function NcCalculator() {
  const { lang, t } = useLanguage();
  const ref = useReferenceData();
  const excluded = useExcludedCodes();
  const [excludeInput, setExcludeInput] = useState('');

  const [refLoading, setRefLoading] = useState(false);
  const [productionLoading, setProductionLoading] = useState(false);
  const [isCalculating, setIsCalculating] = useState(false);

  const [fileLabel, setFileLabel] = useState(t.fileNotChosen);
  const [records, setRecords] = useState([]);
  const hasData = records.length > 0;

  const [startDateTime, setStartDateTime] = useState(null);
  const [endDateTime, setEndDateTime] = useState(null);

  const [status, setStatus] = useState({ msg: '', isError: false });
  const [summary, setSummary] = useState([]);
  const [missingList, setMissingList] = useState([]);
  const [hasCalculated, setHasCalculated] = useState(false);

  const refInputRef = useRef(null);
  const fileInputRef = useRef(null);

  const grandProduction = summary.reduce((s, r) => s + r.production, 0);
  const grandBlended = summary.reduce((s, r) => s + r.blended, 0);
  const grandNet = summary.reduce((s, r) => s + r.net, 0);
  const canCalculate = hasData && startDateTime && endDateTime;

  function setMsg(msg, isError = false) {
    setStatus({ msg: msg || '', isError: !!isError });
  }

  async function onRefFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setRefLoading(true);
    setMsg('');
    try {
      await ref.loadFromFile(file);
    } catch (err) {
      setMsg(err.message, true);
    } finally {
      setRefLoading(false);
      if (refInputRef.current) refInputRef.current.value = '';
    }
  }

  async function onProductionFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    setMsg('');
    setProductionLoading(true);
    setSummary([]);
    setMissingList([]);
    setHasCalculated(false);

    try {
      const workbook = await readWorkbookFile(file);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
      const parsed = parseProductionRows(rows);

      setRecords(parsed);
      setFileLabel(t.fileValidRows(file.name, parsed.length));

      let min = parsed[0].startDate, max = parsed[0].startDate;
      parsed.forEach((r) => {
        if (r.startDate < min) min = r.startDate;
        if (r.startDate > max) max = r.startDate;
      });
      setStartDateTime(min);
      setEndDateTime(max);
    } catch (err) {
      setFileLabel(t.fileNotChosen);
      setRecords([]);
      setMsg(err.message, true);
    } finally {
      setProductionLoading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  }

  function onCalculate() {
    if (!records.length) { setMsg(t.msgNoProdFile, true); return; }
    if (!startDateTime || !endDateTime) { setMsg(t.msgNoRange, true); return; }
    if (startDateTime > endDateTime) { setMsg(t.msgRangeInvalid, true); return; }

    setIsCalculating(true);
    setTimeout(() => {
      const result = groupProduction(records, startDateTime, endDateTime, ref.ratioMap, ref.codeMap, excluded.codesSet);
      setHasCalculated(true);

      if (!result.rangeRecords.length) {
        setMsg(t.msgNoDataInRange(formatVNDateTime(startDateTime), formatVNDateTime(endDateTime)), true);
        setSummary([]);
        setMissingList([]);
      } else {
        setSummary(result.summary);
        setMissingList(result.missingList);
        const gNet = result.summary.reduce((s, r) => s + r.net, 0);
        const gCount = result.summary.reduce((s, r) => s + r.count, 0);
        let msg = t.msgCalcDone(result.summary.length, gCount, gNet.toFixed(2));
        if (result.missingList.length) msg += t.msgMissingAppend(result.missingList.length);
        setMsg(msg, result.missingList.length > 0);
      }
      setIsCalculating(false);
    }, 30);
  }

  function onExport() {
    const rangeLabel = `${formatVNDateTime(startDateTime)} - ${formatVNDateTime(endDateTime)}`;

    const wsData = [
      [t.xlsxTitle(rangeLabel)],
      [],
      ['NC编码', '胶种', '产量', '掺用量', '净产出']
    ];
    summary.forEach((r) => wsData.push([
      r.ncCode, r.baseCode,
      Number(r.production.toFixed(2)), Number(r.blended.toFixed(2)), Number(r.net.toFixed(2))
    ]));
    wsData.push(['', t.totalRow, Number(grandProduction.toFixed(2)), Number(grandBlended.toFixed(2)), Number(grandNet.toFixed(2))]);

    if (missingList.length) {
      wsData.push([]);
      wsData.push([t.xlsxMissingTitle]);
      wsData.push([t.colFullCode, t.colBatchCount, t.colUncountedWeight]);
      missingList.forEach((r) => wsData.push([r.code, r.count, Number(r.total.toFixed(2))]));
    }

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [{ wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Tong hop NC');
    XLSX.writeFile(wb, `tong_hop_NC_${toDateInputValue(startDateTime)}_${toDateInputValue(endDateTime)}.xlsx`);
  }

  function onCopy() {
    let text = `NC编码\t胶种\t产量\t掺用量\t净产出\n`;
    summary.forEach((r) => {
      text += `${r.ncCode}\t${r.baseCode}\t${r.production.toFixed(2)}\t${r.blended.toFixed(2)}\t${r.net.toFixed(2)}\n`;
    });
    text += `\t${t.totalRow}\t${grandProduction.toFixed(2)}\t${grandBlended.toFixed(2)}\t${grandNet.toFixed(2)}\n`;

    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(text).then(
        () => setMsg(t.msgCopied),
        () => fallbackCopy(text)
      );
    } else {
      fallbackCopy(text);
    }
  }

  function fallbackCopy(text) {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    try {
      document.execCommand('copy');
      setMsg(t.msgCopied);
    } catch {
      setMsg(t.msgCopyFailed, true);
    }
    document.body.removeChild(ta);
  }

  return (
    <div>
      <header className="mb-7">
        <h1 className="text-[22px] font-bold tracking-tight">{t.pageTitle}</h1>
        <p className="text-ink-soft mt-1.5 max-w-[70ch]">{t.pageSubtitle}</p>
      </header>

      {/* Bang tra cuu */}
      <section className="bg-surface border border-line rounded-xl2 shadow-card p-5 mb-5">
        <div className="flex flex-wrap items-center gap-4">
          <label className="w-56 shrink-0 font-semibold text-[13.5px]">{t.refLabel}</label>
          <FileButton inputRef={refInputRef} onChange={onRefFile} disabled={refLoading} label={t.chooseFile} />
          <div className="flex items-center gap-2 text-[13px] font-mono text-ink-soft">
            {refLoading ? (
              <span className="flex items-center gap-1.5 text-ink-soft"><Loader2 size={14} className="animate-spin" /> {t.readingFile}</span>
            ) : (
              <>
                <span>{ref.source === 'uploaded' ? ref.fileName : t.refDefaultText(Object.keys(ref.codeMap).length, Object.keys(ref.ratioMap).length)}</span>
                <Tag isUploaded={ref.source === 'uploaded'} labelDefault={t.tagDefault} labelUploaded={t.tagUpdated} />
              </>
            )}
          </div>
          {ref.source === 'uploaded' && !refLoading && (
            <button
              onClick={ref.resetToDefault}
              className="flex items-center gap-1.5 text-[12.5px] font-semibold text-ink-soft border border-line rounded-md px-2.5 py-1.5 hover:border-accent hover:text-accent-dark transition-colors"
            >
              <RotateCcw size={13} /> {t.resetDefaultRef}
            </button>
          )}
        </div>
        <p className="text-[12.5px] text-ink-faint mt-3 leading-relaxed">
          {t.refHelpPre}{' '}
          <code className="font-mono bg-canvas px-1.5 py-0.5 rounded">产量NC.xlsx</code>{' '}
          {t.refHelpPost}
        </p>
      </section>

      {/* Ma loai tru khoi tinh toan */}
      <section className="bg-surface border border-line rounded-xl2 shadow-card p-5 mb-5">
        <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
          <label className="font-semibold text-[13.5px]">
            {t.excludeLabel} <span className="text-ink-faint font-normal">{t.excludeSub}</span>
          </label>
          <div className="flex items-center gap-2">
            <Tag isUploaded={excluded.source === 'edited'} labelDefault={t.tagDefault} labelUploaded={t.tagEdited} />
            {excluded.source === 'edited' && (
              <button
                onClick={excluded.resetToDefault}
                className="flex items-center gap-1.5 text-[12.5px] font-semibold text-ink-soft border border-line rounded-md px-2.5 py-1.5 hover:border-accent hover:text-accent-dark transition-colors"
              >
                <RotateCcw size={13} /> {t.resetDefaultExclude}
              </button>
            )}
          </div>
        </div>

        <div className="flex items-center gap-2 mb-3">
          <input
            type="text"
            value={excludeInput}
            onChange={(e) => setExcludeInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') { excluded.addCode(excludeInput); setExcludeInput(''); }
            }}
            placeholder={t.excludePlaceholder}
            className="font-mono text-[13px] border border-line rounded-lg px-3 py-2 w-40"
          />
          <button
            onClick={() => { excluded.addCode(excludeInput); setExcludeInput(''); }}
            className="flex items-center gap-1.5 text-[13px] font-semibold border border-line rounded-lg px-3 py-2 hover:border-accent hover:text-accent-dark transition-colors"
          >
            <Plus size={14} /> {t.addBtn}
          </button>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {excluded.codes.length === 0 && (
            <span className="text-[12.5px] text-ink-faint">{t.excludeNone}</span>
          )}
          {excluded.codes.map((code) => (
            <span key={code} className="flex items-center gap-1.5 bg-canvas border border-line rounded-full pl-3 pr-1.5 py-1 text-[12.5px] font-mono">
              {code}
              <button
                onClick={() => excluded.removeCode(code)}
                className="text-ink-faint hover:text-red-600 rounded-full p-0.5"
                aria-label={t.removeAria(code)}
              >
                <X size={12} />
              </button>
            </span>
          ))}
        </div>

        <p className="text-[12.5px] text-ink-faint mt-3 leading-relaxed">
          {t.excludeHelpPre} <code className="font-mono bg-canvas px-1.5 py-0.5 rounded">F001</code>,{' '}
          <code className="font-mono bg-canvas px-1.5 py-0.5 rounded">AQPS33</code>{t.excludeHelpMid}
          <strong> {t.excludeHelpContains}</strong> {t.excludeHelpPost}
        </p>
      </section>

      {/* File san xuat + khoang thoi gian */}
      <section className="bg-surface border border-line rounded-xl2 shadow-card p-5 mb-5 divide-y divide-line">
        <div className="flex flex-wrap items-center gap-4 pb-4">
          <label className="w-56 shrink-0 font-semibold text-[13.5px]">{t.step1Label}</label>
          <FileButton inputRef={fileInputRef} onChange={onProductionFile} disabled={productionLoading} label={t.chooseFile} />
          <span className="flex items-center gap-1.5 text-[13px] font-mono text-ink-soft">
            {productionLoading && <Loader2 size={14} className="animate-spin" />}
            {productionLoading ? t.readingFile : fileLabel}
          </span>
        </div>

        <div className="flex flex-wrap items-center gap-4 py-4">
          <label className="w-56 shrink-0 font-semibold text-[13.5px]">{t.step2Label}</label>
          <div className="flex flex-wrap items-center gap-5">
            <RangeDateTimePicker label={t.rangeFrom} value={startDateTime} onChange={setStartDateTime} disabled={!hasData} locale={DATE_LOCALES[lang]} />
            <RangeDateTimePicker label={t.rangeTo} value={endDateTime} onChange={setEndDateTime} disabled={!hasData} locale={DATE_LOCALES[lang]} />
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-4 pt-4">
          <button
            onClick={onCalculate}
            disabled={!canCalculate || isCalculating}
            className="flex items-center gap-2 bg-accent text-[#241605] font-semibold text-[13.5px] px-5 py-2.5 rounded-lg disabled:opacity-40 disabled:cursor-not-allowed hover:bg-accent-dark transition-colors"
          >
            {isCalculating && <Loader2 size={14} className="animate-spin" />}
            {isCalculating ? t.calculating : t.calcBtn}
          </button>
          {status.msg && (
            <span className={`flex items-center gap-1.5 text-[13px] ${status.isError ? 'text-red-600' : 'text-ink-soft'}`}>
              {status.isError ? <XCircle size={14} className="shrink-0" /> : <CheckCircle2 size={14} className="shrink-0 text-emerald-600" />}
              {status.msg}
            </span>
          )}
        </div>
      </section>

      {/* Ket qua */}
      <section className="bg-surface border border-line rounded-xl2 shadow-card p-5 mb-5">
        <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
          <h2 className="font-bold text-[15px]">{t.resultTitle}</h2>
          <div className="flex gap-2">
            <ActionButton icon={Download} label={t.exportBtn} onClick={onExport} disabled={!summary.length} />
            <ActionButton icon={Copy} label={t.copyBtn} onClick={onCopy} disabled={!summary.length} />
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-[14px]">
            <thead>
              <tr className="text-[12.5px] text-ink-soft font-semibold">
                <th className="text-left py-2 px-2.5 border-b border-line">NC编码</th>
                <th className="text-left py-2 px-2.5 border-b border-line">胶种</th>
                <th className="text-right py-2 px-2.5 border-b border-line">产量</th>
                <th className="text-right py-2 px-2.5 border-b border-line">掺用量</th>
                <th className="text-right py-2 px-2.5 border-b border-line">净产出</th>
              </tr>
            </thead>
            <tbody>
              {!summary.length && (
                <tr>
                  <td colSpan={5} className="text-center text-ink-faint py-8">
                    {hasCalculated ? t.emptyResultCalculated : t.emptyResultInitial}
                  </td>
                </tr>
              )}
              {summary.map((row) => (
                <tr key={row.baseCode} className="hover:bg-canvas/60">
                  <td className="py-2.5 px-2.5 border-b border-line font-mono text-ink-soft">{row.ncCode || '—'}</td>
                  <td className="py-2.5 px-2.5 border-b border-line">{row.baseCode}</td>
                  <td className="py-2.5 px-2.5 border-b border-line text-right font-mono">{row.production.toFixed(2)}</td>
                  <td className="py-2.5 px-2.5 border-b border-line text-right font-mono">{row.blended.toFixed(2)}</td>
                  <td className="py-2.5 px-2.5 border-b border-line text-right font-mono font-semibold">{row.net.toFixed(2)}</td>
                </tr>
              ))}
            </tbody>
            {summary.length > 0 && (
              <tfoot>
                <tr className="font-bold font-mono bg-amber-100/70">
                  <td colSpan={2} className="py-3 px-2.5 border-t-2 border-ink font-sans">{t.totalRow}</td>
                  <td className="py-3 px-2.5 border-t-2 border-ink text-right">{grandProduction.toFixed(2)}</td>
                  <td className="py-3 px-2.5 border-t-2 border-ink text-right">{grandBlended.toFixed(2)}</td>
                  <td className="py-3 px-2.5 border-t-2 border-ink text-right">{grandNet.toFixed(2)}</td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      </section>

      {/* Canh bao */}
      {missingList.length > 0 && (
        <section className="bg-amber-50 border border-amber-300 rounded-xl2 p-5">
          <h2 className="font-bold text-[14px] flex items-center gap-2 text-amber-900">
            <AlertTriangle size={16} /> {t.warningTitle}
          </h2>
          <p className="text-[12.5px] text-amber-800/80 mt-2 mb-3 leading-relaxed">
            {t.warningDescPre}{' '}
            <code className="font-mono bg-amber-100 px-1.5 py-0.5 rounded">产量NC.xlsx</code>{' '}
            {t.warningDescPost}
          </p>
          <div className="overflow-x-auto">
            <table className="w-full text-[14px]">
              <thead>
                <tr className="text-[12.5px] text-amber-800 font-semibold">
                  <th className="text-left py-2 px-2.5 border-b border-amber-200">{t.colFullCode}</th>
                  <th className="text-right py-2 px-2.5 border-b border-amber-200">{t.colBatchCount}</th>
                  <th className="text-right py-2 px-2.5 border-b border-amber-200">{t.colUncountedWeight}</th>
                </tr>
              </thead>
              <tbody>
                {missingList.map((row) => (
                  <tr key={row.code}>
                    <td className="py-2 px-2.5 border-b border-amber-100 font-mono">{row.code}</td>
                    <td className="py-2 px-2.5 border-b border-amber-100 text-right font-mono">{row.count}</td>
                    <td className="py-2 px-2.5 border-b border-amber-100 text-right font-mono">{row.total.toFixed(2)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}
    </div>
  );
}

function FileButton({ inputRef, onChange, disabled, label }) {
  return (
    <label className={`flex items-center gap-2 text-[13px] font-semibold border border-line rounded-lg px-3.5 py-2 cursor-pointer hover:border-accent hover:text-accent-dark transition-colors ${disabled ? 'opacity-40 pointer-events-none' : ''}`}>
      <UploadCloud size={14} />
      {label}
      <input ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={onChange} disabled={disabled} />
    </label>
  );
}

function Tag({ isUploaded, labelDefault, labelUploaded }) {
  return (
    <span className={`text-[11px] font-mono px-2 py-0.5 rounded-full ${isUploaded ? 'bg-emerald-100 text-emerald-700' : 'bg-canvas text-ink-soft'}`}>
      {isUploaded ? labelUploaded : labelDefault}
    </span>
  );
}

function RangeDateTimePicker({ label, value, onChange, disabled, locale }) {
  const hour = value ? value.getHours() : 0;
  const minute = value ? value.getMinutes() : 0;

  function updateDate(newDate) {
    if (!newDate) return;
    const combined = new Date(newDate);
    combined.setHours(hour, minute, 0, 0);
    onChange(combined);
  }

  function updateHour(h) {
    const base = value ? new Date(value) : new Date();
    base.setHours(h);
    onChange(base);
  }

  function updateMinute(m) {
    const base = value ? new Date(value) : new Date();
    base.setMinutes(m);
    onChange(base);
  }

  return (
    <div className={`flex items-center rounded-xl border border-line bg-surface overflow-hidden ${disabled ? 'opacity-50' : 'focus-within:ring-2 focus-within:ring-accent/30 focus-within:border-accent'}`}>
      <span className="text-[12px] font-semibold text-accent-dark bg-accent-soft px-3 py-2.5 shrink-0">
        {label}
      </span>
      <DatePicker
        selected={value}
        onChange={updateDate}
        disabled={disabled}
        locale={locale}
        dateFormat="dd/MM/yyyy"
        className="font-mono text-[13.5px] pl-2.5 pr-1.5 py-2 bg-transparent outline-none w-[100px] disabled:text-ink-faint"
      />
      <span className="w-px self-stretch bg-line" />
      <select
        value={hour}
        onChange={(e) => updateHour(Number(e.target.value))}
        disabled={disabled}
        className="font-mono text-[13.5px] pl-2 pr-1 py-2 bg-transparent outline-none disabled:text-ink-faint"
      >
        {Array.from({ length: 24 }, (_, h) => (
          <option key={h} value={h}>{String(h).padStart(2, '0')}</option>
        ))}
      </select>
      <span className="text-ink-faint font-mono">:</span>
      <select
        value={minute}
        onChange={(e) => updateMinute(Number(e.target.value))}
        disabled={disabled}
        className="font-mono text-[13.5px] pl-1 pr-2 py-2 bg-transparent outline-none disabled:text-ink-faint"
      >
        {Array.from({ length: 60 }, (_, m) => (
          <option key={m} value={m}>{String(m).padStart(2, '0')}</option>
        ))}
      </select>
      <Clock size={13} className="text-ink-faint mr-2.5 shrink-0" />
    </div>
  );
}

function ActionButton({ icon: Icon, label, onClick, disabled }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className="flex items-center gap-1.5 text-[13px] font-semibold border border-line rounded-lg px-3.5 py-2 disabled:opacity-40 disabled:cursor-not-allowed hover:border-accent hover:text-accent-dark transition-colors"
    >
      <Icon size={14} /> {label}
    </button>
  );
}
